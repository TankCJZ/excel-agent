import type { AgentWorkerEnv, WorkbookMutationPayload, WorkbookMutationResult } from '#agent/types'
import { publishMutationResult } from '#agent/persistence/workbookMutations'
import { dispatchWorkbookMutation } from '#agent/workflows/mutationDispatch'
import { failTask } from '#agent/persistence/agentRepository'
import { releaseTurnCredits, settleTurnCredits } from '#agent/billing/billingRepository'
import { calculateActualTurnCredits, buildCreditUsageActions } from '#shared/billing/planCatalog'
import { parseMutationPayload, parseMutationResult } from '#agent/persistence/taskProtocol'
import { WorkbookMutationError } from '#agent/workbook/mutation/errors'
import { discardMutationArtifacts } from '#agent/workflows/mutationCleanup'

export async function reconcileWorkbookMutations(env: AgentWorkerEnv, taskId?: string) {
  const records = await env.DB.prepare(`SELECT m.task_id, m.payload_json, m.result_json, m.state, m.updated_at FROM excel_agent_mutations m
    JOIN excel_agent_tasks t ON t.id = m.task_id WHERE m.state IN ('queued','staged','cancelled','failing')
    AND (t.status IN ('queued','processing') OR (t.status='failed' AND m.state IN ('failing','cancelled')))
    AND (? IS NULL OR m.task_id = ?) AND m.updated_at < ?
    ORDER BY m.updated_at LIMIT 20`).bind(taskId || null, taskId || null, new Date(Date.now() - 60_000).toISOString())
    .all<{ task_id: string; payload_json: string; result_json: string | null; state: string; updated_at: string }>()
  for (const row of records.results) {
    try {
      // Polls and cron may overlap. Claim a bounded recovery lease before I/O.
      const claim = await env.DB.prepare('UPDATE excel_agent_mutations SET updated_at=? WHERE task_id=? AND state=? AND updated_at=?')
        .bind(new Date().toISOString(), row.task_id, row.state, row.updated_at).run()
      if (!claim.meta.changes) continue
      let payload: WorkbookMutationPayload
      try {
        payload = parseMutationPayload(JSON.parse(row.payload_json))
        if (payload.taskId !== row.task_id) throw new Error('Mutation task identity mismatch')
      } catch {
        await quarantineUnreadableMutation(env, row.task_id)
        continue
      }
      const deleting = await env.DB.prepare('SELECT 1 FROM excel_agent_conversation_tombstones WHERE conversation_id=?').bind(payload.threadId).first()
      if (deleting && row.state !== 'cancelled' && row.state !== 'failing') {
        await env.DB.prepare("UPDATE excel_agent_mutations SET state='cancelled' WHERE task_id=? AND state IN ('queued','staged')").bind(row.task_id).run()
        await finalizeUnpublishedMutation(env, payload, true)
        continue
      }
      if (row.state === 'cancelled' || row.state === 'failing') {
        await finalizeUnpublishedMutation(env, payload, row.state === 'cancelled')
        continue
      }
      const sourceExists = await env.DB.prepare(`SELECT 1 FROM excel_agent_workbooks w
        JOIN excel_agent_conversations c ON c.id=? AND c.user_id=w.user_id AND c.session_id=w.session_id
        WHERE w.id=? AND w.user_id=? AND w.session_id=? AND w.r2_key=?`)
        .bind(payload.threadId, payload.plan.sourceWorkbookId, payload.userId, payload.sessionId, payload.inputR2Key).first()
      if (!sourceExists) {
        const fenced = await env.DB.prepare("UPDATE excel_agent_mutations SET state='failing' WHERE task_id=? AND state IN ('queued','staged')").bind(row.task_id).run()
        if (fenced.meta.changes) await finalizeUnpublishedMutation(env, payload, false)
        continue
      }
      if (row.state === 'staged' && row.result_json) {
        let result: WorkbookMutationResult
        try {
          result = parseMutationResult(JSON.parse(row.result_json))
          if (result.taskId !== row.task_id || result.sourceWorkbookId !== payload.plan.sourceWorkbookId || ![payload.resultWorkbookId, payload.plan.sourceWorkbookId].includes(result.workbookId)) throw new Error('Mutation result identity mismatch')
        } catch {
          await quarantineUnreadableMutation(env, row.task_id)
          continue
        }
        const artifact = await env.WORKBOOKS.head(result.outputR2Key)
        if (artifact?.etag === result.sourceEtag && await env.WORKBOOKS.head(result.contextR2Key)) {
          await publishMutationResult(env, payload, result)
          continue
        }
      }
      let status: InstanceStatus
      try {
        const instance = await env.EXCEL_AGENT_WORKFLOW.get(`edit-${row.task_id}`)
        status = await instance.status()
      } catch (error) {
        // A status failure is not proof that a workflow is dead. create using
        // the stable id is safe and distinguishes absent instances from outages.
        if (row.state === 'queued') await dispatchWorkbookMutation(env, payload)
        else throw error
        continue
      }
      if (status.status === 'unknown' && row.state === 'queued') { await dispatchWorkbookMutation(env, payload); continue }
      if (!['errored', 'terminated', 'complete'].includes(status.status)) continue
      // Fence publication before settlement. A partial finalization remains
      // recoverable even when settlement or task persistence itself fails.
      const fenced = await env.DB.prepare("UPDATE excel_agent_mutations SET state='failing' WHERE task_id=? AND state IN ('queued','staged')").bind(row.task_id).run()
      if (fenced.meta.changes) await finalizeUnpublishedMutation(env, payload, false)
    } catch (error) {
      if (error instanceof WorkbookMutationError && ['INVALID_RESULT', 'RESULT_VERSION_CONFLICT'].includes(error.code)) {
        await quarantineUnreadableMutation(env, row.task_id)
        continue
      }
      console.error('Workbook mutation reconciliation will retry', { taskId: row.task_id, message: error instanceof Error ? error.message : 'Unknown failure' })
    }
  }
}

async function quarantineUnreadableMutation(env: AgentWorkerEnv, taskId: string) {
  // Do not trust reservation IDs in a broken/unknown payload. Resolve held
  // credits from the authenticated task row, and preserve the JSON for audit.
  const fenced = await env.DB.prepare(`UPDATE excel_agent_mutations SET state='failing' WHERE task_id=? AND state IN ('queued','staged','failing','cancelled')
    AND EXISTS (SELECT 1 FROM excel_agent_tasks WHERE id=? AND status IN ('queued','processing','failed'))`).bind(taskId, taskId).run()
  if (!fenced.meta.changes) return
  await discardMutationArtifacts(env, taskId, true)
  const holds = await env.DB.prepare(`SELECT r.id FROM credit_reservations r JOIN excel_agent_tasks t ON t.id=r.task_id AND t.user_id=r.user_id
    WHERE t.id=? AND r.status='reserved'`).bind(taskId).all<{ id: string }>()
  for (const hold of holds.results) await releaseTurnCredits(env, hold.id, 'Unsupported or corrupted workbook edit protocol')
  await failTask(env, taskId, 'The saved workbook edit uses an unsupported or corrupted format. The source file is unchanged. Start a new task.')
  await env.DB.prepare("UPDATE excel_agent_mutations SET state='failed',updated_at=? WHERE task_id=? AND state='failing'").bind(new Date().toISOString(), taskId).run()
}

async function finalizeUnpublishedMutation(env: AgentWorkerEnv, payload: WorkbookMutationPayload, cancelled: boolean) {
  await discardMutationArtifacts(env, payload.taskId, true)
  const tools = payload.creditToolNames.filter(name => name !== 'editWorkbook')
  if (payload.creditReservationId) {
    if (cancelled) await releaseTurnCredits(env, payload.creditReservationId, 'Workbook edit cancelled before publication')
    else await settleTurnCredits(env, {
      reservationId: payload.creditReservationId, taskId: payload.taskId,
      actualCredits: calculateActualTurnCredits(tools), actions: buildCreditUsageActions(payload.creditTurnId, tools)
    })
  }
  await failTask(env, payload.taskId, cancelled
    ? 'The workbook edit was cancelled. The original file is unchanged.'
    : 'The edit could not publish a verified workbook. Check that the source file is still available before retrying.', cancelled ? 'TASK_CANCELLED' : 'WORKBOOK_EDIT_FAILED')
  await env.DB.prepare("UPDATE excel_agent_mutations SET state='failed', updated_at=? WHERE task_id=? AND state IN ('failing','cancelled')")
    .bind(new Date().toISOString(), payload.taskId).run()
}
