import type { AgentWorkerEnv, WorkbookMutationExecutionPlan, WorkbookMutationPayload, WorkbookMutationResult } from '#agent/types'
import { requireMutation } from '#agent/workbook/mutation/errors'
import { settleTurnCredits } from '#agent/billing/billingRepository'
import { calculateActualTurnCredits, buildCreditUsageActions } from '#shared/billing/planCatalog'
import { parseMutationPayload, parseMutationResult } from '#agent/persistence/taskProtocol'

export type MutationRecord = { task_id: string; payload_json: string; result_json: string | null; state: string }

export async function stableMutationId(userId: string, threadId: string, turnId: string) {
  const bytes = new Uint8Array(await crypto.subtle.digest('SHA-256', new TextEncoder().encode(JSON.stringify(['mutation-v1', userId, threadId, turnId]))))
  bytes[6] = (bytes[6]! & 0x0f) | 0x50
  bytes[8] = (bytes[8]! & 0x3f) | 0x80
  const hex = [...bytes.slice(0, 16)].map(value => value.toString(16).padStart(2, '0')).join('')
  return `${hex.slice(0, 8)}-${hex.slice(8, 12)}-${hex.slice(12, 16)}-${hex.slice(16, 20)}-${hex.slice(20)}`
}

export async function createMutationTask(env: AgentWorkerEnv, payload: WorkbookMutationPayload, plan: WorkbookMutationExecutionPlan, prompt: string) {
  payload = parseMutationPayload(payload)
  const now = new Date().toISOString()
  const statements = [
    env.DB.prepare(`INSERT INTO excel_agent_tasks (id, user_id, session_id, thread_id, task_type, workflow_instance_id, input_r2_key, input_name, prompt, plan_json, status, created_at, updated_at)
      SELECT ?, ?, ?, ?, 'workbook_mutation', ?, ?, ?, ?, ?, 'queued', ?, ?
      WHERE EXISTS (SELECT 1 FROM excel_agent_conversations c WHERE id = ? AND user_id = ? AND session_id = ?
        AND NOT EXISTS (SELECT 1 FROM excel_agent_conversation_tombstones d WHERE d.conversation_id=c.id))
      AND NOT EXISTS (SELECT 1 FROM excel_agent_tasks WHERE thread_id=? AND task_type='workbook_mutation' AND status IN ('queued','processing') AND id<>?)
      ON CONFLICT(id) DO NOTHING`).bind(payload.taskId, payload.userId, payload.sessionId, payload.threadId, `edit-${payload.taskId}`, payload.inputR2Key, payload.inputName, prompt, JSON.stringify(plan), now, now, payload.threadId, payload.userId, payload.sessionId, payload.threadId, payload.taskId),
    env.DB.prepare(`INSERT INTO excel_agent_mutations (task_id, conversation_id, turn_id, payload_json, state, created_at, updated_at)
      SELECT ?, ?, ?, ?, 'queued', ?, ? WHERE EXISTS (SELECT 1 FROM excel_agent_tasks WHERE id = ? AND user_id = ?)
      ON CONFLICT(task_id) DO NOTHING`).bind(payload.taskId, payload.threadId, payload.turnId, JSON.stringify(payload), now, now, payload.taskId, payload.userId),
    env.DB.prepare(`UPDATE excel_agent_messages SET task_id = ?, updated_at = ? WHERE conversation_id = ? AND turn_id = ? AND role = 'assistant'
      AND EXISTS (SELECT 1 FROM excel_agent_mutations WHERE task_id = ?)`).bind(payload.taskId, now, payload.threadId, payload.turnId, payload.taskId)
  ]
  if (payload.creditReservationId) statements.push(env.DB.prepare(`UPDATE credit_reservations SET task_id = ?, updated_at = ? WHERE id = ? AND user_id = ? AND status = 'reserved'
    AND EXISTS (SELECT 1 FROM excel_agent_mutations WHERE task_id = ?)`)
    .bind(payload.taskId, now, payload.creditReservationId, payload.userId, payload.taskId))
  await env.DB.batch(statements)
  const saved = await getMutationRecord(env, payload.taskId)
  requireMutation(saved, 'CONVERSATION_BUSY', 'This conversation is unavailable or already has a workbook edit running. Wait for it to finish before trying again.')
  return parseMutationPayload(JSON.parse(saved.payload_json))
}

export function getMutationRecord(env: AgentWorkerEnv, taskId: string) {
  return env.DB.prepare('SELECT task_id, payload_json, result_json, state FROM excel_agent_mutations WHERE task_id = ?').bind(taskId).first<MutationRecord>()
}

export async function assertMutationActive(env: AgentWorkerEnv, taskId: string) {
  const row = await env.DB.prepare(`SELECT t.status FROM excel_agent_tasks t JOIN excel_agent_mutations m ON m.task_id = t.id
    JOIN excel_agent_conversations c ON c.id = t.thread_id AND c.user_id = t.user_id AND c.session_id = t.session_id
    WHERE t.id = ? AND m.state IN ('queued', 'staged') AND t.status IN ('queued', 'processing')
      AND NOT EXISTS (SELECT 1 FROM excel_agent_conversation_tombstones d WHERE d.conversation_id=c.id)`).bind(taskId).first()
  requireMutation(row, 'TASK_CANCELLED', 'The workbook edit is no longer active.')
}

export async function stageMutationResult(env: AgentWorkerEnv, taskId: string, result: WorkbookMutationResult) {
  result = parseMutationResult(result)
  requireMutation(result.taskId === taskId, 'INVALID_RESULT', 'Workbook result does not belong to this task.')
  await assertMutationActive(env, taskId)
  const update = await env.DB.prepare(`UPDATE excel_agent_mutations SET result_json = ?, state = 'staged', updated_at = ?
    WHERE task_id = ? AND state IN ('queued', 'staged')`).bind(JSON.stringify(result), new Date().toISOString(), taskId).run()
  requireMutation(update.meta.changes, 'TASK_CANCELLED', 'The workbook edit was cancelled before publication.')
}

function publicationStatements(env: AgentWorkerEnv, payload: WorkbookMutationPayload, result: WorkbookMutationResult, finalizationKey?: string) {
  const now = new Date().toISOString()
  const billingGuard = finalizationKey ? 'AND EXISTS (SELECT 1 FROM credit_reservations WHERE id = ? AND finalization_key = ? AND status = \'settled\')' : ''
  const billingValues = finalizationKey ? [payload.creditReservationId!, finalizationKey] : []
  const active = `EXISTS (SELECT 1 FROM excel_agent_mutations m JOIN excel_agent_tasks t ON m.task_id = t.id
    JOIN excel_agent_conversations c ON c.id = t.thread_id AND c.user_id = t.user_id AND c.session_id = t.session_id
    JOIN excel_agent_workbooks w ON w.id = json_extract(m.payload_json, '$.plan.sourceWorkbookId')
      AND w.user_id = t.user_id AND w.session_id = t.session_id AND w.r2_key = t.input_r2_key
    WHERE m.task_id = ? AND m.state = 'staged' AND t.status = 'processing'
      AND NOT EXISTS (SELECT 1 FROM excel_agent_conversation_tombstones d WHERE d.conversation_id=c.id)) ${billingGuard}`
  const statements: D1PreparedStatement[] = []
  if (result.workbookId !== payload.plan.sourceWorkbookId) statements.push(env.DB.prepare(`
    INSERT INTO excel_agent_workbooks (id, user_id, session_id, r2_key, file_name, size_bytes, summary_json, context_r2_key, source_etag, context_version, parsed_at, parent_workbook_id, root_workbook_id, source_task_id, created_at, updated_at)
    SELECT ?, user_id, session_id, ?, ?, ?, ?, ?, ?, ?, ?, id, COALESCE(root_workbook_id, id), ?, ?, ? FROM excel_agent_workbooks
    WHERE id = ? AND user_id = ? AND session_id = ? AND ${active}
    ON CONFLICT(id) DO UPDATE SET r2_key = CASE
      WHEN excel_agent_workbooks.user_id = excluded.user_id AND excel_agent_workbooks.session_id = excluded.session_id
        AND excel_agent_workbooks.source_task_id = excluded.source_task_id AND excel_agent_workbooks.r2_key = excluded.r2_key
        AND excel_agent_workbooks.parent_workbook_id = excluded.parent_workbook_id THEN excel_agent_workbooks.r2_key
      ELSE NULL END
  `).bind(result.workbookId, result.outputR2Key, result.outputName, result.sizeBytes, JSON.stringify(result.summary), result.contextR2Key, result.sourceEtag, result.contextVersion, now, payload.taskId, now, now, payload.plan.sourceWorkbookId, payload.userId, payload.sessionId, payload.taskId, ...billingValues))
  if (result.workbookId !== payload.plan.sourceWorkbookId) statements.push(env.DB.prepare(`UPDATE excel_agent_conversations SET workbook_id = ?, workbook_revision = workbook_revision + 1, updated_at = ?
    WHERE id = ? AND user_id = ? AND session_id = ? AND workbook_id = ? AND workbook_revision = ?
      AND EXISTS (SELECT 1 FROM excel_agent_workbooks WHERE id = ? AND user_id = ?)
      AND ${active}`).bind(result.workbookId, now, payload.threadId, payload.userId, payload.sessionId, payload.plan.sourceWorkbookId, payload.sourceRevision, result.workbookId, payload.userId, payload.taskId, ...billingValues))
  statements.push(env.DB.prepare(`UPDATE excel_agent_tasks SET status = 'completed', output_r2_key = ?, output_name = ?, result_json = ?, error_message = NULL, completed_at = ?, updated_at = ?
    WHERE id = ? AND ${active}`).bind(result.outputR2Key, result.outputName, JSON.stringify(result), now, now, payload.taskId, payload.taskId, ...billingValues))
  statements.push(env.DB.prepare(`UPDATE excel_agent_mutations SET state = 'completed', updated_at = ? WHERE task_id = ? AND state = 'staged'
    AND EXISTS (SELECT 1 FROM excel_agent_tasks WHERE id = ? AND status = 'completed') ${billingGuard}`).bind(now, payload.taskId, payload.taskId, ...billingValues))
  return statements
}

export async function publishMutationResult(env: AgentWorkerEnv, payload: WorkbookMutationPayload, result: WorkbookMutationResult) {
  payload = parseMutationPayload(payload)
  result = parseMutationResult(result)
  requireMutation(result.taskId === payload.taskId && result.sourceWorkbookId === payload.plan.sourceWorkbookId && [payload.resultWorkbookId, payload.plan.sourceWorkbookId].includes(result.workbookId), 'INVALID_RESULT', 'Workbook result does not match its source task and version.')
  if (result.workbookId !== payload.plan.sourceWorkbookId) {
    const conflict = await env.DB.prepare(`SELECT 1 FROM excel_agent_workbooks WHERE id=?
      AND NOT (user_id=? AND session_id=? AND r2_key=? AND COALESCE(source_task_id,'')=? AND COALESCE(parent_workbook_id,'')=?)`)
      .bind(result.workbookId, payload.userId, payload.sessionId, result.outputR2Key, payload.taskId, payload.plan.sourceWorkbookId).first()
    requireMutation(!conflict, 'RESULT_VERSION_CONFLICT', 'This output version conflicts with an existing workbook. Start a new task; the existing file is unchanged.')
  }
  const hasChanges = result.changes.changedCells > 0 || result.changes.addedSheets.length > 0
  const tools = [...payload.creditToolNames.filter(name => name !== 'editWorkbook'), ...(hasChanges ? ['editWorkbook'] : [])]
  if (payload.creditReservationId) {
    await settleTurnCredits(env, {
      reservationId: payload.creditReservationId,
      taskId: payload.taskId,
      actualCredits: calculateActualTurnCredits(tools, hasChanges),
      actions: buildCreditUsageActions(payload.creditTurnId, tools, hasChanges),
      atomicCommit: { taskId: payload.taskId, statements: key => publicationStatements(env, payload, result, key) }
    })
  } else await env.DB.batch(publicationStatements(env, payload, result))
  const record = await getMutationRecord(env, payload.taskId)
  requireMutation(record?.state === 'completed', 'PUBLICATION_INCOMPLETE', 'Workbook publication did not complete. The task will be reconciled.')
}
