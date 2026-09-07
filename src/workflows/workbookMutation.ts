import type { WorkflowStep } from 'cloudflare:workers'
import { NonRetryableError } from 'cloudflare:workflows'
import type { AgentWorkerEnv, WorkbookMutationPayload, WorkbookMutationResult } from '#agent/types'
import { getWorkbookForSession, markTaskProcessing, appendTaskEvent, failTask } from '#agent/persistence/agentRepository'
import { assertMutationActive, getMutationRecord, stageMutationResult, publishMutationResult } from '#agent/persistence/workbookMutations'
import { workbookMutationPlanSchema } from '#shared/agent/workbookMutation'
import { applyWorkbookMutation } from '#agent/workbook/mutation/engine'
import { prepareEditableBytes } from '#agent/workbook/mutation/source'
import { WorkbookMutationError, requireMutation } from '#agent/workbook/mutation/errors'
import { parseWorkbookBytes } from '#agent/workbook/service'
import { writeWorkbookContext, WORKBOOK_CONTEXT_VERSION } from '#agent/workbook/context'
import { settleTurnCredits } from '#agent/billing/billingRepository'
import { calculateActualTurnCredits, buildCreditUsageActions } from '#shared/billing/planCatalog'
import { parseMutationPayload, parseMutationResult } from '#agent/persistence/taskProtocol'
import { compactResultSummary, assertWorkflowResultBudget } from '#agent/workbook/resultPreview'
import { discardMutationArtifacts } from '#agent/workflows/mutationCleanup'

export async function runWorkbookMutation(env: AgentWorkerEnv, payload: WorkbookMutationPayload, step: WorkflowStep) {
  payload = parseMutationPayload(payload)
  try {
    await step.do('edit v1 validate and start', async () => {
      await assertMutationActive(env, payload.taskId)
      workbookMutationPlanSchema.parse(payload.plan)
      const workbook = await getWorkbookForSession(env, payload.plan.sourceWorkbookId, payload.userId, payload.sessionId)
      requireMutation(workbook && workbook.r2Key === payload.inputR2Key, 'SOURCE_NOT_FOUND', 'Source workbook was not found for this account.')
      if (workbook.sourceTaskId) {
        const sourceTask = await env.DB.prepare('SELECT thread_id FROM excel_agent_tasks WHERE id = ? AND user_id = ?').bind(workbook.sourceTaskId, payload.userId).first<{ thread_id: string }>()
        requireMutation(sourceTask?.thread_id === payload.threadId, 'SOURCE_NOT_FOUND', 'Edited workbook versions must stay in their original conversation.')
      }
      await markTaskProcessing(env, payload.taskId, 'Validating workbook edits')
      return true
    })
    const result = await step.do('edit v1 transform and stage', { retries: { limit: 2, delay: '2 seconds', backoff: 'exponential' }, timeout: '5 minutes' }, async () => {
      try {
        await assertMutationActive(env, payload.taskId)
        const stored = await getMutationRecord(env, payload.taskId)
        if (stored?.result_json) {
          const result = parseMutationResult(JSON.parse(stored.result_json))
          const object = await env.WORKBOOKS.head(result.outputR2Key)
          if (object?.etag === result.sourceEtag && await env.WORKBOOKS.head(result.contextR2Key)) return result
        }
        const source = await env.WORKBOOKS.get(payload.inputR2Key)
        requireMutation(source && source.etag === payload.plan.sourceEtag, 'SOURCE_CHANGED', 'The original workbook is missing or changed. Re-upload it before retrying.')
        const input = prepareEditableBytes(new Uint8Array(await source.arrayBuffer()), payload.inputName)
        await appendTaskEvent(env, payload.taskId, 'mutation.applying', 'Applying validated workbook changes', { operationCount: payload.plan.operations.length })
        const output = applyWorkbookMutation(input.bytes, { title: payload.plan.title, operations: payload.plan.operations })
        output.changes.convertedFrom = input.convertedFrom
        const hasChanges = output.changes.changedCells > 0 || output.changes.addedSheets.length > 0
        const outputR2Key = `agent-demo/results/${payload.taskId}/edited.xlsx`
        const outputName = `${payload.inputName.replace(/\.(xlsx|xls|csv)$/i, '').replace(/[\\/\x00-\x1f]/g, '-').slice(0, 100)}-edited.xlsx`
        const context = await parseWorkbookBytes(output.bytes, { pendingCalculation: output.changes.pendingCalculation })
        const summary = compactResultSummary(context.summary)
        const contextR2Key = `agent-demo/results/${payload.taskId}/context-v${WORKBOOK_CONTEXT_VERSION}.json.gz`
        const preview = output.changes.sheets[0]
        const result: WorkbookMutationResult = {
          kind: 'workbook_mutation', schemaVersion: 1, taskId: payload.taskId, outputR2Key, outputName,
          summary, analysis: null, preview: preview?.rows.map(row => row.cells.map(cell => cell.formula || cell.value)) || summary.preview,
          previewSheetName: preview?.name || context.summary.primarySheet, previewRowCount: preview?.rows.length || context.summary.rowCount,
          sizeBytes: output.bytes.length, changes: output.changes, sourceWorkbookId: payload.plan.sourceWorkbookId,
          workbookId: hasChanges ? payload.resultWorkbookId : payload.plan.sourceWorkbookId, sourceEtag: '0'.repeat(64), contextR2Key, contextVersion: WORKBOOK_CONTEXT_VERSION
        }
        // Check the serialized step result before creating any R2 output. The
        // placeholder reserves more space than an R2 object ETag requires.
        assertWorkflowResultBudget(result)
        await assertMutationActive(env, payload.taskId)
        const object = await env.WORKBOOKS.put(outputR2Key, output.bytes, {
          httpMetadata: { contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' },
          customMetadata: { taskId: payload.taskId, userId: payload.userId, sourceWorkbookId: payload.plan.sourceWorkbookId, pendingCalculation: String(output.changes.pendingCalculation) }
        })
        result.sourceEtag = object.etag
        await writeWorkbookContext(env.WORKBOOKS, contextR2Key, context, object.etag)
        await stageMutationResult(env, payload.taskId, result)
        await appendTaskEvent(env, payload.taskId, 'mutation.validated', 'Workbook changes verified', { current: output.changes.changedCells, total: output.changes.changedCells, formulaCells: output.changes.formulaCells })
        return result
      } catch (error) {
        if (error instanceof WorkbookMutationError) throw new NonRetryableError(error.message)
        throw error
      }
    })
    await step.do('edit v1 publish version and settle', { retries: { limit: 5, delay: '5 seconds', backoff: 'exponential' } }, () => publishMutationResult(env, payload, result))
    await step.do('edit v1 record completion', () => appendTaskEvent(env, payload.taskId, 'workflow.completed', 'Edited workbook is ready', { changedCells: result.changes.changedCells, formulaCells: result.changes.formulaCells, workbookId: result.workbookId, pendingCalculation: result.changes.pendingCalculation }))
    return result
  } catch (error) {
    // Staged artifacts are recoverable. Never turn a publication outage into a
    // failed/refunded task after an atomic successful publication.
    const record = await getMutationRecord(env, payload.taskId)
    if (record?.state === 'completed') return parseMutationResult(JSON.parse(record.result_json!))
    if (record?.state === 'staged') throw error
    await step.do('edit v1 finalize failure', async () => {
      await discardMutationArtifacts(env, payload.taskId)
      const completedTools = payload.creditToolNames.filter(name => name !== 'editWorkbook')
      if (payload.creditReservationId) await settleTurnCredits(env, { reservationId: payload.creditReservationId, taskId: payload.taskId, actualCredits: calculateActualTurnCredits(completedTools), actions: buildCreditUsageActions(payload.creditTurnId, completedTools) })
      await failTask(env, payload.taskId, error instanceof Error ? error.message : 'Workbook edit failed', 'WORKBOOK_EDIT_FAILED')
      await env.DB.prepare("UPDATE excel_agent_mutations SET state = 'failed', updated_at = ? WHERE task_id = ? AND state = 'queued'").bind(new Date().toISOString(), payload.taskId).run()
      return true
    })
    throw error
  }
}
