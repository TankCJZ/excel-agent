import type { AgentWorkerEnv, WorkbookMutationPayload } from '#agent/types'
import { getMutationRecord } from '#agent/persistence/workbookMutations'
import { requireMutation } from '#agent/workbook/mutation/errors'

export async function dispatchWorkbookMutation(env: AgentWorkerEnv, payload: WorkbookMutationPayload) {
  const id = `edit-${payload.taskId}`
  const record = await getMutationRecord(env, payload.taskId)
  requireMutation(record && record.state === 'queued', 'TASK_CANCELLED', 'The edit is no longer queued.')
  try {
    return await env.EXCEL_AGENT_WORKFLOW.create({ id, params: payload })
  } catch (error) {
    // create may have succeeded before a network error. Probe the SAME id.
    try {
      const instance = await env.EXCEL_AGENT_WORKFLOW.get(id)
      const status = await instance.status()
      if (status.status !== 'unknown') return instance
    } catch { /* Leave durable outbox queued for reconciliation. */ }
    throw error
  }
}
