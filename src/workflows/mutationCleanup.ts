import type { AgentWorkerEnv } from '#agent/types'
import { requireMutation } from '#agent/workbook/mutation/errors'
import { WORKBOOK_CONTEXT_VERSION } from '#agent/workbook/context'

// Only unpublished, task-owned output keys are disposable. Never derive a
// deletion target from a saved result/payload, which may itself be corrupted.
export async function discardMutationArtifacts(env: AgentWorkerEnv, taskId: string, stopWorkflow = false) {
  requireMutation(/^[a-zA-Z0-9_-]+$/.test(taskId), 'INVALID_RESULT', 'Invalid workbook edit identity.')
  const task = await env.DB.prepare(`SELECT t.status, t.input_r2_key, t.workflow_instance_id, m.state
    FROM excel_agent_tasks t JOIN excel_agent_mutations m ON m.task_id=t.id WHERE t.id=?`).bind(taskId)
    .first<{ status: string; input_r2_key: string; workflow_instance_id: string | null; state: string }>()
  if (!task || task.status === 'completed' || !['queued', 'failing', 'cancelled'].includes(task.state)) return
  if (stopWorkflow && task.workflow_instance_id) {
    const instance = await env.EXCEL_AGENT_WORKFLOW.get(task.workflow_instance_id)
    const status = await instance.status()
    if (!['complete', 'errored', 'terminated', 'unknown'].includes(status.status)) await instance.terminate()
  }
  const keys = [`agent-demo/results/${taskId}/edited.xlsx`, `agent-demo/results/${taskId}/context-v${WORKBOOK_CONTEXT_VERSION}.json.gz`]
  requireMutation(!keys.includes(task.input_r2_key), 'INVALID_RESULT', 'The source workbook cannot be discarded.')
  await env.WORKBOOKS.delete(keys)
}
