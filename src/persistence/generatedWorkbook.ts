import type { AgentWorkerEnv, SpreadsheetGenerationResult, SpreadsheetWorkflowPayload } from '#agent/types'
import { stableMutationId } from '#agent/persistence/workbookMutations'
import { getWorkbookForSession } from '#agent/persistence/agentRepository'
import { parseWorkbookBytes } from '#agent/workbook/service'
import { WORKBOOK_CONTEXT_VERSION, writeWorkbookContext } from '#agent/workbook/context'
import { HttpError } from '#agent/http/responses'

// Register before publication; the completed task and its current selection are
// then published in a single DB batch. Retries always reuse the same workbook ID.
export async function registerGeneratedWorkbook(env: AgentWorkerEnv, payload: SpreadsheetWorkflowPayload, result: SpreadsheetGenerationResult) {
  const selection = payload.workbookSelection
  if (!selection) return result // Old in-flight Workflow payloads remain compatible.
  const workbookId = await stableMutationId(payload.userId, selection.threadId, `adopt:${payload.taskId}`)
  const existing = await getWorkbookForSession(env, workbookId, payload.userId, payload.sessionId)
  if (existing) {
    if (existing.sourceTaskId !== payload.taskId || existing.r2Key !== result.outputR2Key) throw new HttpError(409, 'INVALID_RESULT', 'Workbook registration conflicts with another result')
    return { ...result, workbookId }
  }
  const object = await env.WORKBOOKS.get(result.outputR2Key)
  if (!object) throw new HttpError(410, 'WORKBOOK_FILE_EXPIRED', 'The generated workbook is unavailable')
  const context = await parseWorkbookBytes(await object.arrayBuffer())
  const contextKey = `agent-demo/results/${payload.taskId}/adopted-context-v${WORKBOOK_CONTEXT_VERSION}.json.gz`
  await writeWorkbookContext(env.WORKBOOKS, contextKey, context, object.etag)
  const now = new Date().toISOString()
  await env.DB.prepare(`INSERT INTO excel_agent_workbooks
    (id,user_id,session_id,r2_key,file_name,size_bytes,summary_json,context_r2_key,source_etag,context_version,parsed_at,root_workbook_id,source_task_id,created_at,updated_at)
    SELECT ?,t.user_id,t.session_id,?,?,?,?,?,?,?,?,?,t.id,?,? FROM excel_agent_tasks t
    JOIN excel_agent_conversations c ON c.id=t.thread_id AND c.user_id=t.user_id AND c.session_id=t.session_id
    WHERE t.id=? AND t.user_id=? AND t.session_id=? AND t.thread_id=? AND t.task_type='spreadsheet_generation' AND t.status='processing'
      AND NOT EXISTS (SELECT 1 FROM excel_agent_conversation_tombstones d WHERE d.conversation_id=c.id)
    ON CONFLICT(id) DO NOTHING`).bind(workbookId, result.outputR2Key, result.outputName, object.size,
      JSON.stringify(context.summary), contextKey, object.etag, WORKBOOK_CONTEXT_VERSION, now, workbookId, now, now,
      payload.taskId, payload.userId, payload.sessionId, selection.threadId).run()
  const saved = await getWorkbookForSession(env, workbookId, payload.userId, payload.sessionId)
  if (!saved || saved.sourceTaskId !== payload.taskId || saved.r2Key !== result.outputR2Key) throw new HttpError(409, 'RESULT_NOT_READY', 'This workbook task is no longer active')
  return { ...result, workbookId }
}
