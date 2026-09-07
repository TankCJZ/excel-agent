import type { AgentWorkerEnv } from '#agent/types'
import { getTaskForSession, getWorkbookForSession } from '#agent/persistence/agentRepository'
import { getConversationForUser, assertConversationWritable } from '#agent/persistence/conversationRepository'
import { stableMutationId } from '#agent/persistence/workbookMutations'
import { parseWorkbookBytes } from '#agent/workbook/service'
import { WORKBOOK_CONTEXT_VERSION, writeWorkbookContext } from '#agent/workbook/context'
import { HttpError } from '#agent/http/responses'

// Old generated files become registered inputs lazily. Never depend on a
// deserializable result_json or charge a download just to continue editing.
export async function adoptTaskWorkbook(env: AgentWorkerEnv, input: {
  userId: string; sessionId: string; threadId: string; taskId: string; workbookRevision: number
}) {
  const conversation = await getConversationForUser(env, input.threadId, input.userId)
  const stored = await getTaskForSession(env, input.taskId, input.userId, input.sessionId)
  if (!conversation || conversation.sessionId !== input.sessionId || !stored || stored.task.threadId !== input.threadId) throw new HttpError(404, 'NOT_FOUND', 'Workbook result was not found in this conversation')
  await assertConversationWritable(env, input.threadId)
  const task = stored.task
  if (task.status !== 'completed' || !task.outputR2Key) throw new HttpError(409, 'RESULT_NOT_READY', 'Wait for the workbook to finish before using it')
  // A completed mutation already owns a version. A generated/exported legacy
  // task gets an isolated deterministic ID, making retries idempotent.
  const workbookId = task.result?.kind === 'workbook_mutation' ? task.result.workbookId : await stableMutationId(input.userId, input.threadId, `adopt:${input.taskId}`)
  const existing = await getWorkbookForSession(env, workbookId, input.userId, input.sessionId)
  const object = await env.WORKBOOKS.get(existing?.r2Key || task.outputR2Key)
  if (!object) throw new HttpError(410, 'WORKBOOK_FILE_EXPIRED', 'This workbook file is no longer available. Upload your downloaded copy to continue')
  if (existing && conversation.workbookId === workbookId) return { conversation, workbook: publicWorkbook(existing) }
  if (conversation.workbookRevision !== input.workbookRevision) throw new HttpError(409, 'WORKBOOK_VERSION_CONFLICT', 'The attached workbook changed. Reload the conversation and choose the result again')
  if (!existing) {
    const context = await parseWorkbookBytes(await object.arrayBuffer())
    const contextKey = `agent-demo/results/${task.id}/adopted-context-v${WORKBOOK_CONTEXT_VERSION}.json.gz`
    await writeWorkbookContext(env.WORKBOOKS, contextKey, context, object.etag)
    const now = new Date().toISOString()
    await env.DB.prepare(`INSERT INTO excel_agent_workbooks
      (id,user_id,session_id,r2_key,file_name,size_bytes,summary_json,context_r2_key,source_etag,context_version,parsed_at,parent_workbook_id,root_workbook_id,source_task_id,created_at,updated_at)
      SELECT ?,t.user_id,t.session_id,t.output_r2_key,?,?,?, ?,?,?,?,w.id,COALESCE(w.root_workbook_id,w.id,?),t.id,?,?
      FROM excel_agent_tasks t LEFT JOIN excel_agent_workbooks w ON w.r2_key=t.input_r2_key AND w.user_id=t.user_id AND w.session_id=t.session_id
      JOIN excel_agent_conversations c ON c.id=t.thread_id AND c.user_id=t.user_id AND c.session_id=t.session_id
      WHERE t.id=? AND t.user_id=? AND t.session_id=? AND t.thread_id=? AND t.status='completed' AND t.output_r2_key=?
        AND NOT EXISTS (SELECT 1 FROM excel_agent_conversation_tombstones d WHERE d.conversation_id=c.id)
      LIMIT 1 ON CONFLICT(id) DO NOTHING`)
      .bind(workbookId, task.outputName || 'excelgen-result.xlsx', object.size, JSON.stringify(context.summary), contextKey, object.etag, WORKBOOK_CONTEXT_VERSION, now, workbookId, now, now, task.id, input.userId, input.sessionId, input.threadId, task.outputR2Key).run()
  }
  const workbook = await getWorkbookForSession(env, workbookId, input.userId, input.sessionId)
  if (!workbook || (workbook.sourceTaskId !== task.id && task.result?.kind !== 'workbook_mutation')) throw new HttpError(409, 'INVALID_RESULT', 'This result could not be registered as a workbook')
  await env.DB.prepare(`UPDATE excel_agent_conversations SET workbook_id=?,workbook_revision=workbook_revision+1,updated_at=?
    WHERE id=? AND user_id=? AND session_id=? AND workbook_revision=?
      AND NOT EXISTS (SELECT 1 FROM excel_agent_conversation_tombstones WHERE conversation_id=?)`)
    .bind(workbookId, new Date().toISOString(), input.threadId, input.userId, input.sessionId, input.workbookRevision, input.threadId).run()
  const current = await getConversationForUser(env, input.threadId, input.userId)
  if (!current || current.workbookId !== workbookId) throw new HttpError(409, 'WORKBOOK_VERSION_CONFLICT', 'The attached workbook changed. Reload the conversation and choose the result again')
  return { conversation: current, workbook: publicWorkbook(workbook) }
}

function publicWorkbook(workbook: NonNullable<Awaited<ReturnType<typeof getWorkbookForSession>>>) {
  return { workbookId: workbook.id, fileName: workbook.fileName, sizeBytes: workbook.sizeBytes, summary: workbook.summary }
}
