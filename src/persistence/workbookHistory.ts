import { workbookSummarySchema } from '#agent/contracts'
import { getConversationForUser, assertConversationWritable, updateConversation } from '#agent/persistence/conversationRepository'
import { adoptTaskWorkbook } from '#agent/persistence/adoptTaskWorkbook'
import { decodeStoredJson } from '#agent/persistence/taskProtocol'
import { parseWorkbookBytes } from '#agent/workbook/service'
import { HttpError } from '#agent/http/responses'
import type { AgentWorkerEnv } from '#agent/types'
import type { WorkbookHistoryEntry, WorkbookHistoryPage } from '#shared/agent/workbookSelection'

// Rank the full lineage before pagination so a branch/version number does not
// change as the user loads older history. Unregistered legacy outputs remain
// visible without a GET request mutating the registry or changing selection.
const historyQuery = `WITH files AS (
  SELECT w.id,w.id AS workbook_id,w.source_task_id AS task_id,COALESCE(w.root_workbook_id,w.id) AS root_id,
    w.parent_workbook_id,w.r2_key,w.file_name,COALESCE(root.file_name,w.file_name) AS display_name,
    w.size_bytes,w.summary_json,COALESCE(t.completed_at,w.created_at) AS created_at,
    COALESCE(t.prompt,'') AS description,COALESCE(t.task_type,'upload') AS kind
  FROM excel_agent_conversation_workbooks h JOIN excel_agent_workbooks w ON w.id=h.workbook_id
  JOIN excel_agent_conversations c ON c.id=h.conversation_id AND c.user_id=w.user_id AND c.session_id=w.session_id
  LEFT JOIN excel_agent_tasks t ON t.id=w.source_task_id AND t.thread_id=c.id AND t.user_id=c.user_id AND t.session_id=c.session_id
  LEFT JOIN excel_agent_workbooks root ON root.id=w.root_workbook_id AND root.user_id=w.user_id AND root.session_id=w.session_id
  WHERE c.id=? AND c.user_id=? AND (w.source_task_id IS NULL OR t.status='completed')
  UNION ALL
  SELECT 'task:'||t.id,NULL,t.id,'task:'||t.id,NULL,t.output_r2_key,t.output_name,t.output_name,
    NULL,CASE WHEN json_valid(t.result_json) THEN json_extract(t.result_json,'$.summary') ELSE NULL END,
    COALESCE(t.completed_at,t.created_at),t.prompt,t.task_type
  FROM excel_agent_tasks t JOIN excel_agent_conversations c ON c.id=t.thread_id AND c.user_id=t.user_id AND c.session_id=t.session_id
  WHERE c.id=? AND c.user_id=? AND t.status='completed' AND t.output_r2_key IS NOT NULL
    AND t.task_type IN ('spreadsheet_generation','workbook_export')
    AND NOT EXISTS (SELECT 1 FROM excel_agent_workbooks w WHERE w.source_task_id=t.id AND w.user_id=t.user_id AND w.session_id=t.session_id)
), ranked AS (
  SELECT *,ROW_NUMBER() OVER (PARTITION BY root_id ORDER BY created_at,id) AS version,
    ROW_NUMBER() OVER (PARTITION BY root_id ORDER BY created_at DESC,id DESC)=1 AS is_latest FROM files
)`

type HistoryRow = {
  id: string; workbook_id: string | null; task_id: string | null; root_id: string; parent_workbook_id: string | null
  r2_key: string; file_name: string; display_name: string; size_bytes: number | null; summary_json: string | null
  created_at: string; description: string; kind: WorkbookHistoryEntry['kind']; version: number; is_latest: number
}

async function requireConversation(env: AgentWorkerEnv, conversationId: string, userId: string) {
  const conversation = await getConversationForUser(env, conversationId, userId)
  if (!conversation) throw new HttpError(404, 'NOT_FOUND', 'Conversation was not found')
  await assertConversationWritable(env, conversationId)
  return conversation
}

function publicEntry(row: HistoryRow): WorkbookHistoryEntry {
  const summary = workbookSummarySchema.safeParse(decodeStoredJson(row.summary_json))
  // Keep the original document name, but show the actual version's format:
  // editing a CSV produces an XLSX, not another CSV.
  const extension = /\.(?:xlsx|xls|csv)$/i.exec(row.file_name)?.[0]
  const displayName = extension ? row.display_name.replace(/\.(?:xlsx|xls|csv)$/i, '') + extension : row.display_name
  return {
    id: row.id, workbookId: row.workbook_id, taskId: row.task_id, rootId: row.root_id,
    parentWorkbookId: row.parent_workbook_id, fileName: row.file_name, displayName,
    version: row.version, isLatest: Boolean(row.is_latest), createdAt: row.created_at,
    description: row.description.slice(0, 180), kind: row.kind,
    rowCount: summary.success ? summary.data.rowCount : null,
    sheetCount: summary.success ? summary.data.sheetNames.length : null,
    calculationPending: summary.success && summary.data.calculation?.status === 'pending'
  }
}

export async function listWorkbookHistory(env: AgentWorkerEnv, conversationId: string, userId: string, offset = 0): Promise<WorkbookHistoryPage> {
  await requireConversation(env, conversationId, userId)
  const { results } = await env.DB.prepare(`${historyQuery} SELECT * FROM ranked ORDER BY created_at DESC,id DESC LIMIT 51 OFFSET ?`)
    .bind(conversationId, userId, conversationId, userId, offset).all<HistoryRow>()
  return { entries: results.slice(0, 50).map(publicEntry), nextOffset: results.length > 50 ? offset + 50 : null }
}

async function historyRow(env: AgentWorkerEnv, conversationId: string, userId: string, entryId: string) {
  await requireConversation(env, conversationId, userId)
  const row = await env.DB.prepare(`${historyQuery} SELECT * FROM ranked WHERE id=? LIMIT 1`)
    .bind(conversationId, userId, conversationId, userId, entryId).first<HistoryRow>()
  if (!row) throw new HttpError(404, 'NOT_FOUND', 'Workbook was not found in this conversation')
  return row
}

export async function previewHistoryWorkbook(env: AgentWorkerEnv, conversationId: string, userId: string, entryId: string) {
  const row = await historyRow(env, conversationId, userId, entryId)
  const object = await env.WORKBOOKS.head(row.r2_key)
  if (!object) throw new HttpError(410, 'WORKBOOK_FILE_EXPIRED', 'This workbook is no longer available. Upload your saved copy to continue')
  const parsed = workbookSummarySchema.safeParse(decodeStoredJson(row.summary_json))
  let summary = parsed.success ? parsed.data : null
  if (!summary) {
    const file = await env.WORKBOOKS.get(row.r2_key)
    if (!file) throw new HttpError(410, 'WORKBOOK_FILE_EXPIRED', 'This workbook is no longer available. Upload your saved copy to continue')
    summary = (await parseWorkbookBytes(await file.arrayBuffer())).summary
  }
  return { entry: publicEntry(row), workbook: { workbookId: row.workbook_id || row.id, fileName: row.file_name, sizeBytes: object.size, summary } }
}

export async function selectHistoryWorkbook(env: AgentWorkerEnv, input: { conversationId: string; userId: string; entryId: string; workbookRevision: number }) {
  const conversation = await requireConversation(env, input.conversationId, input.userId)
  const row = await historyRow(env, input.conversationId, input.userId, input.entryId)
  // Even an idempotent selection must check file availability, never resurrect a
  // missing binary from its cached summary or historical model messages.
  const preview = await previewHistoryWorkbook(env, input.conversationId, input.userId, input.entryId)
  if (!row.workbook_id && row.task_id) return adoptTaskWorkbook(env, {
    userId: input.userId, sessionId: conversation.sessionId, threadId: conversation.id,
    taskId: row.task_id, workbookRevision: input.workbookRevision
  })
  if (conversation.workbookRevision !== input.workbookRevision) throw new HttpError(409, 'WORKBOOK_VERSION_CONFLICT', 'The current workbook changed. Refresh the file list before choosing again')
  if (conversation.workbookId === row.workbook_id) return { conversation, workbook: preview.workbook }
  const updated = await updateConversation(env, {
    conversationId: conversation.id, userId: input.userId, sessionId: conversation.sessionId,
    workbookId: row.workbook_id, workbookRevision: input.workbookRevision
  })
  return { conversation: updated, workbook: preview.workbook }
}
