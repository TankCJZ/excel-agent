import { index, integer, primaryKey, sqliteTable, text, uniqueIndex } from 'drizzle-orm/sqlite-core'

export const excelAgentWorkbooks = sqliteTable('excel_agent_workbooks', {
  id: text('id').primaryKey(),
  userId: text('user_id').notNull().default('legacy-demo'),
  sessionId: text('session_id').notNull(),
  r2Key: text('r2_key').notNull(),
  fileName: text('file_name').notNull(),
  sizeBytes: integer('size_bytes').notNull(),
  summaryJson: text('summary_json').notNull(),
  contextR2Key: text('context_r2_key'),
  sourceEtag: text('source_etag'),
  contextVersion: integer('context_version'),
  parsedAt: text('parsed_at'),
  parentWorkbookId: text('parent_workbook_id'),
  rootWorkbookId: text('root_workbook_id'),
  sourceTaskId: text('source_task_id'),
  createdAt: text('created_at').notNull(),
  updatedAt: text('updated_at').notNull()
}, (table) => ({
  userIdx: index('excel_agent_workbooks_user_id_idx').on(table.userId),
  ownerSessionIdx: index('excel_agent_workbooks_owner_session_idx').on(table.userId, table.sessionId),
  sessionIdx: index('excel_agent_workbooks_session_id_idx').on(table.sessionId),
  r2KeyUnique: uniqueIndex('excel_agent_workbooks_r2_key_unique').on(table.r2Key),
  createdAtIdx: index('excel_agent_workbooks_created_at_idx').on(table.createdAt)
}))

export const excelAgentConversations = sqliteTable('excel_agent_conversations', {
  id: text('id').primaryKey(),
  userId: text('user_id').notNull(),
  sessionId: text('session_id').notNull(),
  title: text('title').notNull().default(''),
  workbookId: text('workbook_id'),
  workbookRevision: integer('workbook_revision').notNull().default(0),
  createdAt: text('created_at').notNull(),
  updatedAt: text('updated_at').notNull(),
  lastMessageAt: text('last_message_at')
}, (table) => ({
  userIdx: index('excel_agent_conversations_user_id_idx').on(table.userId),
  userUpdatedIdx: index('excel_agent_conversations_user_updated_idx').on(table.userId, table.updatedAt),
  sessionIdx: index('excel_agent_conversations_session_id_idx').on(table.sessionId)
}))

// Kept after deletion to fence late model requests and Workflow dispatches.
export const excelAgentConversationTombstones = sqliteTable('excel_agent_conversation_tombstones', {
  conversationId: text('conversation_id').primaryKey(),
  userId: text('user_id').notNull(),
  sessionId: text('session_id').notNull(),
  deletedAt: text('deleted_at').notNull()
})

// Attachment history is independent of the current selection, including files
// removed before the first chat message. Deleting a conversation removes links only.
export const excelAgentConversationWorkbooks = sqliteTable('excel_agent_conversation_workbooks', {
  conversationId: text('conversation_id').notNull().references(() => excelAgentConversations.id, { onDelete: 'cascade' }),
  workbookId: text('workbook_id').notNull().references(() => excelAgentWorkbooks.id, { onDelete: 'cascade' }),
  addedAt: text('added_at').notNull()
}, table => ({
  pk: primaryKey({ columns: [table.conversationId, table.workbookId] })
}))

export const excelAgentTasks = sqliteTable('excel_agent_tasks', {
  id: text('id').primaryKey(),
  userId: text('user_id').notNull().default('legacy-demo'),
  sessionId: text('session_id').notNull(),
  threadId: text('thread_id').notNull(),
  taskType: text('task_type').notNull().default('workbook_export'),
  workflowInstanceId: text('workflow_instance_id'),
  inputR2Key: text('input_r2_key').notNull(),
  inputName: text('input_name').notNull(),
  outputR2Key: text('output_r2_key'),
  outputName: text('output_name'),
  prompt: text('prompt').notNull(),
  planJson: text('plan_json'),
  resultJson: text('result_json'),
  status: text('status').notNull(),
  errorMessage: text('error_message'),
  createdAt: text('created_at').notNull(),
  updatedAt: text('updated_at').notNull(),
  completedAt: text('completed_at')
}, (table) => ({
  userIdx: index('excel_agent_tasks_user_id_idx').on(table.userId),
  ownerThreadIdx: index('excel_agent_tasks_owner_thread_idx').on(table.userId, table.threadId),
  sessionIdx: index('excel_agent_tasks_session_id_idx').on(table.sessionId),
  threadIdx: index('excel_agent_tasks_thread_id_idx').on(table.threadId),
  statusIdx: index('excel_agent_tasks_status_idx').on(table.status),
  createdAtIdx: index('excel_agent_tasks_created_at_idx').on(table.createdAt)
}))

export const excelAgentMessages = sqliteTable('excel_agent_messages', {
  id: text('id').primaryKey(),
  conversationId: text('conversation_id').notNull().references(() => excelAgentConversations.id, { onDelete: 'cascade' }),
  turnId: text('turn_id').notNull(),
  role: text('role').notNull(),
  status: text('status').notNull(),
  content: text('content').notNull().default(''),
  taskId: text('task_id'),
  workbookId: text('workbook_id'),
  metadataJson: text('metadata_json'),
  errorCode: text('error_code'),
  createdAt: text('created_at').notNull(),
  updatedAt: text('updated_at').notNull()
}, (table) => ({
  conversationIdx: index('excel_agent_messages_conversation_id_idx').on(table.conversationId),
  conversationCreatedIdx: index('excel_agent_messages_conversation_created_idx').on(table.conversationId, table.createdAt),
  taskIdx: index('excel_agent_messages_task_id_idx').on(table.taskId),
  turnRoleUnique: uniqueIndex('excel_agent_messages_turn_role_unique').on(table.conversationId, table.turnId, table.role)
}))

export const excelAgentTaskEvents = sqliteTable('excel_agent_task_events', {
  id: text('id').primaryKey(),
  taskId: text('task_id').notNull().references(() => excelAgentTasks.id, { onDelete: 'cascade' }),
  type: text('type').notNull(),
  message: text('message').notNull(),
  metadataJson: text('metadata_json'),
  createdAt: text('created_at').notNull()
}, (table) => ({
  taskIdx: index('excel_agent_task_events_task_id_idx').on(table.taskId),
  createdAtIdx: index('excel_agent_task_events_created_at_idx').on(table.createdAt)
}))

export const excelAgentRateLimits = sqliteTable('excel_agent_rate_limits', {
  key: text('key').primaryKey(),
  count: integer('count').notNull(),
  expiresAt: text('expires_at').notNull(),
  updatedAt: text('updated_at').notNull()
}, (table) => ({
  expiresAtIdx: index('excel_agent_rate_limits_expires_at_idx').on(table.expiresAt)
}))

// Durable dispatch/finalization record. Old task payloads and results stay intact.
export const excelAgentMutations = sqliteTable('excel_agent_mutations', {
  taskId: text('task_id').primaryKey().references(() => excelAgentTasks.id, { onDelete: 'cascade' }),
  conversationId: text('conversation_id').notNull(),
  turnId: text('turn_id').notNull(),
  payloadJson: text('payload_json').notNull(),
  resultJson: text('result_json'),
  state: text('state').notNull().default('queued'),
  createdAt: text('created_at').notNull(),
  updatedAt: text('updated_at').notNull()
}, table => ({
  turnUnique: uniqueIndex('excel_agent_mutations_conversation_turn_unique').on(table.conversationId, table.turnId),
  stateIdx: index('excel_agent_mutations_state_idx').on(table.state, table.updatedAt)
}))

export type ExcelAgentTask = typeof excelAgentTasks.$inferSelect
export type ExcelAgentTaskEvent = typeof excelAgentTaskEvents.$inferSelect
export type ExcelAgentConversation = typeof excelAgentConversations.$inferSelect
export type ExcelAgentMessage = typeof excelAgentMessages.$inferSelect
