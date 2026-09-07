import { getConversationForUser } from '#agent/persistence/conversationRepository'
import { listAllTasksForThread } from '#agent/persistence/agentRepository'
import { createExcelAgentMemory } from '#agent/mastra/runtime'
import type { AgentWorkerEnv, ExcelAgentTaskRecord } from '#agent/types'
import { releaseTurnCredits } from '#agent/billing/billingRepository'

const ACTIVE_TASK_STATUSES = new Set(['planning', 'queued', 'processing'])
const SAFE_ARTIFACT_ROOT = 'agent-demo/'
const DELETION_STAGE_ATTEMPTS = 2

export type ConversationDeletionStage =
  | 'load_conversation'
  | 'load_tasks'
  | 'stop_workflows'
  | 'list_artifacts'
  | 'delete_artifacts'
  | 'delete_memory'
  | 'delete_records'

export class ConversationDeletionBusyError extends Error {
  constructor() {
    super('The conversation still has a task that cannot be safely stopped')
    this.name = 'ConversationDeletionBusyError'
  }
}

export class ConversationDeletionStageError extends Error {
  readonly stage: ConversationDeletionStage
  readonly cause: unknown

  constructor(stage: ConversationDeletionStage, cause: unknown) {
    super(`Conversation deletion failed during ${stage}`)
    this.name = 'ConversationDeletionStageError'
    this.stage = stage
    this.cause = cause
  }
}

export async function deleteConversationForUser(
  env: AgentWorkerEnv,
  conversationId: string,
  userId: string
) {
  const startedAt = Date.now()
  const conversation = await runDeletionStage(conversationId, 'load_conversation', () => (
    getConversationForUser(env, conversationId, userId)
  ))
  if (!conversation) return null

  await env.DB.prepare(`INSERT INTO excel_agent_conversation_tombstones (conversation_id,user_id,session_id,deleted_at)
    VALUES (?,?,?,?) ON CONFLICT(conversation_id) DO NOTHING`).bind(conversation.id, userId, conversation.sessionId, new Date().toISOString()).run()

  const tasks = await runDeletionStage(conversation.id, 'load_tasks', () => (
    listAllTasksForThread(env, conversation.id, userId, conversation.sessionId)
  ))
  // Invalidate publication before terminating; termination need not execute
  // a Workflow finally block. Completed publications keep their settled charge.
  for (const task of tasks.filter(task => task.taskType === 'workbook_mutation' && ACTIVE_TASK_STATUSES.has(task.status))) {
    await env.DB.prepare("UPDATE excel_agent_mutations SET state='cancelled', updated_at=? WHERE task_id=? AND state IN ('queued','staged')").bind(new Date().toISOString(), task.id).run()
  }
  await runDeletionStage(conversation.id, 'stop_workflows', () => terminateActiveWorkflows(env, tasks))
  for (const task of tasks.filter(task => task.taskType === 'workbook_mutation')) {
    const reserved = await env.DB.prepare("SELECT id FROM credit_reservations WHERE task_id=? AND user_id=? AND status='reserved'").bind(task.id, userId).all<{ id: string }>()
    for (const hold of reserved.results) await releaseTurnCredits(env, hold.id, 'Conversation deleted before workbook edit publication')
  }

  const artifactKeys = await runDeletionStage(conversation.id, 'list_artifacts', () => (
    collectConversationArtifactKeys(env.WORKBOOKS, conversation.sessionId, tasks)
  ))
  await runDeletionStage(conversation.id, 'delete_artifacts', () => deleteR2Keys(env.WORKBOOKS, artifactKeys))

  await runDeletionStage(conversation.id, 'delete_memory', async () => {
    const { memory } = createExcelAgentMemory(env)
    await memory.deleteThread(conversation.id)
  })

  await runDeletionStage(conversation.id, 'delete_records', async () => {
    const statements: D1PreparedStatement[] = [
      env.DB.prepare(`DELETE FROM excel_agent_workbooks WHERE user_id=? AND source_task_id IN (SELECT id FROM excel_agent_tasks WHERE user_id=? AND thread_id=? AND session_id=?)`).bind(userId, userId, conversation.id, conversation.sessionId),
      env.DB.prepare(`
        DELETE FROM excel_agent_task_events
        WHERE task_id IN (
          SELECT id FROM excel_agent_tasks WHERE user_id = ? AND thread_id = ? AND session_id = ?
        )
      `).bind(userId, conversation.id, conversation.sessionId),
      env.DB.prepare('DELETE FROM excel_agent_tasks WHERE user_id = ? AND thread_id = ? AND session_id = ?')
        .bind(userId, conversation.id, conversation.sessionId),
      env.DB.prepare('DELETE FROM excel_agent_messages WHERE conversation_id = ?').bind(conversation.id),
      env.DB.prepare('DELETE FROM excel_agent_conversations WHERE id = ? AND user_id = ?').bind(conversation.id, userId)
    ]
    await env.DB.batch(statements)
  })

  console.info('Conversation deleted', {
    conversationId: conversation.id,
    taskCount: tasks.length,
    artifactCount: artifactKeys.length,
    durationMs: Date.now() - startedAt
  })

  return {
    conversationId: conversation.id,
    deletedTaskCount: tasks.length,
    deletedArtifactCount: artifactKeys.length
  }
}

async function runDeletionStage<T>(
  conversationId: string,
  stage: ConversationDeletionStage,
  operation: () => Promise<T>
): Promise<T> {
  let lastError: unknown
  for (let attempt = 1; attempt <= DELETION_STAGE_ATTEMPTS; attempt += 1) {
    try {
      return await operation()
    } catch (error) {
      if (error instanceof ConversationDeletionBusyError) throw error
      lastError = error
      console.warn('Conversation deletion stage failed', {
        conversationId,
        stage,
        attempt,
        retrying: attempt < DELETION_STAGE_ATTEMPTS,
        message: error instanceof Error ? error.message : 'Unknown deletion error'
      })
      if (attempt < DELETION_STAGE_ATTEMPTS) {
        await new Promise(resolve => setTimeout(resolve, 25 * attempt))
      }
    }
  }
  throw new ConversationDeletionStageError(stage, lastError)
}

export function conversationArtifactPrefixes(sessionId: string, tasks: ExcelAgentTaskRecord[]) {
  return [
    `agent-demo/research/${sessionId}/`,
    ...tasks.flatMap(task => [
      `agent-demo/tasks/${task.id}/`,
      `agent-demo/results/${task.id}/`
    ])
  ]
}

async function terminateActiveWorkflows(env: AgentWorkerEnv, tasks: ExcelAgentTaskRecord[]) {
  for (const task of tasks) {
    if (!ACTIVE_TASK_STATUSES.has(task.status)) continue
    if (!task.workflowInstanceId) throw new ConversationDeletionBusyError()
    const binding = task.taskType === 'spreadsheet_generation'
      ? env.SPREADSHEET_GENERATION_WORKFLOW
      : env.EXCEL_AGENT_WORKFLOW
    try {
      const instance = await binding.get(task.workflowInstanceId)
      const status = await instance.status()
      if (['queued', 'running', 'paused', 'waiting', 'waitingForPause'].includes(status.status)) {
        await instance.terminate()
      }
    } catch (error) {
      console.error('Failed to stop workflow before deleting conversation', {
        conversationId: task.threadId,
        taskId: task.id,
        message: error instanceof Error ? error.message : 'Unknown workflow error'
      })
      throw new ConversationDeletionBusyError()
    }
  }
}

async function collectConversationArtifactKeys(
  bucket: R2Bucket,
  sessionId: string,
  tasks: ExcelAgentTaskRecord[]
) {
  const keys = new Set<string>()
  for (const prefix of conversationArtifactPrefixes(sessionId, tasks)) {
    let cursor: string | undefined
    do {
      const listed = await bucket.list({ prefix, ...(cursor ? { cursor } : {}), limit: 1000 })
      for (const object of listed.objects) keys.add(object.key)
      cursor = listed.truncated ? listed.cursor : undefined
    } while (cursor)
  }

  for (const task of tasks) {
    if (isOwnedGeneratedArtifactKey(task.inputR2Key, sessionId, task.id)) keys.add(task.inputR2Key)
    if (task.outputR2Key && isOwnedGeneratedArtifactKey(task.outputR2Key, sessionId, task.id)) keys.add(task.outputR2Key)
    const resultKey = task.result?.outputR2Key
    if (resultKey && isOwnedGeneratedArtifactKey(resultKey, sessionId, task.id)) keys.add(resultKey)
  }
  return [...keys].filter(key => key.startsWith(SAFE_ARTIFACT_ROOT))
}

export function isOwnedGeneratedArtifactKey(key: string, sessionId: string, taskId: string) {
  return key.startsWith(`agent-demo/tasks/${taskId}/`)
    || key.startsWith(`agent-demo/results/${taskId}/`)
    || key.startsWith(`agent-demo/research/${sessionId}/`)
}

async function deleteR2Keys(bucket: R2Bucket, keys: string[]) {
  for (let index = 0; index < keys.length; index += 1000) {
    await bucket.delete(keys.slice(index, index + 1000))
  }
}
