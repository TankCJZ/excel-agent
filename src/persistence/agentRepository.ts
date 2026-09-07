import { and, asc, eq, inArray, lt, ne, sql } from 'drizzle-orm'
import { drizzle } from 'drizzle-orm/d1'
// This repository is the only Worker module that maps Agent domain records to D1 rows.
import {
  excelAgentRateLimits,
  excelAgentConversations,
  excelAgentTaskEvents,
  excelAgentTasks,
  excelAgentWorkbooks
} from '#shared/database/agentSchema'
import { workbookSummarySchema, type WorkbookSummary } from '#agent/contracts'
import { decodeStoredJson, readTaskPlan, readTaskResult } from '#agent/persistence/taskProtocol'
import type {
  AgentTaskPlan,
  AgentTaskResult,
  AgentWorkerEnv,
  ExcelAgentTaskEvent,
  ExcelAgentTaskRecord,
  ExcelAgentTaskType,
  ExcelAgentWorkbookRecord
} from '#agent/types'

function getAgentDb(env: AgentWorkerEnv) {
  return drizzle(env.DB)
}

export async function createExcelAgentWorkbook(
  env: AgentWorkerEnv,
  input: {
    id: string
    userId: string
    sessionId: string
    r2Key: string
    fileName: string
    sizeBytes: number
    summary: WorkbookSummary
    contextR2Key: string
    sourceEtag: string
    contextVersion: number
    parsedAt: string
  }
) {
  const now = new Date().toISOString()
  await getAgentDb(env).insert(excelAgentWorkbooks).values({
    id: input.id,
    userId: input.userId,
    sessionId: input.sessionId,
    r2Key: input.r2Key,
    fileName: input.fileName,
    sizeBytes: input.sizeBytes,
    summaryJson: JSON.stringify(input.summary),
    contextR2Key: input.contextR2Key,
    sourceEtag: input.sourceEtag,
    contextVersion: input.contextVersion,
    parsedAt: input.parsedAt,
    createdAt: now,
    updatedAt: now
  })

  return {
    workbookId: input.id,
    fileName: input.fileName,
    sizeBytes: input.sizeBytes,
    summary: input.summary
  }
}

export async function getWorkbookForSession(
  env: AgentWorkerEnv,
  workbookId: string,
  userId: string,
  sessionId: string
): Promise<ExcelAgentWorkbookRecord | null> {
  const [row] = await getAgentDb(env)
    .select()
    .from(excelAgentWorkbooks)
    .where(and(
      eq(excelAgentWorkbooks.id, workbookId),
      eq(excelAgentWorkbooks.userId, userId),
      eq(excelAgentWorkbooks.sessionId, sessionId)
    ))
    .limit(1)
  if (!row) return null
  return mapWorkbookRow(row)
}

export async function updateWorkbookContextCache(
  env: AgentWorkerEnv,
  input: {
    workbookId: string
    userId: string
    sessionId: string
    contextR2Key: string
    sourceEtag: string
    contextVersion: number
    parsedAt: string
    summary: WorkbookSummary
  }
) {
  await getAgentDb(env)
    .update(excelAgentWorkbooks)
    .set({
      contextR2Key: input.contextR2Key,
      sourceEtag: input.sourceEtag,
      contextVersion: input.contextVersion,
      parsedAt: input.parsedAt,
      summaryJson: JSON.stringify(input.summary),
      updatedAt: new Date().toISOString()
    })
    .where(and(
      eq(excelAgentWorkbooks.id, input.workbookId),
      eq(excelAgentWorkbooks.userId, input.userId),
      eq(excelAgentWorkbooks.sessionId, input.sessionId)
    ))
}

export async function createExcelAgentTask(
  env: AgentWorkerEnv,
  input: {
    id: string
    userId: string
    sessionId: string
    threadId: string
    inputR2Key: string
    inputName: string
    prompt: string
    taskType?: ExcelAgentTaskType
  }
) {
  const now = new Date().toISOString()
  const db = getAgentDb(env)
  await db.batch([
    db.insert(excelAgentTasks).values({
      id: input.id,
      userId: input.userId,
      sessionId: input.sessionId,
      threadId: input.threadId,
      taskType: input.taskType || 'workbook_export',
      inputR2Key: input.inputR2Key,
      inputName: input.inputName,
      prompt: input.prompt,
      status: 'planning',
      createdAt: now,
      updatedAt: now
    }),
    db.insert(excelAgentTaskEvents).values(taskEvent(input.id, 'task.created', 'Task persisted to D1', now))
  ])
}

export async function saveTaskPlan(env: AgentWorkerEnv, taskId: string, plan: AgentTaskPlan) {
  const now = new Date().toISOString()
  const db = getAgentDb(env)
  await db.batch([
    db.update(excelAgentTasks)
      .set({ planJson: JSON.stringify(plan), status: 'queued', updatedAt: now })
      .where(eq(excelAgentTasks.id, taskId)),
    db.insert(excelAgentTaskEvents).values(taskEvent(
      taskId,
      'agent.plan',
      'Mastra generated a structured execution plan',
      now,
      { action: plan.action, stepCount: plan.steps.length }
    ))
  ])
}

export async function saveWorkflowInstance(env: AgentWorkerEnv, taskId: string, workflowInstanceId: string) {
  const now = new Date().toISOString()
  const db = getAgentDb(env)
  await db.batch([
    db.update(excelAgentTasks)
      .set({ workflowInstanceId, updatedAt: now })
      .where(eq(excelAgentTasks.id, taskId)),
    db.insert(excelAgentTaskEvents).values(taskEvent(
      taskId,
      'workflow.queued',
      'Cloudflare Workflow instance created',
      now,
      { workflowInstanceId }
    ))
  ])
}

export async function markTaskProcessing(
  env: AgentWorkerEnv,
  taskId: string,
  message = 'Workflow started processing the R2 workbook'
) {
  const now = new Date().toISOString()
  const db = getAgentDb(env)
  await db.batch([
    db.update(excelAgentTasks)
      .set({ status: 'processing', updatedAt: now })
      .where(and(
        eq(excelAgentTasks.id, taskId),
        ne(excelAgentTasks.status, 'processing'),
        ne(excelAgentTasks.status, 'completed'),
        ne(excelAgentTasks.status, 'failed')
      )),
    conditionalTaskEvent(
      db,
      taskId,
      'processing',
      'workflow.processing',
      message,
      now
    )
  ])
}

export async function completeTask(env: AgentWorkerEnv, taskId: string, result: AgentTaskResult, selection?: { threadId: string; workbookId: string | null; revision: number; userId: string; sessionId: string }) {
  const now = new Date().toISOString()
  const db = getAgentDb(env)
  const writes: [Parameters<typeof db.batch>[0][number], ...Parameters<typeof db.batch>[0][number][]] = [
    db.update(excelAgentTasks)
      .set({
        status: 'completed',
        outputR2Key: result.outputR2Key,
        outputName: result.outputName,
        resultJson: JSON.stringify(result),
        errorMessage: null,
        updatedAt: now,
        completedAt: now
      })
      .where(and(
        eq(excelAgentTasks.id, taskId),
        ne(excelAgentTasks.status, 'completed'),
        ne(excelAgentTasks.status, 'failed')
      )),
    conditionalTaskEvent(db, taskId, 'completed', 'workflow.completed',
      'Result workbook saved to R2', now, {
        outputR2Key: result.outputR2Key,
        sizeBytes: result.sizeBytes
      })
  ]
  // Only generation of a new workbook changes selection here. Analysis exports
  // never replace their source. A stale completion cannot undo a detach/switch.
  if (selection && result.kind === 'spreadsheet_generation' && result.workbookId) {
    writes.push(db.update(excelAgentConversations).set({
      workbookId: result.workbookId,
      workbookRevision: sql`${excelAgentConversations.workbookRevision} + 1`,
      updatedAt: now
    }).where(and(
      eq(excelAgentConversations.id, selection.threadId), eq(excelAgentConversations.userId, selection.userId),
      eq(excelAgentConversations.sessionId, selection.sessionId),
      sql`${excelAgentConversations.workbookId} IS ${selection.workbookId}`,
      eq(excelAgentConversations.workbookRevision, selection.revision),
      sql`EXISTS (SELECT 1 FROM excel_agent_tasks t JOIN excel_agent_workbooks w ON w.source_task_id=t.id
        WHERE t.id=${taskId} AND t.status='completed' AND t.thread_id=${selection.threadId}
          AND t.user_id=${selection.userId} AND w.id=${result.workbookId})`
    )))
  }
  await db.batch(writes)
}

export async function failTask(env: AgentWorkerEnv, taskId: string, errorMessage: string, errorCode = 'WORKFLOW_FAILED') {
  const db = getAgentDb(env)
  const now = new Date().toISOString()
  await db.batch([
    db.update(excelAgentTasks)
      .set({
        status: 'failed',
        errorMessage: errorMessage.slice(0, 800),
        updatedAt: now,
        completedAt: now
      })
      .where(and(
        eq(excelAgentTasks.id, taskId),
        ne(excelAgentTasks.status, 'completed'),
        ne(excelAgentTasks.status, 'failed')
      )),
    conditionalTaskEvent(db, taskId, 'failed', 'workflow.failed', 'Workflow failed', now, {
      error: errorMessage.slice(0, 400), errorCode
    })
  ])
}

export async function appendTaskEvent(
  env: AgentWorkerEnv,
  taskId: string,
  type: string,
  message: string,
  metadata?: Record<string, unknown>
) {
  const now = new Date().toISOString()
  await getAgentDb(env)
    .insert(excelAgentTaskEvents)
    .values(idempotentTaskEvent(taskId, type, message, now, metadata))
    .onConflictDoNothing({ target: excelAgentTaskEvents.id })
}

export async function getTaskForSession(env: AgentWorkerEnv, taskId: string, userId: string, sessionId: string) {
  const db = getAgentDb(env)
  const [tasks, events] = await Promise.all([
    db.select()
      .from(excelAgentTasks)
      .where(and(
        eq(excelAgentTasks.id, taskId),
        eq(excelAgentTasks.userId, userId),
        eq(excelAgentTasks.sessionId, sessionId)
      ))
      .limit(1),
    db.select()
      .from(excelAgentTaskEvents)
      .where(eq(excelAgentTaskEvents.taskId, taskId))
      .orderBy(asc(excelAgentTaskEvents.createdAt))
  ])
  const task = tasks[0]
  if (!task) return null
  return {
    task: mapTaskRow(task),
    events: events.map(mapEventRow)
  }
}

export async function listTasksForThread(env: AgentWorkerEnv, threadId: string, userId: string, sessionId: string) {
  const rows = await getAgentDb(env)
    .select()
    .from(excelAgentTasks)
    .where(and(
      eq(excelAgentTasks.threadId, threadId),
      eq(excelAgentTasks.userId, userId),
      eq(excelAgentTasks.sessionId, sessionId)
    ))
    .orderBy(asc(excelAgentTasks.createdAt))
    .limit(100)
  return rows.map(mapTaskRow)
}

export async function listAllTasksForThread(env: AgentWorkerEnv, threadId: string, userId: string, sessionId: string) {
  const rows = await getAgentDb(env)
    .select()
    .from(excelAgentTasks)
    .where(and(
      eq(excelAgentTasks.threadId, threadId),
      eq(excelAgentTasks.userId, userId),
      eq(excelAgentTasks.sessionId, sessionId)
    ))
    .orderBy(asc(excelAgentTasks.createdAt))
  return rows.map(mapTaskRow)
}

export async function getThreadTaskHistory(env: AgentWorkerEnv, threadId: string, userId: string, sessionId: string) {
  const tasks = await listTasksForThread(env, threadId, userId, sessionId)
  if (!tasks.length) return { tasks, events: [] as ExcelAgentTaskEvent[] }
  const rows = await getAgentDb(env)
    .select()
    .from(excelAgentTaskEvents)
    .where(inArray(excelAgentTaskEvents.taskId, tasks.map(task => task.id)))
    .orderBy(asc(excelAgentTaskEvents.createdAt))
  return { tasks, events: rows.map(mapEventRow) }
}

export async function consumeAgentRateLimit(
  env: AgentWorkerEnv,
  key: string,
  expiresAt: string,
  updatedAt: string
) {
  const db = getAgentDb(env)
  const [result] = await db
    .insert(excelAgentRateLimits)
    .values({ key, count: 1, expiresAt, updatedAt })
    .onConflictDoUpdate({
      target: excelAgentRateLimits.key,
      set: {
        count: sql`${excelAgentRateLimits.count} + 1`,
        updatedAt
      }
    })
    .returning({ count: excelAgentRateLimits.count })
  if (result?.count === 1) {
    await db.delete(excelAgentRateLimits).where(lt(excelAgentRateLimits.expiresAt, updatedAt))
  }
  return result?.count || 1
}

export async function checkAgentPersistence(env: AgentWorkerEnv) {
  const db = getAgentDb(env)
  await Promise.all([
    db.select({ id: excelAgentTasks.id }).from(excelAgentTasks).limit(1),
    db.select({ id: excelAgentWorkbooks.id }).from(excelAgentWorkbooks).limit(1),
    db.select({ key: excelAgentRateLimits.key }).from(excelAgentRateLimits).limit(1)
  ])
  return true
}

function taskEvent(
  taskId: string,
  type: string,
  message: string,
  createdAt: string,
  metadata?: Record<string, unknown>
) {
  return {
    id: crypto.randomUUID(),
    taskId,
    type,
    message,
    metadataJson: metadata ? JSON.stringify(metadata) : null,
    createdAt
  }
}

function idempotentTaskEvent(
  taskId: string,
  type: string,
  message: string,
  createdAt: string,
  metadata?: Record<string, unknown>
) {
  return {
    ...taskEvent(taskId, type, message, createdAt, metadata),
    id: workflowEventId(taskId, type)
  }
}

export function workflowEventId(taskId: string, type: string) {
  return `workflow:${taskId}:${type}`
}

function conditionalTaskEvent(
  db: ReturnType<typeof getAgentDb>,
  taskId: string,
  status: string,
  type: string,
  message: string,
  createdAt: string,
  metadata?: Record<string, unknown>
) {
  const metadataJson = metadata ? JSON.stringify(metadata) : null
  return db.insert(excelAgentTaskEvents).select(
    db.select({
      id: sql<string>`${crypto.randomUUID()}`.as('id'),
      taskId: excelAgentTasks.id,
      type: sql<string>`${type}`.as('type'),
      message: sql<string>`${message}`.as('message'),
      metadataJson: sql<string | null>`${metadataJson}`.as('metadata_json'),
      createdAt: sql<string>`${createdAt}`.as('created_at')
    })
      .from(excelAgentTasks)
      .where(and(
        eq(excelAgentTasks.id, taskId),
        eq(excelAgentTasks.status, status),
        eq(excelAgentTasks.updatedAt, createdAt)
      ))
  )
}

function mapWorkbookRow(row: typeof excelAgentWorkbooks.$inferSelect): ExcelAgentWorkbookRecord {
  return {
    id: row.id,
    userId: row.userId,
    sessionId: row.sessionId,
    r2Key: row.r2Key,
    fileName: row.fileName,
    sizeBytes: row.sizeBytes,
    summary: workbookSummarySchema.safeParse(decodeStoredJson(row.summaryJson)).data || emptyWorkbookSummary(),
    contextR2Key: row.contextR2Key,
    sourceEtag: row.sourceEtag,
    contextVersion: row.contextVersion,
    parsedAt: row.parsedAt,
    parentWorkbookId: row.parentWorkbookId,
    rootWorkbookId: row.rootWorkbookId,
    sourceTaskId: row.sourceTaskId,
    createdAt: row.createdAt,
    updatedAt: row.updatedAt
  }
}

function mapTaskRow(row: typeof excelAgentTasks.$inferSelect): ExcelAgentTaskRecord {
  return {
    id: row.id,
    userId: row.userId,
    sessionId: row.sessionId,
    threadId: row.threadId,
    taskType: row.taskType as ExcelAgentTaskType,
    workflowInstanceId: row.workflowInstanceId,
    inputR2Key: row.inputR2Key,
    inputName: row.inputName,
    outputR2Key: row.outputR2Key,
    outputName: row.outputName,
    prompt: row.prompt,
    plan: readTaskPlan(decodeStoredJson(row.planJson)),
    result: readTaskResult(decodeStoredJson(row.resultJson), row.taskType),
    status: row.status as ExcelAgentTaskRecord['status'],
    errorMessage: row.errorMessage,
    createdAt: row.createdAt,
    updatedAt: row.updatedAt,
    completedAt: row.completedAt
  }
}

function mapEventRow(row: typeof excelAgentTaskEvents.$inferSelect): ExcelAgentTaskEvent {
  return {
    id: row.id,
    taskId: row.taskId,
    type: row.type,
    message: row.message,
    metadata: parseJson<Record<string, unknown>>(row.metadataJson),
    createdAt: row.createdAt
  }
}

function parseJson<T>(value: string | null): T | null {
  if (!value) return null
  try {
    return JSON.parse(value) as T
  } catch {
    return null
  }
}

function emptyWorkbookSummary(): WorkbookSummary {
  return {
    sheetNames: [],
    primarySheet: '',
    rowCount: 0,
    columnCount: 0,
    headers: [],
    numericColumns: [],
    blankCellCount: 0,
    duplicateRowCount: 0,
    sheets: [],
    preview: []
  }
}
