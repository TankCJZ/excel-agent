import { and, asc, desc, eq, sql } from 'drizzle-orm'
import { drizzle } from 'drizzle-orm/d1'

import {
  excelAgentConversations,
  excelAgentMessages
} from '#shared/database/agentSchema'
import type {
  AgentTaskPlan,
  AgentWorkerEnv,
  ConversationAssistantMetadata,
  ConversationTimelineEvent,
  ExcelAgentConversationRecord,
  ExcelAgentMessageRecord,
  ExcelAgentWorkbookRecord
} from '#agent/types'
import { webSearchResultSchema, type RequestedOutcome, type WebSearchResult, type WorkbookAnalysis } from '#agent/contracts'

const MAX_METADATA_CHARS = 500_000

export class ConversationStateError extends Error {
  readonly code: 'CONVERSATION_DELETED' | 'WORKBOOK_VERSION_CONFLICT'
  constructor(code: ConversationStateError['code']) {
    super(code === 'CONVERSATION_DELETED' ? 'This conversation is being deleted. Start a new conversation.' : 'The attached workbook changed. Reload this conversation before trying again.')
    this.name = 'ConversationStateError'
    this.code = code
  }
}

export async function assertConversationWritable(env: AgentWorkerEnv, id: string) {
  const tombstone = await env.DB.prepare('SELECT conversation_id FROM excel_agent_conversation_tombstones WHERE conversation_id=?').bind(id).first()
  if (tombstone) throw new ConversationStateError('CONVERSATION_DELETED')
}

export class ConversationOwnershipError extends Error {
  constructor() {
    super('Conversation does not belong to the authenticated user or session')
    this.name = 'ConversationOwnershipError'
  }
}

function getAgentDb(env: AgentWorkerEnv) {
  return drizzle(env.DB)
}

export async function createConversation(
  env: AgentWorkerEnv,
  input: { id: string, userId: string, sessionId: string, workbookId?: string | null }
) {
  await assertConversationWritable(env, input.id)
  const db = getAgentDb(env)
  const existing = await findConversation(db, input.id)
  if (existing) {
    if (existing.userId !== input.userId || existing.sessionId !== input.sessionId) throw new ConversationOwnershipError()
    return mapConversationRow(existing)
  }

  const now = new Date().toISOString()
  await db.insert(excelAgentConversations).values({
    id: input.id,
    userId: input.userId,
    sessionId: input.sessionId,
    title: '',
    workbookId: input.workbookId || null,
    createdAt: now,
    updatedAt: now,
    lastMessageAt: null
  })
  return (await getConversationForUser(env, input.id, input.userId))!
}

export async function listConversationsForUser(env: AgentWorkerEnv, userId: string, limit: number) {
  const rows = await getAgentDb(env)
    .select()
    .from(excelAgentConversations)
    .where(eq(excelAgentConversations.userId, userId))
    .orderBy(desc(excelAgentConversations.updatedAt))
    .limit(limit)
  return rows.map(mapConversationRow)
}

export async function getConversationForUser(env: AgentWorkerEnv, conversationId: string, userId: string) {
  const [row] = await getAgentDb(env)
    .select()
    .from(excelAgentConversations)
    .where(and(
      eq(excelAgentConversations.id, conversationId),
      eq(excelAgentConversations.userId, userId)
    ))
    .limit(1)
  return row ? mapConversationRow(row) : null
}

export async function listConversationMessages(env: AgentWorkerEnv, conversationId: string) {
  const rows = await getAgentDb(env)
    .select()
    .from(excelAgentMessages)
    .where(eq(excelAgentMessages.conversationId, conversationId))
    .orderBy(asc(excelAgentMessages.createdAt), asc(excelAgentMessages.role))
    .limit(200)
  return rows.map(mapMessageRow)
}

export async function updateConversation(
  env: AgentWorkerEnv,
  input: {
    conversationId: string
    userId: string
    sessionId: string
    workbookId?: string | null
    workbookRevision?: number
    title?: string
  }
) {
  await assertConversationWritable(env, input.conversationId)
  const existing = await requireOwnedConversation(env, input.conversationId, input.userId, input.sessionId)
  const now = new Date().toISOString()
  const updated = await getAgentDb(env)
    .update(excelAgentConversations)
    .set({
      ...(input.workbookId !== undefined ? { workbookId: input.workbookId, workbookRevision: sql`${excelAgentConversations.workbookRevision} + 1` } : {}),
      ...(input.title !== undefined ? { title: input.title } : {}),
      updatedAt: now
    })
    .where(and(eq(excelAgentConversations.id, existing.id), input.workbookRevision !== undefined ? eq(excelAgentConversations.workbookRevision, input.workbookRevision) : undefined))
  if (!updated.meta.changes) throw new ConversationStateError('WORKBOOK_VERSION_CONFLICT')
  return (await getConversationForUser(env, existing.id, input.userId))!
}

export async function prepareConversationTurn(
  env: AgentWorkerEnv,
  input: {
    conversationId: string
    turnId: string
    userId: string
    sessionId: string
    message: string
    requestedOutcome?: RequestedOutcome
    workbook: ExcelAgentWorkbookRecord | null
    preserveWorkbookSelection?: boolean
    workbookRevision?: number
  }
) {
  await assertConversationWritable(env, input.conversationId)
  const db = getAgentDb(env)
  const existing = await findConversation(db, input.conversationId)
  if (existing && (existing.userId !== input.userId || existing.sessionId !== input.sessionId)) {
    throw new ConversationOwnershipError()
  }
  if (existing && input.workbookRevision !== undefined && existing.workbookRevision !== input.workbookRevision) throw new ConversationStateError('WORKBOOK_VERSION_CONFLICT')

  const now = new Date().toISOString()
  const title = conversationTitle(input.message)
  const conversationWrite = existing
    ? db.update(excelAgentConversations)
        .set({
          title: existing.title || title,
          // New clients attach through PATCH, never through a historical retry.
          ...(!input.preserveWorkbookSelection && input.workbook ? {
            workbookId: input.workbook.id,
            workbookRevision: sql`CASE WHEN ${excelAgentConversations.workbookId} IS NOT ${input.workbook.id} THEN ${excelAgentConversations.workbookRevision} + 1 ELSE ${excelAgentConversations.workbookRevision} END`
          } : {}),
          updatedAt: now,
          lastMessageAt: now
        })
        .where(eq(excelAgentConversations.id, input.conversationId))
    : db.insert(excelAgentConversations).values({
        id: input.conversationId,
        userId: input.userId,
        sessionId: input.sessionId,
        title,
        workbookId: input.workbook?.id || null,
        createdAt: now,
        updatedAt: now,
        lastMessageAt: now
      })

  const userMetadata = stringifyMetadata({
    requestedOutcome: input.requestedOutcome || null,
    workbook: input.workbook
      ? {
          workbookId: input.workbook.id,
          fileName: input.workbook.fileName,
          sizeBytes: input.workbook.sizeBytes,
          summary: input.workbook.summary
        }
      : null
  })
  const emptyAssistantMetadata = stringifyMetadata(emptyAssistantState())

  await db.batch([
    conversationWrite,
    db.insert(excelAgentMessages).values({
      id: crypto.randomUUID(),
      conversationId: input.conversationId,
      turnId: input.turnId,
      role: 'user',
      status: 'completed',
      content: input.message,
      taskId: null,
      workbookId: input.workbook?.id || null,
      metadataJson: userMetadata,
      errorCode: null,
      createdAt: now,
      updatedAt: now
    }).onConflictDoUpdate({
      target: [excelAgentMessages.conversationId, excelAgentMessages.turnId, excelAgentMessages.role],
      set: {
        content: input.message,
        workbookId: input.workbook?.id || null,
        metadataJson: userMetadata,
        status: 'completed',
        errorCode: null,
        updatedAt: now
      }
    }),
    db.insert(excelAgentMessages).values({
      id: crypto.randomUUID(),
      conversationId: input.conversationId,
      turnId: input.turnId,
      role: 'assistant',
      status: 'streaming',
      content: '',
      taskId: null,
      workbookId: input.workbook?.id || null,
      metadataJson: emptyAssistantMetadata,
      errorCode: null,
      createdAt: now,
      updatedAt: now
    }).onConflictDoUpdate({
      target: [excelAgentMessages.conversationId, excelAgentMessages.turnId, excelAgentMessages.role],
      set: {
        status: 'streaming',
        content: '',
        taskId: null,
        workbookId: input.workbook?.id || null,
        metadataJson: emptyAssistantMetadata,
        errorCode: null,
        updatedAt: now
      }
    })
  ])
}

export async function completeConversationTurn(
  env: AgentWorkerEnv,
  input: {
    conversationId: string
    turnId: string
    content: string
    taskId: string | null
    analysis: WorkbookAnalysis | null
    research: WebSearchResult | null
    followups: ConversationAssistantMetadata['followups']
    plan: AgentTaskPlan | null
    timeline: ConversationTimelineEvent[]
  }
) {
  const now = new Date().toISOString()
  const metadata = stringifyMetadata({
    analysis: input.analysis,
    research: input.research,
    followups: input.followups,
    plan: input.plan,
    timeline: input.timeline
  } satisfies ConversationAssistantMetadata)
  const db = getAgentDb(env)
  await db.batch([
    db.update(excelAgentMessages)
      .set({
        content: input.content.slice(0, 50_000),
        status: 'completed',
        taskId: input.taskId,
        metadataJson: metadata,
        errorCode: null,
        updatedAt: now
      })
      .where(and(
        eq(excelAgentMessages.conversationId, input.conversationId),
        eq(excelAgentMessages.turnId, input.turnId),
        eq(excelAgentMessages.role, 'assistant')
      )),
    db.update(excelAgentConversations)
      .set({ updatedAt: now, lastMessageAt: now })
      .where(eq(excelAgentConversations.id, input.conversationId))
  ])
}

export async function failConversationTurn(
  env: AgentWorkerEnv,
  input: {
    conversationId: string
    turnId: string
    content: string
    errorCode: string
    research: WebSearchResult | null
    timeline: ConversationTimelineEvent[]
  }
) {
  const now = new Date().toISOString()
  const metadata = stringifyMetadata({
    ...emptyAssistantState(),
    research: input.research,
    timeline: input.timeline
  } satisfies ConversationAssistantMetadata)
  const db = getAgentDb(env)
  await db.batch([
    db.update(excelAgentMessages)
      .set({
        content: input.content.slice(0, 50_000),
        status: 'failed',
        metadataJson: metadata,
        errorCode: input.errorCode,
        updatedAt: now
      })
      .where(and(
        eq(excelAgentMessages.conversationId, input.conversationId),
        eq(excelAgentMessages.turnId, input.turnId),
        eq(excelAgentMessages.role, 'assistant')
      )),
    db.update(excelAgentConversations)
      .set({ updatedAt: now, lastMessageAt: now })
      .where(eq(excelAgentConversations.id, input.conversationId))
  ])
}

export async function getReusableTurnResearch(
  env: AgentWorkerEnv,
  input: { conversationId: string, turnId: string, userId: string, sessionId: string, message: string }
) {
  await requireOwnedConversation(env, input.conversationId, input.userId, input.sessionId)
  const rows = await getAgentDb(env)
    .select({
      role: excelAgentMessages.role,
      status: excelAgentMessages.status,
      content: excelAgentMessages.content,
      metadataJson: excelAgentMessages.metadataJson
    })
    .from(excelAgentMessages)
    .where(and(
      eq(excelAgentMessages.conversationId, input.conversationId),
      eq(excelAgentMessages.turnId, input.turnId)
    ))
    .limit(2)
  const userMessage = rows.find(row => row.role === 'user')
  const assistantMessage = rows.find(row => row.role === 'assistant')
  if (userMessage?.content !== input.message || assistantMessage?.status !== 'failed') return null
  const metadata = parseJson<Record<string, unknown>>(assistantMessage.metadataJson)
  const parsed = webSearchResultSchema.safeParse(metadata?.research)
  return parsed.success ? parsed.data : null
}

async function requireOwnedConversation(
  env: AgentWorkerEnv,
  conversationId: string,
  userId: string,
  sessionId: string
) {
  const conversation = await getConversationForUser(env, conversationId, userId)
  if (!conversation || conversation.sessionId !== sessionId) throw new ConversationOwnershipError()
  return conversation
}

async function findConversation(db: ReturnType<typeof getAgentDb>, conversationId: string) {
  const [row] = await db
    .select()
    .from(excelAgentConversations)
    .where(eq(excelAgentConversations.id, conversationId))
    .limit(1)
  return row || null
}

function conversationTitle(message: string) {
  return message.replace(/\s+/g, ' ').trim().slice(0, 80)
}

function emptyAssistantState(): ConversationAssistantMetadata {
  return {
    analysis: null,
    research: null,
    followups: [],
    plan: null,
    timeline: []
  }
}

function stringifyMetadata(metadata: ConversationAssistantMetadata | Record<string, unknown>) {
  const json = JSON.stringify(metadata)
  if (json.length <= MAX_METADATA_CHARS) return json
  const withoutTimeline = JSON.stringify({ ...metadata, timeline: [] })
  if (withoutTimeline.length <= MAX_METADATA_CHARS) return withoutTimeline
  return JSON.stringify({ truncated: true })
}

function mapConversationRow(row: typeof excelAgentConversations.$inferSelect): ExcelAgentConversationRecord {
  return {
    id: row.id,
    userId: row.userId,
    sessionId: row.sessionId,
    title: row.title,
    workbookId: row.workbookId,
    workbookRevision: row.workbookRevision,
    createdAt: row.createdAt,
    updatedAt: row.updatedAt,
    lastMessageAt: row.lastMessageAt
  }
}

function mapMessageRow(row: typeof excelAgentMessages.$inferSelect): ExcelAgentMessageRecord {
  return {
    id: row.id,
    conversationId: row.conversationId,
    turnId: row.turnId,
    role: row.role as ExcelAgentMessageRecord['role'],
    status: row.status as ExcelAgentMessageRecord['status'],
    content: row.content,
    taskId: row.taskId,
    workbookId: row.workbookId,
    metadata: parseJson<Record<string, unknown>>(row.metadataJson),
    errorCode: row.errorCode,
    createdAt: row.createdAt,
    updatedAt: row.updatedAt
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
