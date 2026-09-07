import {
  createSampleWorkbookForUser,
  downloadTaskWorkbook,
  uploadWorkbookForUser
} from '#agent/http/handlers/workbooks'
import {
  authenticateAgentRequest,
  enforceAgentRateLimit,
  getAgentEnvironment,
  isLocalBillingEnabled,
  requireAgentScope
} from '#agent/platform/auth'
import {
  chatRequestSchema,
  adoptTaskWorkbookSchema,
  conversationCreateSchema,
  conversationListQuerySchema,
  conversationUpdateSchema,
  evalRequestSchema,
  taskIdSchema,
  taskQuerySchema,
  threadIdSchema,
  threadQuerySchema
} from '#agent/contracts'
import {
  checkAgentPersistence,
  getTaskForSession,
  getThreadTaskHistory,
  getWorkbookForSession,
  listTasksForThread
} from '#agent/persistence/agentRepository'
import {
  ConversationOwnershipError,
  createConversation,
  getReusableTurnResearch,
  getConversationForUser,
  listConversationMessages,
  listConversationsForUser,
  updateConversation,
  assertConversationWritable,
  ConversationStateError
} from '#agent/persistence/conversationRepository'
import {
  ConversationDeletionBusyError,
  ConversationDeletionStageError,
  deleteConversationForUser
} from '#agent/persistence/conversationDeletion'
import { createAgentStream } from '#agent/http/agentStream'
import { corsHeaders, errorResponse, HttpError, jsonResponse } from '#agent/http/responses'
import { runArchitectureEvals } from '#agent/mastra/evals/architecture'
import { DEFAULT_MODEL_ID } from '#agent/mastra/models/cloudflareResponses'
import { resolveAgentAnswerTimeoutMs, resolveAgentRequestTimeoutMs } from '#agent/platform/requestTimeout'
import type { AgentWorkerEnv } from '#agent/types'
import {
  assertConcurrentTaskCapacity,
  getBillingAccess,
  getTurnCredits,
  getTaskCredits,
  listConversationCredits,
  releaseExpiredReservationsForUser
} from '#agent/billing/billingRepository'
import type { CreditTurnSummary } from '#shared/billing/creditSummary'
import type { ExcelAgentRuntimeDependencies } from '#agent/mastra/runtime'
import { reconcileWorkbookMutations } from '#agent/workflows/mutationRecovery'
import { adoptTaskWorkbook } from '#agent/persistence/adoptTaskWorkbook'
import { listWorkbookHistory, previewHistoryWorkbook, selectHistoryWorkbook } from '#agent/persistence/workbookHistory'
import { z } from 'zod'
import { getLocalizedFollowups } from '#agent/persistence/followupPresentation'

export async function handleAgentRequest(
  request: Request, env: AgentWorkerEnv, context?: ExecutionContext,
  runtimeDependencies: ExcelAgentRuntimeDependencies = {}
) {
  const requestId = request.headers.get('cf-ray') || crypto.randomUUID()

  try {
    if (request.method === 'OPTIONS') {
      return new Response(null, { status: 204, headers: corsHeaders(request, env) })
    }

    const url = new URL(request.url)
    if (request.method === 'GET' && url.pathname === '/') {
      return jsonResponse(request, env, requestId, {
        service: 'excelgen-mastra-agent',
        endpoints: ['/api/health', '/api/files/sample', '/api/files/upload', '/api/chat', '/api/conversations', '/api/conversations/:id', '/api/tasks/:id', '/api/threads/:id/tasks', '/api/evals/run']
      })
    }
    if (request.method === 'GET' && url.pathname === '/api/health') {
      return jsonResponse(request, env, requestId, await checkHealth(env))
    }

    const identity = await authenticateAgentRequest(request, env)
    const turnCreditsMatch = url.pathname.match(/^\/api\/credits\/turns\/([^/]+)$/)
    if (request.method === 'GET' && turnCreditsMatch) {
      const turnId = taskIdSchema.parse(turnCreditsMatch[1])
      if (identity.billingEnabled) await releaseExpiredReservationsForUser(env, identity.userId)
      return jsonResponse(request, env, requestId, {
        credit: identity.billingEnabled ? await getTurnCredits(env, identity.userId, turnId) : null
      })
    }
    if (request.method === 'POST' && url.pathname === '/api/files/sample') {
      await enforceAgentRateLimit(env, identity, 'files', 20, 60 * 60)
      return jsonResponse(request, env, requestId, await createSampleWorkbookForUser(request, env, identity), 201)
    }
    if (request.method === 'POST' && url.pathname === '/api/files/upload') {
      await enforceAgentRateLimit(env, identity, 'files', 20, 60 * 60)
      return jsonResponse(request, env, requestId, await uploadWorkbookForUser(request, env, identity), 201)
    }
    if (request.method === 'POST' && url.pathname === '/api/chat') {
      await enforceAgentRateLimit(env, identity, 'chat', 60, 60 * 60)
      if (identity.billingEnabled) {
        const access = await getBillingAccess(env, identity.userId)
        await assertConcurrentTaskCapacity(env, identity.userId, access)
      }
      const body = chatRequestSchema.parse(await request.json())
      await withConversationAccess(() => assertConversationWritable(env, body.threadId))
      const currentConversation = await getConversationForUser(env, body.threadId, identity.userId)
      if (body.workbookMutationV1 && currentConversation && !body.retryOfTurnId && (
        (body.workbookId || null) !== currentConversation.workbookId ||
        (body.workbookRevision !== undefined && body.workbookRevision !== currentConversation.workbookRevision)
      )) throw new HttpError(409, 'WORKBOOK_VERSION_CONFLICT', 'The attached workbook changed. Reload this conversation before trying again')
      const workbook = body.workbookId
        ? await getWorkbookForSession(env, body.workbookId, identity.userId, body.sessionId)
        : null
      if (body.workbookId && !workbook) throw new HttpError(404, 'NOT_FOUND', 'Workbook was not found')
      if (workbook && !await env.WORKBOOKS.head(workbook.r2Key)) {
        throw new HttpError(410, 'WORKBOOK_FILE_EXPIRED', 'This workbook is no longer available. Upload your saved copy to continue')
      }
      const reusableResearch = body.retryOfTurnId
        ? await withConversationAccess(() => getReusableTurnResearch(env, {
            conversationId: body.threadId,
            turnId: body.retryOfTurnId!,
            userId: identity.userId,
            sessionId: body.sessionId,
            message: body.message
          }))
        : null
      return createAgentStream(request, env, identity, body, workbook, requestId, reusableResearch, context, runtimeDependencies)
    }
    if (request.method === 'GET' && url.pathname === '/api/conversations') {
      const query = conversationListQuerySchema.parse({ limit: url.searchParams.get('limit') || undefined })
      const conversations = await listConversationsForUser(env, identity.userId, query.limit)
      return jsonResponse(request, env, requestId, {
        conversations: conversations.map(toPublicConversation)
      })
    }
    if (request.method === 'POST' && url.pathname === '/api/conversations') {
      const body = conversationCreateSchema.parse(await request.json())
      if (body.workbookId) {
        const workbook = await getWorkbookForSession(env, body.workbookId, identity.userId, body.sessionId)
        if (!workbook) throw new HttpError(404, 'NOT_FOUND', 'Workbook was not found')
      }
      const conversation = await withConversationAccess(() => createConversation(env, {
        id: body.id,
        userId: identity.userId,
        sessionId: body.sessionId,
        workbookId: body.workbookId
      }))
      return jsonResponse(request, env, requestId, { conversation: toPublicConversation(conversation) }, 201)
    }
    if (request.method === 'POST' && url.pathname === '/api/evals/run') {
      requireAgentScope(identity, 'agent:eval')
      await enforceAgentRateLimit(env, identity, 'evals', 3, 60 * 60)
      const body = evalRequestSchema.parse(await request.json())
      return jsonResponse(request, env, requestId, await runArchitectureEvals(env, body.sessionId))
    }

    const downloadMatch = url.pathname.match(/^\/api\/tasks\/([^/]+)\/download$/)
    if (request.method === 'GET' && downloadMatch) {
      return await downloadTaskWorkbook(request, env, identity, downloadMatch[1], url.searchParams.get('sessionId'))
    }
    const adoptMatch = url.pathname.match(/^\/api\/tasks\/([^/]+)\/use-workbook$/)
    if (request.method === 'POST' && adoptMatch) {
      await enforceAgentRateLimit(env, identity, 'files', 20, 60 * 60)
      const body = adoptTaskWorkbookSchema.parse(await request.json())
      const result = await withConversationAccess(() => adoptTaskWorkbook(env, { ...body, taskId: taskIdSchema.parse(adoptMatch[1]), userId: identity.userId }))
      return jsonResponse(request, env, requestId, { workbook: result.workbook, conversation: toPublicConversation(result.conversation) })
    }
    const historyMatch = url.pathname.match(/^\/api\/conversations\/([^/]+)\/workbooks(?:\/([^/]+))?$/)
    if (historyMatch && ['GET', 'POST'].includes(request.method)) {
      const conversationId = threadIdSchema.parse(historyMatch[1])
      const entryId = historyMatch[2] ? decodeURIComponent(historyMatch[2]) : null
      if (request.method === 'GET') {
        const result = entryId
          ? await withConversationAccess(() => previewHistoryWorkbook(env, conversationId, identity.userId, entryId))
          : await withConversationAccess(() => listWorkbookHistory(env, conversationId, identity.userId, z.coerce.number().int().min(0).max(100000).parse(url.searchParams.get('offset') || 0)))
        return jsonResponse(request, env, requestId, result)
      }
      if (!entryId) throw new HttpError(400, 'VALIDATION_ERROR', 'Choose a workbook version')
      const body = z.object({ workbookRevision: z.number().int().nonnegative() }).strict().parse(await request.json())
      const result = await withConversationAccess(() => selectHistoryWorkbook(env, { conversationId, userId: identity.userId, entryId, workbookRevision: body.workbookRevision }))
      return jsonResponse(request, env, requestId, { conversation: toPublicConversation(result.conversation), workbook: result.workbook })
    }
    const followupsMatch = url.pathname.match(/^\/api\/conversations\/([^/]+)\/turns\/([^/]+)\/followups$/)
    if (followupsMatch && request.method === 'POST') {
      requireAgentScope(identity, 'agent:use')
      const presentation = await getLocalizedFollowups(env, {
        userId: identity.userId,
        conversationId: threadIdSchema.parse(followupsMatch[1]),
        turnId: taskIdSchema.parse(followupsMatch[2]),
        signal: request.signal
      }, {
        generate: runtimeDependencies.followupCopyGenerator,
        beforeGenerate: () => enforceAgentRateLimit(env, identity, 'followup-copy', 20, 60)
      })
      return jsonResponse(request, env, requestId, { presentation })
    }
    const conversationMatch = url.pathname.match(/^\/api\/conversations\/([^/]+)$/)
    if (conversationMatch && request.method === 'GET') {
      const conversationId = threadIdSchema.parse(conversationMatch[1])
      const conversation = await getConversationForUser(env, conversationId, identity.userId)
      if (!conversation) throw new HttpError(404, 'NOT_FOUND', 'Conversation was not found')
      if (identity.billingEnabled) await releaseExpiredReservationsForUser(env, identity.userId)
      const [messages, history, workbook, credits] = await Promise.all([
        listConversationMessages(env, conversation.id),
        getThreadTaskHistory(env, conversation.id, identity.userId, conversation.sessionId),
        conversation.workbookId
          ? getWorkbookForSession(env, conversation.workbookId, identity.userId, conversation.sessionId)
          : null,
        identity.billingEnabled ? listConversationCredits(env, identity.userId, conversation.id) : []
      ])
      return jsonResponse(request, env, requestId, {
        conversation: toPublicConversation(conversation),
        workbook: workbook ? toPublicWorkbook(workbook) : null,
        turns: toPublicConversationTurns(messages, history.tasks, history.events, credits)
      })
    }
    if (conversationMatch && request.method === 'PATCH') {
      const conversationId = threadIdSchema.parse(conversationMatch[1])
      const body = conversationUpdateSchema.parse(await request.json())
      if (body.workbookId) {
        const workbook = await getWorkbookForSession(env, body.workbookId, identity.userId, body.sessionId)
        if (!workbook) throw new HttpError(404, 'NOT_FOUND', 'Workbook was not found')
      }
      const conversation = await withConversationAccess(() => updateConversation(env, {
        conversationId,
        userId: identity.userId,
        sessionId: body.sessionId,
        ...(body.workbookId !== undefined ? { workbookId: body.workbookId } : {}),
        workbookRevision: body.workbookRevision,
        ...(body.title !== undefined ? { title: body.title } : {})
      }))
      return jsonResponse(request, env, requestId, { conversation: toPublicConversation(conversation) })
    }
    if (conversationMatch && request.method === 'DELETE') {
      const conversationId = threadIdSchema.parse(conversationMatch[1])
      try {
        const deleted = await deleteConversationForUser(env, conversationId, identity.userId)
        if (!deleted) throw new HttpError(404, 'NOT_FOUND', 'Conversation was not found')
        return jsonResponse(request, env, requestId, { deleted })
      } catch (error) {
        if (error instanceof ConversationDeletionBusyError) {
          throw new HttpError(409, 'CONVERSATION_BUSY', 'Wait for the current task to stop, then try deleting the conversation again')
        }
        if (error instanceof ConversationDeletionStageError) {
          throw new HttpError(503, 'CONVERSATION_DELETE_FAILED', 'The conversation could not be deleted. Try again')
        }
        throw error
      }
    }
    const threadTasksMatch = url.pathname.match(/^\/api\/threads\/([^/]+)\/tasks$/)
    if (request.method === 'GET' && threadTasksMatch) {
      const threadId = threadIdSchema.parse(threadTasksMatch[1])
      const query = threadQuerySchema.parse({ sessionId: url.searchParams.get('sessionId') })
      const tasks = await listTasksForThread(env, threadId, identity.userId, query.sessionId)
      return jsonResponse(request, env, requestId, { tasks: tasks.map(toPublicTask) })
    }
    const taskMatch = url.pathname.match(/^\/api\/tasks\/([^/]+)$/)
    if (request.method === 'GET' && taskMatch) {
      const taskId = taskIdSchema.parse(taskMatch[1])
      const query = taskQuerySchema.parse({ sessionId: url.searchParams.get('sessionId') })
      let record = await getTaskForSession(env, taskId, identity.userId, query.sessionId)
      if (!record) throw new HttpError(404, 'NOT_FOUND', 'Task was not found')
      if (record.task.taskType === 'workbook_mutation' && ['queued', 'processing'].includes(record.task.status)) {
        const recovery = reconcileWorkbookMutations(env, taskId)
        if (context) context.waitUntil(recovery)
        else {
          await recovery
          record = await getTaskForSession(env, taskId, identity.userId, query.sessionId)
          if (!record) throw new HttpError(404, 'NOT_FOUND', 'Task was not found')
        }
      }
      return jsonResponse(request, env, requestId, {
        task: toPublicTask(record.task),
        credit: identity.billingEnabled ? await getTaskCredits(env, identity.userId, taskId) : null,
        events: record.events.map(event => ({
          id: event.id,
          type: event.type,
          metadata: event.metadata,
          createdAt: event.createdAt
        }))
      })
    }

    throw new HttpError(404, 'NOT_FOUND', 'Endpoint was not found')
  } catch (error) {
    return errorResponse(request, env, requestId, error)
  }
}

async function checkHealth(env: AgentWorkerEnv) {
  const [persistence, objects] = await Promise.all([
    checkAgentPersistence(env),
    env.WORKBOOKS.list({ prefix: 'agent-demo/', limit: 1 })
  ])
  const model = env.MASTRA_MODEL || DEFAULT_MODEL_ID
  return {
    ok: true,
    model,
    mode: model.startsWith('mock/') ? 'mock' : 'provider',
    environment: getAgentEnvironment(env),
    requestTimeoutMs: resolveAgentRequestTimeoutMs(env.AGENT_REQUEST_TIMEOUT_MS),
    answerTimeoutMs: resolveAgentAnswerTimeoutMs(env.AGENT_ANSWER_TIMEOUT_MS),
    bindings: {
      authentication: Boolean(env.AGENT_TOKEN_SECRET),
      responsesApi: Boolean(env.CLOUDFLARE_ACCOUNT_ID && env.CLOUDFLARE_API_TOKEN),
      localBillingEnabled: isLocalBillingEnabled(env),
      d1: persistence,
      workbookRegistry: persistence,
      mastraMemory: Boolean(env.DB),
      r2: Boolean(objects),
      workflows: Boolean(env.EXCEL_AGENT_WORKFLOW && env.SPREADSHEET_GENERATION_WORKFLOW)
    },
    packages: {
      mastra: '1.61.0',
      mastraD1: '1.3.0',
      sheetjs: '0.20.3'
    }
  }
}

function toPublicTask(task: Awaited<ReturnType<typeof listTasksForThread>>[number]) {
  const {
    inputR2Key: _inputR2Key,
    outputR2Key: _outputR2Key,
    userId: _userId,
    sessionId: _sessionId,
    ...publicTask
  } = task
  if (!publicTask.result) return publicTask
  if (publicTask.result.kind === 'workbook_mutation') {
    const { outputR2Key: _resultR2Key, contextR2Key: _contextKey, sourceEtag: _etag, ...publicResult } = publicTask.result
    return { ...publicTask, result: publicResult }
  }
  const { outputR2Key: _resultR2Key, ...publicResult } = publicTask.result
  return { ...publicTask, result: publicResult }
}

function toPublicConversation(conversation: Awaited<ReturnType<typeof createConversation>>) {
  const { userId: _userId, ...publicConversation } = conversation
  return publicConversation
}

function toPublicWorkbook(workbook: NonNullable<Awaited<ReturnType<typeof getWorkbookForSession>>>) {
  return {
    workbookId: workbook.id,
    fileName: workbook.fileName,
    sizeBytes: workbook.sizeBytes,
    summary: workbook.summary
  }
}

function toPublicConversationTurns(
  messages: Awaited<ReturnType<typeof listConversationMessages>>,
  tasks: Awaited<ReturnType<typeof listTasksForThread>>,
  events: Awaited<ReturnType<typeof getThreadTaskHistory>>['events'],
  credits: CreditTurnSummary[] = []
) {
  const turns = new Map<string, {
    turnId: string
    user: (typeof messages)[number] | null
    assistant: (typeof messages)[number] | null
  }>()
  for (const message of messages) {
    const turn = turns.get(message.turnId) || { turnId: message.turnId, user: null, assistant: null }
    if (message.role === 'user') turn.user = message
    else turn.assistant = message
    turns.set(message.turnId, turn)
  }
  const tasksById = new Map(tasks.map(task => [task.id, task]))
  const creditsByTurn = new Map(credits.map(credit => [credit.turnId, credit]))
  const eventsByTask = new Map<string, typeof events>()
  for (const event of events) {
    const list = eventsByTask.get(event.taskId) || []
    list.push(event)
    eventsByTask.set(event.taskId, list)
  }
  return [...turns.values()]
    .filter(turn => turn.user)
    .map(turn => {
      const task = turn.assistant?.taskId ? tasksById.get(turn.assistant.taskId) : undefined
      return {
        turnId: turn.turnId,
        credit: creditsByTurn.get(turn.turnId) || null,
        user: toPublicConversationMessage(turn.user!),
        assistant: turn.assistant ? toPublicConversationMessage(turn.assistant) : null,
        task: task ? toPublicTask(task) : null,
        events: task
          ? (eventsByTask.get(task.id) || []).map(event => ({
              id: event.id,
              type: event.type,
              metadata: event.metadata,
              createdAt: event.createdAt
            }))
          : []
      }
    })
}

function toPublicConversationMessage(message: Awaited<ReturnType<typeof listConversationMessages>>[number]) {
  const { conversationId: _conversationId, ...publicMessage } = message
  return publicMessage
}

async function withConversationAccess<T>(operation: () => Promise<T>) {
  try {
    return await operation()
  } catch (error) {
    if (error instanceof ConversationStateError) throw new HttpError(409, error.code, error.message)
    if (error instanceof ConversationOwnershipError) {
      throw new HttpError(404, 'NOT_FOUND', 'Conversation was not found')
    }
    throw error
  }
}
