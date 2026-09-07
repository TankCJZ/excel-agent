import { createUIMessageStreamResponse, type UIMessageChunk } from 'ai'
import { toAISdkStream, withSseHeartbeat } from '@mastra/ai-sdk'
import type { AgentIdentity } from '#agent/platform/auth'
import type { ChatRequest, WebSearchResult } from '#agent/contracts'
import {
  attachReservationTask,
  ensureTurnCreditReservation,
  releaseTurnCredits,
  reserveTurnCredits,
  settleTurnCredits,
  type CreditReservation
} from '#agent/billing/billingRepository'
import {
  buildCreditUsageActions,
  calculateActualTurnCredits,
  estimateTurnCredits
} from '#shared/billing/planCatalog'
import {
  createExcelAgentTask,
  failTask,
  saveTaskPlan,
  saveWorkflowInstance
} from '#agent/persistence/agentRepository'
import {
  completeConversationTurn,
  failConversationTurn,
  prepareConversationTurn
} from '#agent/persistence/conversationRepository'
import { getConversationForUser } from '#agent/persistence/conversationRepository'
import { createMutationTask, stableMutationId } from '#agent/persistence/workbookMutations'
import { dispatchWorkbookMutation } from '#agent/workflows/mutationDispatch'
import { corsHeaders, HttpError, internalErrorMessage, publicStreamError } from '#agent/http/responses'
import { toolOutputErrorMessage } from '#agent/http/toolOutput'
import { createWorkbookFollowups } from '#agent/mastra/followups'
import { DEFAULT_MODEL_ID } from '#agent/mastra/models/cloudflareResponses'
import { createExcelAgentRuntime, type ExcelAgentRuntimeDependencies } from '#agent/mastra/runtime'
import {
  resolveAgentAnswerTimeoutMs,
  resolveAgentRequestTimeoutMs
} from '#agent/platform/requestTimeout'
import { researchSnapshotR2Key } from '#agent/research/webSearch'
import type {
  AgentWorkerEnv,
  ConversationAssistantMetadata,
  ConversationTimelineEvent,
  ExcelAgentWorkbookRecord
} from '#agent/types'
import type { ExcelGenAgentDataParts, ExcelGenAgentMessage } from '#shared/agent/stream'

export function createAgentStream(
  request: Request,
  env: AgentWorkerEnv,
  identity: AgentIdentity,
  body: ChatRequest,
  workbook: ExcelAgentWorkbookRecord | null,
  requestId: string,
  initialResearch: WebSearchResult | null = null,
  executionContext?: Pick<ExecutionContext, 'waitUntil'>,
  runtimeDependencies: ExcelAgentRuntimeDependencies = {}
) {
  type AgentPhase = 'thinking' | 'synthesizing' | 'answering'

  const runAbortController = new AbortController()
  let closed = false
  let streamFinished = false
  const abortRun = (reason?: unknown) => {
    if (!runAbortController.signal.aborted) runAbortController.abort(reason)
  }
  const abortFromRequest = () => abortRun(request.signal.reason)
  if (request.signal.aborted) abortFromRequest()
  else request.signal.addEventListener('abort', abortFromRequest, { once: true })

  const stream = new ReadableStream<UIMessageChunk<ExcelGenAgentMessage['metadata'], ExcelGenAgentDataParts>>({
    start(controller) {
      controller.enqueue({
        type: 'start',
        messageId: `assistant:${body.turnId}`,
        messageMetadata: { requestId, turnId: body.turnId }
      })
      const completion = (async () => {
        let taskId: string | null = null
        let requestTimeout: ReturnType<typeof setTimeout> | null = null
        let answerIdleTimeout: ReturnType<typeof setTimeout> | null = null
        let timedOut = false
        let timeoutErrorCode = ''
        let turnPrepared = false
        let turnFinalized = false
        let persistedAnswer = ''
        let timeoutPersistence: Promise<void> | null = null
        let creditReservation: CreditReservation | null = null
        let creditDelegatedToWorkflow = false
        let creditFinalized = false
        let sourceWorkbookRevision = -1
        let generationSelection: import('#agent/types').SpreadsheetWorkflowPayload['workbookSelection']
        const completedBillableTools = new Set<string>()
        const timeline: ConversationTimelineEvent[] = []
        let runtime: ReturnType<typeof createExcelAgentRuntime> | null = null
        const send = (event: ExcelGenAgentDataParts['agent']['event'], data: Record<string, unknown>) => {
          if (closed || runAbortController.signal.aborted) return false
          try {
            controller.enqueue({
              type: 'data-agent',
              ...streamPartId(event, data),
              data: { event, data }
            })
            if ((event === 'phase' || event === 'tool' || event === 'tools') && timeline.length < 100) {
              timeline.push({ event, data })
            }
            return true
          } catch {
            closed = true
            abortRun('UI message response was closed')
            return false
          }
        }
        const finishStream = (finishReason: 'stop' | 'error' = 'stop') => {
          if (closed || streamFinished) return
          streamFinished = true
          try {
            controller.enqueue({ type: 'finish', finishReason })
          } catch {
            closed = true
          }
        }
        const phaseStartedAt = new Map<AgentPhase, number>()
        const startPhase = (phase: AgentPhase) => {
          if (phaseStartedAt.has(phase)) return
          phaseStartedAt.set(phase, Date.now())
          send('phase', { phase, status: 'started' })
        }
        const finishPhase = (phase: AgentPhase, failed = false) => {
          const startedAt = phaseStartedAt.get(phase)
          if (!startedAt) return
          phaseStartedAt.delete(phase)
          send('phase', {
            phase,
            status: failed ? 'failed' : 'completed',
            durationMs: Math.max(0, Date.now() - startedAt)
          })
        }

        const persistFailure = async (errorCode: string) => {
          if (!turnPrepared || turnFinalized) return
          turnFinalized = true
          try {
            await failConversationTurn(env, {
              conversationId: body.threadId,
              turnId: body.turnId,
              content: persistedAnswer,
              errorCode,
              research: runtime?.getResearch() || initialResearch,
              timeline
            })
          } catch (persistError) {
            console.error('Conversation failure could not be persisted', {
              requestId,
              error: internalErrorMessage(persistError)
            })
          }
        }

        const clearAnswerIdleTimeout = () => {
          if (!answerIdleTimeout) return
          clearTimeout(answerIdleTimeout)
          answerIdleTimeout = null
        }
        const triggerTimeout = (errorCode: 'AGENT_TIMEOUT' | 'AGENT_ANSWER_TIMEOUT', message: string) => {
          if (closed || timedOut || runAbortController.signal.aborted) return
          timedOut = true
          timeoutErrorCode = errorCode
          clearAnswerIdleTimeout()
          finishPhase('thinking', true)
          finishPhase('synthesizing', true)
          finishPhase('answering', true)
          send('error', { code: errorCode, message, taskId })
          timeoutPersistence = persistFailure(errorCode)
          finishStream('error')
          closed = true
          try {
            controller.close()
          } catch {
            // The client may have closed the response at the same time as the timeout.
          }
          abortRun(errorCode === 'AGENT_ANSWER_TIMEOUT' ? 'Agent answer timed out' : 'Agent request timed out')
        }
        const armAnswerIdleTimeout = () => {
          clearAnswerIdleTimeout()
          if (closed || timedOut || runAbortController.signal.aborted) return
          answerIdleTimeout = setTimeout(() => {
            triggerTimeout('AGENT_ANSWER_TIMEOUT', 'The AI model stopped responding while generating the answer')
          }, resolveAgentAnswerTimeoutMs(env.AGENT_ANSWER_TIMEOUT_MS))
        }

        try {
          if (identity.billingEnabled) {
            creditReservation = await reserveTurnCredits(env, {
              userId: identity.userId,
              conversationId: body.threadId,
              turnId: body.turnId,
              rejectExisting: true,
              estimatedCredits: estimateTurnCredits(body.message, Boolean(workbook), body.requestedOutcome, {
                reuseWebSearch: Boolean(initialResearch)
              })
            })
            send('credit.estimated', {
              reservationId: creditReservation.id,
              turnId: body.turnId,
              status: 'reserved',
              refundedCredits: 0,
              estimatedCredits: creditReservation.estimatedCredits,
              balance: creditReservation.balanceAfterHold
            })
          }
          await prepareConversationTurn(env, {
            conversationId: body.threadId,
            turnId: body.turnId,
            userId: identity.userId,
            sessionId: body.sessionId,
            message: body.message,
            requestedOutcome: body.requestedOutcome,
            preserveWorkbookSelection: body.workbookMutationV1 === true,
            workbookRevision: body.workbookRevision,
            workbook
          })
          turnPrepared = true
          if (workbook && body.workbookMutationV1) {
            const current = await getConversationForUser(env, body.threadId, identity.userId)
            // Do not adopt a newer revision merely because the user switched
            // away and back to the same file while this request was starting.
            if (current?.workbookId === workbook.id && (body.workbookRevision === undefined || current.workbookRevision === body.workbookRevision)) sourceWorkbookRevision = current.workbookRevision || 0
          }
          if (body.workbookVersionsV1 && !body.retryOfTurnId) {
            const current = await getConversationForUser(env, body.threadId, identity.userId)
            if (current && current.workbookId === (workbook?.id || null) && current.workbookRevision === body.workbookRevision) {
              generationSelection = { threadId: body.threadId, workbookId: current.workbookId, revision: current.workbookRevision || 0 }
            }
          }
          send('connected', {
            requestId,
            model: env.MASTRA_MODEL || DEFAULT_MODEL_ID,
            webSearchModel: env.WEB_SEARCH_MODEL || 'openai/gpt-4o-mini',
            workbookAttached: Boolean(workbook)
          })
          startPhase('thinking')
          requestTimeout = setTimeout(() => {
            triggerTimeout('AGENT_TIMEOUT', 'The Agent did not complete before the request deadline')
          }, resolveAgentRequestTimeoutMs(env.AGENT_REQUEST_TIMEOUT_MS))
          const initialResearchSnapshotR2Key = initialResearch
            ? researchSnapshotR2Key(body.sessionId, initialResearch.researchId)
            : null
          const activeRuntime = createExcelAgentRuntime(env, workbook
            ? {
                workbookId: workbook.id,
                fileKey: workbook.r2Key,
                fileName: workbook.fileName,
                supportsWorkbookMutation: body.workbookMutationV1 === true,
                contextR2Key: workbook.contextR2Key,
                contextVersion: workbook.contextVersion,
                sourceEtag: workbook.sourceEtag,
                summary: workbook.summary,
                userId: identity.userId,
                sessionId: body.sessionId,
                billingEnabled: identity.billingEnabled,
                initialResearch,
                initialResearchSnapshotR2Key,
                requestedOutcome: body.requestedOutcome
              }
            : {
                userId: identity.userId,
                sessionId: body.sessionId,
                billingEnabled: identity.billingEnabled,
                initialResearch,
                initialResearchSnapshotR2Key,
                requestedOutcome: body.requestedOutcome
              }, runtimeDependencies)
          runtime = activeRuntime
          const agentOutput = await activeRuntime.agent.stream(body.message, {
            maxSteps: 10,
            abortSignal: runAbortController.signal,
            persistPartialOnAbort: false,
            memory: {
              thread: body.threadId,
              resource: body.sessionId
            }
          })
          armAnswerIdleTimeout()

          const toolNames = new Set<string>()
          const toolStartedAt = new Map<string, number>()
          const activeToolCalls = new Set<string>()
          const finishedToolCalls = new Set<string>()
          const failedToolNames = new Set<string>()
          const toolNameByCallId = new Map<string, string>()
          let modelText = ''
          let activeTextSegment = ''
          let forwardedModelText = false

          const persistTextSegment = () => {
            if (activeTextSegment && timeline.length < 100) {
              timeline.push({ event: 'text', data: { text: activeTextSegment.slice(0, 12_000) } })
            }
            activeTextSegment = ''
          }

          const startTool = (toolCallId: string, toolName: string) => {
            toolNames.add(toolName)
            toolNameByCallId.set(toolCallId, toolName)
            if (toolStartedAt.has(toolCallId)) return
            clearAnswerIdleTimeout()
            finishPhase('thinking')
            finishPhase('synthesizing')
            finishPhase('answering')
            toolStartedAt.set(toolCallId, Date.now())
            activeToolCalls.add(toolCallId)
            send('tool', { toolCallId, toolName, status: 'started' })
          }

          const finishTool = (toolCallId: string, toolName: string, failed: boolean) => {
            if (finishedToolCalls.has(toolCallId)) return
            finishedToolCalls.add(toolCallId)
            activeToolCalls.delete(toolCallId)
            toolNames.add(toolName)
            if (failed) failedToolNames.add(toolName)
            else {
              failedToolNames.delete(toolName)
              const billableToolName = toolName === 'createDataChart' ? 'createChart' : toolName
              if (billableToolName !== 'editWorkbook' && (billableToolName !== 'searchWeb' || activeRuntime.didUseExternalWebSearch())) {
                completedBillableTools.add(billableToolName)
              }
            }
            const startedAt = toolStartedAt.get(toolCallId)
            const summary = activeRuntime.getSummary()
            const analysis = activeRuntime.getAnalysis()
            const evidence = analysis?.evidence[0]
            const sheet = evidence?.sheet || summary?.primarySheet
            const sheetSummary = summary?.sheets.find(item => item.name === sheet) || summary?.sheets[0]
            const runtimeDetail = activeRuntime.getToolDetail(toolName)
            const detail = runtimeDetail || (toolName === 'inspectWorkbook' && summary
              ? {
                  sheet: sheetSummary?.name || summary?.primarySheet,
                  range: sheetSummary?.sourceRange,
                  rowCount: sheetSummary?.rowCount ?? summary?.rowCount,
                  columnCount: sheetSummary?.columnCount ?? summary?.columnCount
                }
              : evidence
                ? {
                    sheet: evidence.sheet,
                    range: evidence.range,
                    rowCount: sheetSummary?.rowCount,
                    columnCount: sheetSummary?.columnCount
                  }
                : undefined)
            send('tool', {
              toolCallId,
              toolName,
              status: failed ? 'failed' : 'completed',
              ...(startedAt ? { durationMs: Math.max(0, Date.now() - startedAt) } : {}),
              ...(detail ? { detail } : {})
            })
            if (activeToolCalls.size === 0) {
              startPhase('synthesizing')
              armAnswerIdleTimeout()
            }
            const billingError = activeRuntime.getBillingError()
            if (billingError) throw billingError
          }

          const agentUiStream = toAISdkStream(agentOutput, {
            from: 'agent',
            version: 'v7',
            sendStart: false,
            sendFinish: false,
            sendReasoning: false,
            sendSources: false,
            onError: internalErrorMessage
          })

          const agentUiReader = agentUiStream.getReader()
          try {
            while (true) {
              const { done, value: chunk } = await agentUiReader.read()
              if (done || runAbortController.signal.aborted) break
              if (activeToolCalls.size === 0 && ['text-start', 'text-delta', 'text-end'].includes(chunk.type)) {
                armAnswerIdleTimeout()
              }
              if (chunk.type === 'tool-input-start' || chunk.type === 'tool-input-available') {
                startTool(chunk.toolCallId, chunk.toolName)
              } else if (chunk.type === 'tool-input-error') {
                console.error('Agent tool failed', {
                  toolName: chunk.toolName,
                  error: chunk.errorText
                })
                finishTool(chunk.toolCallId, chunk.toolName, true)
              } else if (chunk.type === 'tool-output-available') {
                const toolName = toolNameByCallId.get(chunk.toolCallId)
                const errorText = toolOutputErrorMessage(chunk.output)
                if (toolName && errorText) {
                  console.error('Agent tool failed', { toolName, error: errorText })
                  finishTool(chunk.toolCallId, toolName, true)
                } else if (toolName) {
                  finishTool(chunk.toolCallId, toolName, false)
                }
              } else if (chunk.type === 'tool-output-error') {
                const toolName = toolNameByCallId.get(chunk.toolCallId) || 'unknown'
                console.error('Agent tool failed', { toolName, error: chunk.errorText })
                finishTool(chunk.toolCallId, toolName, true)
              } else if (chunk.type === 'text-delta') {
                modelText += chunk.delta
                activeTextSegment += chunk.delta
                persistedAnswer = modelText.slice(0, 50_000)
                if (chunk.delta) {
                  forwardedModelText = true
                  finishPhase('thinking')
                  finishPhase('synthesizing')
                  startPhase('answering')
                  controller.enqueue(chunk)
                }
              } else if (chunk.type === 'text-start') {
                persistTextSegment()
                controller.enqueue(chunk)
              } else if (chunk.type === 'start-step' || chunk.type === 'finish-step') {
                controller.enqueue(chunk)
              } else if (chunk.type === 'text-end') {
                persistTextSegment()
                controller.enqueue(chunk)
              } else if (chunk.type === 'error') {
                throw new Error(chunk.errorText)
              } else if (chunk.type === 'abort') {
                abortRun('Mastra stream aborted')
              }
            }
          } finally {
            agentUiReader.releaseLock()
          }
          persistTextSegment()

          if (runAbortController.signal.aborted) return

          const billingError = activeRuntime.getBillingError()
          if (billingError) throw billingError

          if (failedToolNames.has('createSpreadsheet') && !activeRuntime.getSpreadsheetPlan()) {
            throw new HttpError(422, 'SPREADSHEET_PLAN_INVALID', 'The spreadsheet plan could not be validated')
          }

          const analysis = activeRuntime.getAnalysis()
          const resolvedModelText = modelText || await agentOutput.text
          clearAnswerIdleTimeout()
          if (requestTimeout) {
            clearTimeout(requestTimeout)
            requestTimeout = null
          }
          const finalText = normalizeAgentText(resolvedModelText)
          if (!finalText) {
            throw new HttpError(502, 'EMPTY_AGENT_RESPONSE', 'The model completed without a user-facing response')
          }
          const research = activeRuntime.getResearch()
          persistedAnswer = finalText

          finishPhase('thinking')
          finishPhase('synthesizing')
          startPhase('answering')
          if (!forwardedModelText) {
            const textPartId = `answer:${body.turnId}`
            controller.enqueue({ type: 'text-start', id: textPartId })
            controller.enqueue({ type: 'text-delta', id: textPartId, delta: finalText })
            controller.enqueue({ type: 'text-end', id: textPartId })
            activeTextSegment = finalText
            persistTextSegment()
          }
          send('answer.done', {
            text: finalText,
            runId: agentOutput.runId,
            grounded: Boolean(analysis || research)
          })
          finishPhase('answering')

          const completedToolNames = [...toolNames]
          send('tools', {
            names: completedToolNames,
            workbookToolsUsed: completedToolNames.filter((name) => ['inspectWorkbook', 'analyzeWorkbook', 'createChart', 'exportAnalysisWorkbook'].includes(name)).length
          })

          const plan = activeRuntime.getPlan()
          const summary = activeRuntime.getSummary()
          send('agent', {
            text: finalText,
            runId: agentOutput.runId
          })
          if (analysis) send('analysis', { analysis })
          if (research) send('research', { research })
          const followups = summary
            ? createWorkbookFollowups({
                summary,
                analysis,
                workflowRequested: activeRuntime.isWorkflowRequested()
              })
            : createGeneralFollowups(Boolean(research), activeRuntime.isSpreadsheetWorkflowRequested())
          send('followups', { suggestions: followups })

          const completeTurn = async () => {
            await completeConversationTurn(env, {
              conversationId: body.threadId,
              turnId: body.turnId,
              content: finalText,
              taskId,
              analysis,
              research,
              followups: followups as ConversationAssistantMetadata['followups'],
              plan,
              timeline
            })
            turnFinalized = true
          }

          if (plan?.action === 'edit_workbook' && workbook) {
            if (creditReservation) {
              // editWorkbook only validates here; its charge is deferred until
              // publication. Include that planned charge in the hold even when
              // a short follow-up produced a low initial estimate. No-op/failure
              // settlement still refunds the unused edit and artifact credits.
              const plannedEditTools = [...completedBillableTools, 'editWorkbook']
              creditReservation = await ensureTurnCreditReservation(env, creditReservation.id, calculateActualTurnCredits(plannedEditTools, true))
              send('credit.estimated', { reservationId: creditReservation.id, turnId: body.turnId, status: 'reserved', refundedCredits: 0, estimatedCredits: creditReservation.estimatedCredits, balance: creditReservation.balanceAfterHold })
            }
            taskId = await stableMutationId(identity.userId, body.threadId, body.turnId)
            const conversation = await getConversationForUser(env, body.threadId, identity.userId)
            if (!conversation) throw new HttpError(409, 'CONVERSATION_DELETED', 'The conversation is no longer available')
            const payload = await createMutationTask(env, {
              kind: 'workbook_mutation', taskId, userId: identity.userId, sessionId: body.sessionId,
              threadId: body.threadId, turnId: body.turnId, inputR2Key: workbook.r2Key, inputName: workbook.fileName,
              plan: plan.mutation, resultWorkbookId: crypto.randomUUID(), sourceRevision: sourceWorkbookRevision,
              creditReservationId: creditReservation?.id || null, creditTurnId: body.turnId, creditToolNames: [...completedBillableTools]
            }, plan, body.message)
            // Durable outbox owns recovery/settlement from this point, even if
            // dispatch or the HTTP response fails after task creation.
            creditDelegatedToWorkflow = true
            send('task', { taskId, status: 'queued', taskType: 'workbook_mutation' })
            send('plan', { plan, summary })
            await completeTurn()
            try {
              const instance = await dispatchWorkbookMutation(env, payload)
              send('workflow', { workflowInstanceId: instance.id, status: 'queued', taskType: 'workbook_mutation' })
            } catch (error) {
              console.error('Workbook edit dispatch will be retried from durable outbox', { taskId, error: internalErrorMessage(error) })
            }
            send('done', { taskId, status: 'queued' })
          } else if (analysis && summary && plan && plan.action !== 'create_spreadsheet' && plan.action !== 'edit_workbook' && activeRuntime.isWorkflowRequested() && workbook) {
            taskId = crypto.randomUUID()
            await createExcelAgentTask(env, {
              id: taskId,
              userId: identity.userId,
              sessionId: body.sessionId,
              threadId: body.threadId,
              inputR2Key: workbook.r2Key,
              inputName: workbook.fileName,
              prompt: body.message,
              taskType: 'workbook_export'
            })
            send('task', { taskId, status: 'planning', taskType: 'workbook_export' })
            await saveTaskPlan(env, taskId, plan)
            send('plan', { plan, summary })

            if (creditReservation) await attachReservationTask(env, creditReservation.id, taskId)
            const workflowInstance = await env.EXCEL_AGENT_WORKFLOW.create({
              id: `excel-${taskId}`,
              params: {
                taskId,
                userId: identity.userId,
                sessionId: body.sessionId,
                inputR2Key: workbook.r2Key,
                inputName: workbook.fileName,
                prompt: body.message,
                plan,
                analysis,
                creditReservationId: creditReservation?.id || null,
                creditTurnId: body.turnId,
                creditToolNames: [...completedBillableTools]
              }
            })
            creditDelegatedToWorkflow = true
            await saveWorkflowInstance(env, taskId, workflowInstance.id)
            send('workflow', {
              workflowInstanceId: workflowInstance.id,
              status: 'queued',
              taskType: 'workbook_export'
            })
            await completeTurn()
            send('done', { taskId, status: 'queued' })
          } else if (plan?.action === 'create_spreadsheet' && activeRuntime.isSpreadsheetWorkflowRequested()) {
            const spreadsheetPlan = activeRuntime.getSpreadsheetPlan()
            if (!spreadsheetPlan) throw new Error('Spreadsheet plan was not available after tool completion')
            taskId = crypto.randomUUID()
            const spreadsheetRecords = activeRuntime.getSpreadsheetRecords()
            if (spreadsheetPlan.dataMode === 'conversation_data' || spreadsheetPlan.dataMode === 'model_knowledge' || spreadsheetRecords.length > 0) {
              const recordsR2Key = `agent-demo/tasks/${taskId}/records.json`
              await env.WORKBOOKS.put(recordsR2Key, JSON.stringify({ schemaVersion: 2, records: spreadsheetRecords }), {
                httpMetadata: { contentType: 'application/json; charset=utf-8' },
                customMetadata: { taskId, kind: spreadsheetPlan.dataMode === 'model_knowledge' ? 'model-knowledge-records' : 'conversation-records' }
              })
            }
            let inputR2Key = activeRuntime.getResearchSnapshotR2Key()
            if (!inputR2Key) {
              inputR2Key = `agent-demo/tasks/${taskId}/request.json`
              await env.WORKBOOKS.put(inputR2Key, JSON.stringify({
                prompt: body.message,
                plan: spreadsheetPlan,
                createdAt: new Date().toISOString()
              }), { httpMetadata: { contentType: 'application/json; charset=utf-8' } })
            }
            await createExcelAgentTask(env, {
              id: taskId,
              userId: identity.userId,
              sessionId: body.sessionId,
              threadId: body.threadId,
              inputR2Key,
              inputName: `${spreadsheetPlan.title}.json`,
              prompt: body.message,
              taskType: 'spreadsheet_generation'
            })
            send('task', { taskId, status: 'planning', taskType: 'spreadsheet_generation' })
            await saveTaskPlan(env, taskId, plan)
            send('plan', { plan })

            if (creditReservation) await attachReservationTask(env, creditReservation.id, taskId)
            const workflowInstance = await env.SPREADSHEET_GENERATION_WORKFLOW.create({
              id: `spreadsheet-${taskId}`,
              params: {
                taskId,
                userId: identity.userId,
                sessionId: body.sessionId,
                prompt: body.message,
                plan: spreadsheetPlan,
                workbookSelection: generationSelection,
                researchSnapshotR2Key: activeRuntime.getResearchSnapshotR2Key(),
                analysis,
                creditReservationId: creditReservation?.id || null,
                creditTurnId: body.turnId,
                creditToolNames: [...completedBillableTools]
              }
            })
            creditDelegatedToWorkflow = true
            await saveWorkflowInstance(env, taskId, workflowInstance.id)
            send('workflow', {
              workflowInstanceId: workflowInstance.id,
              status: 'queued',
              taskType: 'spreadsheet_generation'
            })
            await completeTurn()
            send('done', { taskId, status: 'queued' })
          } else {
            await completeTurn()
            if (creditReservation) {
              const settlement = await settleTurnCredits(env, {
                reservationId: creditReservation.id,
                actualCredits: calculateActualTurnCredits(completedBillableTools),
                actions: buildCreditUsageActions(body.turnId, completedBillableTools)
              })
              creditFinalized = true
              send('credit.settled', settlement)
            }
            send('done', { status: 'completed' })
          }
        } catch (error) {
          if (timedOut) return
          clearAnswerIdleTimeout()
          if (runAbortController.signal.aborted || isAbortError(error)) {
            await persistFailure('REQUEST_CANCELLED')
            return
          }
          finishPhase('thinking', true)
          finishPhase('synthesizing', true)
          finishPhase('answering', true)
          const internalMessage = internalErrorMessage(error)
          if (taskId && !creditDelegatedToWorkflow) {
            try {
              await failTask(env, taskId, internalMessage)
            } catch (persistError) {
              console.error('Agent failure could not be persisted', persistError)
            }
          }
          console.error('Agent pipeline failed', { requestId, taskId, error: internalMessage })
          const publicError = publicStreamError(error)
          await persistFailure(publicError.code)
          if (creditReservation && !creditFinalized && !creditDelegatedToWorkflow) {
            const release = await releaseTurnCredits(env, creditReservation.id, publicError.code)
            creditFinalized = true
            if (release) send('credit.released', release)
          }
          send('error', {
            ...publicError,
            taskId
          })
          finishStream('error')
        } finally {
          if (timeoutPersistence) await timeoutPersistence
          if (runAbortController.signal.aborted && !timedOut) await persistFailure('REQUEST_CANCELLED')
          if (creditReservation && !creditFinalized && !creditDelegatedToWorkflow) {
            try {
              await releaseTurnCredits(env, creditReservation.id, timedOut ? timeoutErrorCode || 'AGENT_TIMEOUT' : 'REQUEST_CANCELLED')
              creditFinalized = true
            } catch (releaseError) {
              console.error('Credit reservation could not be released', {
                requestId,
                error: internalErrorMessage(releaseError)
              })
            }
          }
          if (requestTimeout) clearTimeout(requestTimeout)
          clearAnswerIdleTimeout()
          request.signal.removeEventListener('abort', abortFromRequest)
          if (!closed) {
            finishStream()
            closed = true
            try {
              controller.close()
            } catch {
              // The client may have closed the response between the final event and cleanup.
            }
          }
        }
      })()
      // Keep cancellation/refund cleanup alive after the browser closes SSE.
      executionContext?.waitUntil(completion)
      return completion
    },
    cancel(reason) {
      closed = true
      abortRun(reason)
    }
  })

  const headers = corsHeaders(request, env)
  headers.set('cache-control', 'no-cache, no-transform')
  headers.set('content-encoding', 'identity')
  headers.set('x-accel-buffering', 'no')
  return withSseHeartbeat(createUIMessageStreamResponse({ stream, headers }), 15_000)
}

function isAbortError(error: unknown) {
  return error instanceof Error && error.name === 'AbortError'
}

function normalizeAgentText(value: string) {
  return value.trim()
}

function streamPartId(event: ExcelGenAgentDataParts['agent']['event'], data: Record<string, unknown>) {
  if (event === 'phase' && data.phase) return { id: `phase:${String(data.phase)}` }
  if (event === 'tool' && data.toolCallId) return { id: `tool:${String(data.toolCallId)}` }
  if (event === 'workflow') return { id: 'workflow:output' }
  if (event === 'artifact') return { id: 'artifact:output' }
  return {}
}

function createGeneralFollowups(searched: boolean, spreadsheetRequested: boolean) {
  if (spreadsheetRequested) {
    return [
      { id: 'review-columns', kind: 'review_columns', params: {}, intent: 'analyze' },
      { id: 'check-sources', kind: 'check_sources', params: {}, intent: 'quality' }
    ]
  }
  if (searched) {
    return [
      { id: 'compare-findings', kind: 'compare_findings', params: {}, intent: 'analyze' },
      { id: 'create-research-sheet', kind: 'create_research_sheet', params: {}, intent: 'export' },
      { id: 'show-sources', kind: 'rank_sources', params: {}, intent: 'quality' }
    ]
  }
  return []
}
