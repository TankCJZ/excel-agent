import {
  WorkflowEntrypoint,
  type WorkflowEvent,
  type WorkflowStep
} from 'cloudflare:workers'
import { prepareSpreadsheetData, writeSpreadsheetResult } from '#agent/spreadsheet/generation'
import { appendTaskEvent, completeTask, failTask, markTaskProcessing } from '#agent/persistence/agentRepository'
import type { AgentWorkerEnv, SpreadsheetWorkflowPayload } from '#agent/types'
import { settleTurnCredits } from '#agent/billing/billingRepository'
import { buildCreditUsageActions, calculateActualTurnCredits } from '#shared/billing/planCatalog'
import { registerGeneratedWorkbook } from '#agent/persistence/generatedWorkbook'

export class SpreadsheetGenerationWorkflow extends WorkflowEntrypoint<AgentWorkerEnv, SpreadsheetWorkflowPayload> {
  async run(event: WorkflowEvent<SpreadsheetWorkflowPayload>, step: WorkflowStep) {
    const payload = event.payload
    try {
      await step.do('mark spreadsheet task processing', async () => {
        await markTaskProcessing(this.env, payload.taskId, 'Preparing spreadsheet data')
        return { taskId: payload.taskId, status: 'processing' }
      })

      await step.do('record spreadsheet structure', async () => {
        await appendTaskEvent(this.env, payload.taskId, 'spreadsheet.structuring', 'Validating the Agent-designed spreadsheet structure', {
          title: payload.plan.title,
          dataMode: payload.plan.dataMode,
          columnCount: payload.plan.columns.length,
          maxRows: payload.plan.maxRows
        })
        return { columns: payload.plan.columns.length }
      })

      const prepared = await step.do(
        'prepare spreadsheet records',
        { retries: { limit: 2, delay: '3 seconds', backoff: 'exponential' }, timeout: '3 minutes' },
        async () => {
          await appendTaskEvent(this.env, payload.taskId, 'spreadsheet.generating', 'Generating structured rows and the Excel workbook')
          return prepareSpreadsheetData(this.env, payload)
        }
      )

      const generated = await step.do(
        'render and store spreadsheet workbook',
        { retries: { limit: 2, delay: '3 seconds', backoff: 'exponential' }, timeout: '3 minutes' },
        async () => writeSpreadsheetResult(this.env, payload, prepared)
      )

      const result = payload.workbookSelection
        ? await step.do('register generated workbook version', () => registerGeneratedWorkbook(this.env, payload, generated))
        : generated

      if (payload.creditReservationId) {
        await step.do('settle spreadsheet credits', async () => settleTurnCredits(this.env, {
          reservationId: payload.creditReservationId!,
          taskId: payload.taskId,
          actualCredits: calculateActualTurnCredits(payload.creditToolNames, true),
          actions: buildCreditUsageActions(payload.creditTurnId, payload.creditToolNames, true)
        }))
      }
      await step.do('persist spreadsheet result', async () => {
        await completeTask(this.env, payload.taskId, result, payload.workbookSelection
          ? { ...payload.workbookSelection, userId: payload.userId, sessionId: payload.sessionId }
          : undefined)
        return { taskId: payload.taskId, outputR2Key: result.outputR2Key }
      })
      return result
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : 'Unknown spreadsheet workflow failure'
      if (payload.creditReservationId) {
        await step.do('settle completed Agent work after spreadsheet failure', async () => settleTurnCredits(this.env, {
          reservationId: payload.creditReservationId!,
          taskId: payload.taskId,
          actualCredits: calculateActualTurnCredits(payload.creditToolNames),
          actions: buildCreditUsageActions(payload.creditTurnId, payload.creditToolNames)
        }))
      }
      await step.do('persist spreadsheet workflow failure', async () => {
        await failTask(this.env, payload.taskId, errorMessage)
        return { taskId: payload.taskId, failed: true }
      })
      throw error
    }
  }
}
