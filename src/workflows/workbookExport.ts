import {
  WorkflowEntrypoint,
  type WorkflowEvent,
  type WorkflowStep
} from 'cloudflare:workers'
import { NonRetryableError } from 'cloudflare:workflows'
import type { AgentWorkerEnv, ExcelWorkflowPayload } from '#agent/types'
import { completeTask, failTask, markTaskProcessing } from '#agent/persistence/agentRepository'
import { processWorkbook } from '#agent/workbook/service'
import { settleTurnCredits } from '#agent/billing/billingRepository'
import { buildCreditUsageActions, calculateActualTurnCredits } from '#shared/billing/planCatalog'
import { runWorkbookMutation } from '#agent/workflows/workbookMutation'

export class ExcelAgentWorkflow extends WorkflowEntrypoint<AgentWorkerEnv, ExcelWorkflowPayload> {
  async run(event: WorkflowEvent<ExcelWorkflowPayload>, step: WorkflowStep) {
    const payload = event.payload
    if (payload.kind === 'workbook_mutation') return runWorkbookMutation(this.env, payload, step)
    if (payload.kind !== undefined && payload.kind !== 'workbook_export') throw new NonRetryableError('Unknown workbook workflow protocol')

    try {
      await step.do('mark task processing', async () => {
        await markTaskProcessing(this.env, payload.taskId)
        return { taskId: payload.taskId, status: 'processing' }
      })

      await step.do(
        'validate source workbook',
        {
          retries: { limit: 2, delay: '2 seconds', backoff: 'linear' },
          timeout: '2 minutes'
        },
        async () => {
          const exists = await this.env.WORKBOOKS.head(payload.inputR2Key)
          if (!exists) {
            throw new NonRetryableError('Input workbook does not exist in R2')
          }
          return { inputR2Key: payload.inputR2Key, etag: exists.etag }
        }
      )

      const result = await step.do(
        'transform workbook with SheetJS',
        {
          retries: { limit: 2, delay: '3 seconds', backoff: 'exponential' },
          timeout: '5 minutes'
        },
        async () => processWorkbook(this.env.WORKBOOKS, payload)
      )

      if (payload.creditReservationId) {
        await step.do('settle workbook credits', async () => settleTurnCredits(this.env, {
          reservationId: payload.creditReservationId!,
          taskId: payload.taskId,
          actualCredits: calculateActualTurnCredits(payload.creditToolNames, true),
          actions: buildCreditUsageActions(payload.creditTurnId, payload.creditToolNames, true)
        }))
      }

      await step.do('persist result metadata to D1', async () => {
        await completeTask(this.env, payload.taskId, result)
        return {
          taskId: payload.taskId,
          outputR2Key: result.outputR2Key
        }
      })

      return result
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : 'Unknown workflow failure'
      if (payload.creditReservationId) {
        await step.do('settle completed Agent work after workbook failure', async () => settleTurnCredits(this.env, {
          reservationId: payload.creditReservationId!,
          taskId: payload.taskId,
          actualCredits: calculateActualTurnCredits(payload.creditToolNames),
          actions: buildCreditUsageActions(payload.creditTurnId, payload.creditToolNames)
        }))
      }
      await step.do('persist workflow failure to D1', async () => {
        await failTask(this.env, payload.taskId, errorMessage)
        return { taskId: payload.taskId, failed: true }
      })
      throw error
    }
  }
}
