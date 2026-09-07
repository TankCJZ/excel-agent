import { NonRetryableError } from 'cloudflare:workflows'
import type { WorkflowStep } from 'cloudflare:workers'
import { handleAgentRequest } from '../../src/http/router.ts'
import { releaseExpiredCredits } from '../../src/billing/scheduled.ts'
import { SpreadsheetGenerationWorkflow as ProductionSpreadsheetWorkflow } from '../../src/workflows/spreadsheetGeneration.ts'
import { ExcelAgentWorkflow as ProductionExportWorkflow } from '../../src/workflows/workbookExport.ts'
import { resolveWorkersAIModel } from './localBillingModel.ts'
import type { AgentWorkerEnv } from '../../src/types.ts'

export class ExcelAgentWorkflow extends ProductionExportWorkflow {
  override run(event: Parameters<ProductionExportWorkflow['run']>[0], step: WorkflowStep) {
    return super.run(event, event.payload.prompt.includes('[billing-workflow-fail]')
      ? failAtStep(step, 'validate source workbook') : step)
  }
}

// This entrypoint is explicitly selected by agent:dev:mock; deployments keep src/index.ts.
export class SpreadsheetGenerationWorkflow extends ProductionSpreadsheetWorkflow {
  override run(event: Parameters<ProductionSpreadsheetWorkflow['run']>[0], step: WorkflowStep) {
    if (!event.payload.prompt.includes('[billing-workflow-fail]')) return super.run(event, step)
    return super.run(event, failAtStep(step, 'prepare spreadsheet records'))
  }
}

function failAtStep(step: WorkflowStep, stepName: string) {
  return new Proxy(step, {
    get(target, property) {
      if (property !== 'do') return Reflect.get(target, property)
      return (name: string, ...args: unknown[]) => {
        if (name === stepName) {
          return target.do(name, async () => { throw new NonRetryableError('Simulated workflow failure') })
        }
        return Reflect.apply(target.do, target, [name, ...args])
      }
    }
  })
}

export default {
  fetch(request: Request, env: AgentWorkerEnv, context: ExecutionContext) {
    return handleAgentRequest(request, env, context, { modelFactory: resolveWorkersAIModel })
  },
  scheduled: releaseExpiredCredits
} satisfies ExportedHandler<AgentWorkerEnv>
