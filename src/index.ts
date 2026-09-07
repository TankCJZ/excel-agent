import { handleAgentRequest } from '#agent/http/router'
import { releaseExpiredCredits } from '#agent/billing/scheduled'
import type { AgentWorkerEnv } from '#agent/types'

export { ExcelAgentWorkflow } from '#agent/workflows/workbookExport'
export { SpreadsheetGenerationWorkflow } from '#agent/workflows/spreadsheetGeneration'

export default {
  fetch: handleAgentRequest,
  scheduled: releaseExpiredCredits
} satisfies ExportedHandler<AgentWorkerEnv>
