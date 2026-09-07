import type { RequestedOutcome, WebSearchResult, WorkbookSummary } from '#agent/contracts'
import type { WorkbookAnalysisContext } from '#agent/workbook/service'

export type ExcelAgentRuntimeInput = {
  workbookId?: string
  fileKey?: string
  fileName?: string
  supportsWorkbookMutation?: boolean
  contextR2Key?: string | null
  contextVersion?: number | null
  sourceEtag?: string | null
  summary?: WorkbookSummary
  userId?: string
  sessionId: string
  billingEnabled?: boolean
  initialContext?: WorkbookAnalysisContext
  initialResearch?: WebSearchResult | null
  initialResearchSnapshotR2Key?: string | null
  persistContextMetadata?: boolean
  requestedOutcome?: RequestedOutcome
}
