import type {
  RequestedOutcome,
  SpreadsheetColumn,
  SpreadsheetExecutionPlan,
  SpreadsheetPlan,
  WebSearchResult,
  WebSearchSource,
  WorkbookAnalysis,
  WorkbookPlan,
  WorkbookSummary
} from '#agent/contracts'
import type { MutationChanges, WorkbookMutationPlan } from '#shared/agent/workbookMutation'

export type ExcelAgentTaskStatus = 'planning' | 'queued' | 'processing' | 'completed' | 'failed'
export type ExcelAgentMessageRole = 'user' | 'assistant'
export type ExcelAgentMessageStatus = 'completed' | 'streaming' | 'failed'

export type WorkbookExportPayload = {
  kind?: 'workbook_export'
  taskId: string
  userId: string
  sessionId: string
  inputR2Key: string
  inputName: string
  prompt: string
  plan: WorkbookPlan
  analysis: WorkbookAnalysis
  creditReservationId: string | null
  creditTurnId: string
  creditToolNames: string[]
}

export type WorkbookMutationExecutionPlan = {
  action: 'edit_workbook'
  steps: Array<{ id: string; operation: string }>
  mutation: WorkbookMutationPlan
}

export type WorkbookMutationPayload = {
  kind: 'workbook_mutation'
  taskId: string
  userId: string
  sessionId: string
  threadId: string
  turnId: string
  inputR2Key: string
  inputName: string
  plan: WorkbookMutationPlan
  resultWorkbookId: string
  sourceRevision: number
  creditReservationId: string | null
  creditTurnId: string
  creditToolNames: string[]
}

export type ExcelWorkflowPayload = WorkbookExportPayload | WorkbookMutationPayload

export type WorkbookMutationResult = {
  kind: 'workbook_mutation'
  schemaVersion: 1
  taskId: string
  outputR2Key: string
  outputName: string
  summary: WorkbookSummary
  analysis: null
  preview: Array<Array<string | number | boolean | null>>
  previewSheetName: string
  previewRowCount: number
  sizeBytes: number
  changes: MutationChanges
  sourceWorkbookId: string
  workbookId: string
  sourceEtag: string
  contextR2Key: string
  contextVersion: number
}

export type ExcelWorkflowResult = {
  kind: 'workbook_export'
  taskId: string
  outputR2Key: string
  outputName: string
  summary: WorkbookSummary
  analysis: WorkbookAnalysis
  preview: Array<Array<string | number | boolean | null>>
  previewSheetName?: string
  previewRowCount?: number
  sizeBytes: number
}

export type SpreadsheetWorkflowPayload = {
  taskId: string
  userId: string
  sessionId: string
  prompt: string
  plan: SpreadsheetPlan
  researchSnapshotR2Key: string | null
  analysis: WorkbookAnalysis | null
  creditReservationId: string | null
  creditTurnId: string
  creditToolNames: string[]
  workbookSelection?: { threadId: string; workbookId: string | null; revision: number }
}

export type SpreadsheetGenerationResult = {
  kind: 'spreadsheet_generation'
  taskId: string
  outputR2Key: string
  outputName: string
  summary: WorkbookSummary
  analysis: null
  preview: Array<Array<string | number | boolean | null>>
  previewSheetName: string
  previewRowCount: number
  sizeBytes: number
  columns: SpreadsheetColumn[]
  sources: WebSearchSource[]
  workbookId?: string
}

export type AgentTaskPlan = WorkbookPlan | SpreadsheetExecutionPlan | WorkbookMutationExecutionPlan
export type AgentTaskResult = ExcelWorkflowResult | SpreadsheetGenerationResult | WorkbookMutationResult
export type ExcelAgentTaskType = 'workbook_export' | 'spreadsheet_generation' | 'workbook_mutation'

export type AgentWorkerEnv = Omit<Env, 'WORKBOOK_EDITING_ENABLED' | 'AGENT_ENV'> & {
  AGENT_ENV?: string
  AGENT_TOKEN_SECRET?: string
  CLOUDFLARE_ACCOUNT_ID?: string
  CLOUDFLARE_API_TOKEN?: string
  LOCAL_BILLING_ENABLED?: string
  OPENAI_API_KEY?: string
  ANTHROPIC_API_KEY?: string
  GOOGLE_API_KEY?: string
  WORKBOOK_EDITING_ENABLED?: string
}

export type ExcelAgentWorkbookRecord = {
  id: string
  userId: string
  sessionId: string
  r2Key: string
  fileName: string
  sizeBytes: number
  summary: WorkbookSummary
  contextR2Key: string | null
  sourceEtag: string | null
  contextVersion: number | null
  parsedAt: string | null
  createdAt: string
  updatedAt: string
  parentWorkbookId?: string | null
  rootWorkbookId?: string | null
  sourceTaskId?: string | null
}

export type ExcelAgentTaskRecord = {
  id: string
  userId: string
  sessionId: string
  threadId: string
  workflowInstanceId: string | null
  inputR2Key: string
  inputName: string
  outputR2Key: string | null
  outputName: string | null
  prompt: string
  taskType: ExcelAgentTaskType
  plan: AgentTaskPlan | null
  result: AgentTaskResult | null
  status: ExcelAgentTaskStatus
  errorMessage: string | null
  createdAt: string
  updatedAt: string
  completedAt: string | null
}

export type ExcelAgentTaskEvent = {
  id: string
  taskId: string
  type: string
  message: string
  metadata: Record<string, unknown> | null
  createdAt: string
}

export type ConversationTimelineEvent = {
  event: 'phase' | 'tool' | 'tools' | 'text'
  data: Record<string, unknown>
}

export type ConversationUserMetadata = {
  requestedOutcome: RequestedOutcome | null
  workbook: {
    workbookId: string
    fileName: string
    sizeBytes: number
    summary: WorkbookSummary
  } | null
}

export type ConversationAssistantMetadata = {
  followupPresentation?: import('#shared/agent/followups').FollowupPresentation
  analysis: WorkbookAnalysis | null
  research: WebSearchResult | null
  followups: Array<{
    id: string
    intent: string
    kind: string
    params: Record<string, string>
  }>
  plan: AgentTaskPlan | null
  timeline: ConversationTimelineEvent[]
}

export type ExcelAgentConversationRecord = {
  id: string
  userId: string
  sessionId: string
  title: string
  workbookId: string | null
  workbookRevision?: number
  createdAt: string
  updatedAt: string
  lastMessageAt: string | null
}

export type ExcelAgentMessageRecord = {
  id: string
  conversationId: string
  turnId: string
  role: ExcelAgentMessageRole
  status: ExcelAgentMessageStatus
  content: string
  taskId: string | null
  workbookId: string | null
  metadata: Record<string, unknown> | null
  errorCode: string | null
  createdAt: string
  updatedAt: string
}
