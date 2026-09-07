import type { UIMessage } from 'ai'

export type AgentStreamEventName =
  | 'connected'
  | 'phase'
  | 'tool'
  | 'tools'
  | 'answer.done'
  | 'agent'
  | 'analysis'
  | 'research'
  | 'followups'
  | 'task'
  | 'plan'
  | 'workflow'
  | 'done'
  | 'error'
  | 'credit.estimated'
  | 'credit.settled'
  | 'credit.released'
  | 'artifact'

export type AgentStreamEvent = {
  event: AgentStreamEventName
  data: Record<string, unknown>
}

export type ExcelGenAgentDataParts = {
  agent: AgentStreamEvent
}

export type ExcelGenAgentMessage = UIMessage<
  { requestId?: string, turnId?: string },
  ExcelGenAgentDataParts
>
