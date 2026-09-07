export const DEFAULT_AGENT_REQUEST_TIMEOUT_MS = 90_000
export const DEFAULT_AGENT_ANSWER_TIMEOUT_MS = 45_000

export const MIN_AGENT_REQUEST_TIMEOUT_MS = 5_000
export const MAX_AGENT_REQUEST_TIMEOUT_MS = 120_000
export const MIN_AGENT_ANSWER_TIMEOUT_MS = 5_000
export const MAX_AGENT_ANSWER_TIMEOUT_MS = 90_000

export function resolveAgentRequestTimeoutMs(value?: string) {
  if (!value?.trim()) return DEFAULT_AGENT_REQUEST_TIMEOUT_MS

  const parsed = Number(value)
  if (!Number.isFinite(parsed)) return DEFAULT_AGENT_REQUEST_TIMEOUT_MS

  return Math.min(
    MAX_AGENT_REQUEST_TIMEOUT_MS,
    Math.max(MIN_AGENT_REQUEST_TIMEOUT_MS, Math.floor(parsed))
  )
}

export function resolveAgentAnswerTimeoutMs(value?: string) {
  if (!value?.trim()) return DEFAULT_AGENT_ANSWER_TIMEOUT_MS

  const parsed = Number(value)
  if (!Number.isFinite(parsed)) return DEFAULT_AGENT_ANSWER_TIMEOUT_MS

  return Math.min(
    MAX_AGENT_ANSWER_TIMEOUT_MS,
    Math.max(MIN_AGENT_ANSWER_TIMEOUT_MS, Math.floor(parsed))
  )
}
