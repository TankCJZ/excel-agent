import { followupPresentationSchema, followupSuggestionSchema, type FollowupPresentation, type FollowupSuggestion } from '#shared/agent/followups'
import { createFollowupCopyGenerator, localizeFollowups, type FollowupCopyGenerator } from '#agent/mastra/localizeFollowups'
import { HttpError } from '#agent/http/responses'
import type { AgentWorkerEnv } from '#agent/types'

type Source = { id: string; question: string; answer: string; metadata_json: string }
function readMetadata(source: Source) {
  try {
    const value: unknown = JSON.parse(source.metadata_json)
    return value && typeof value === 'object' && !Array.isArray(value) ? value as Record<string, unknown> : {}
  } catch { return {} }
}
function candidates(metadata: Record<string, unknown>): FollowupSuggestion[] {
  return (Array.isArray(metadata.followups) ? metadata.followups : []).flatMap(value => {
    const parsed = followupSuggestionSchema.safeParse(value)
    return parsed.success ? [parsed.data] : []
  }).slice(0, 3)
}
function cached(metadata: Record<string, unknown>, suggestions: FollowupSuggestion[]): FollowupPresentation | null {
  const parsed = followupPresentationSchema.safeParse(metadata.followupPresentation)
  if (!parsed.success || parsed.data.suggestions.length !== suggestions.length) return null
  return parsed.data.suggestions.every((item, index) => {
    const source = suggestions[index]!
    return item.id === source.id && item.kind === source.kind && item.intent === source.intent && JSON.stringify(item.params) === JSON.stringify(source.params)
  }) ? parsed.data : null
}

export async function getLocalizedFollowups(
  env: AgentWorkerEnv,
  input: { userId: string; conversationId: string; turnId: string; signal: AbortSignal },
  options: { generate?: FollowupCopyGenerator; beforeGenerate?: () => Promise<void> } = {}
) {
  const read = () => env.DB.prepare(`SELECT a.id, a.content AS answer, u.content AS question, a.metadata_json
    FROM excel_agent_messages a
    JOIN excel_agent_conversations c ON c.id = a.conversation_id
    JOIN excel_agent_messages u ON u.conversation_id = a.conversation_id AND u.turn_id = a.turn_id AND u.role = 'user'
    WHERE c.user_id = ? AND c.id = ? AND a.turn_id = ? AND a.role = 'assistant' AND a.status = 'completed'
      AND NOT EXISTS (SELECT 1 FROM excel_agent_conversation_tombstones t WHERE t.conversation_id = c.id)
    LIMIT 1`).bind(input.userId, input.conversationId, input.turnId).first<Source>()
  const source = await read()
  if (!source) throw new HttpError(404, 'NOT_FOUND', 'Completed conversation turn was not found')
  const metadata = readMetadata(source)
  const suggestions = candidates(metadata)
  if (!suggestions.length) return null
  const existing = cached(metadata, suggestions)
  if (existing) return existing
  await options.beforeGenerate?.()

  // Optional presentation work is separate from the answer/task stream and never charges a chat action.
  const timeout = new AbortController()
  const timer = setTimeout(() => timeout.abort(), 15_000)
  const signal = AbortSignal.any([input.signal, timeout.signal])
  try {
    const presentation = await localizeFollowups({ question: source.question, answer: source.answer, suggestions }, options.generate || createFollowupCopyGenerator(env), signal)
    signal.throwIfAborted()
    const saved = await env.DB.prepare(`UPDATE excel_agent_messages
      SET metadata_json = json_set(metadata_json, '$.followupPresentation', json(?))
      WHERE id = ? AND metadata_json = ? AND content = ? AND status = 'completed'
        AND EXISTS (SELECT 1 FROM excel_agent_conversations c WHERE c.id = excel_agent_messages.conversation_id AND c.user_id = ?)
        AND NOT EXISTS (SELECT 1 FROM excel_agent_conversation_tombstones t WHERE t.conversation_id = excel_agent_messages.conversation_id)`)
      .bind(JSON.stringify(presentation), source.id, source.metadata_json, source.answer, input.userId).run()
    if (saved.meta.changes) return presentation
    // A retry, deletion or concurrent localization may have won; never return stale copy.
    const latest = await read()
    return latest ? cached(readMetadata(latest), candidates(readMetadata(latest))) : null
  } catch {
    if (!signal.aborted) console.warn('Follow-up copy unavailable', { reason: 'generation_or_persistence_failed' })
    // A missing optional suggestion must not make a successful answer fail or flash English fallback text.
    return null
  } finally { clearTimeout(timer) }
}
