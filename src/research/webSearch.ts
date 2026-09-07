import type {
  WebSearchInput,
  WebSearchResult,
  WebSearchSource
} from '#agent/contracts'

import type { AgentWorkerEnv } from '#agent/types'

export const DEFAULT_WEB_SEARCH_MODEL = 'openai/gpt-4o-mini'
export const DEFAULT_WEB_SEARCH_TIMEOUT_MS = 15_000
export const AGENT_WEB_SEARCH_SOURCE_LIMIT = 12
export const AGENT_WEB_SEARCH_SUMMARY_CHAR_LIMIT = 10_000
export const AGENT_WEB_SEARCH_SNIPPET_CHAR_LIMIT = 320

type ResponsesRunner = {
  run: (model: string, input: Record<string, unknown>, options?: Record<string, unknown>) => Promise<unknown>
}

export async function searchWeb(
  env: AgentWorkerEnv,
  input: WebSearchInput & { sessionId: string }
) {
  if (!env.AI) throw new Error('The AI binding is required for web search')

  const researchId = crypto.randomUUID()
  const maxToolCalls = clampInteger(env.WEB_SEARCH_MAX_TOOL_CALLS, 1, 5, 3)
  const domainGuidance = input.preferredDomains.length
    ? `Prefer trustworthy information from these domains when relevant: ${input.preferredDomains.join(', ')}.`
    : 'Prefer primary sources, official documentation, public datasets, and reputable reporting.'
  const recencyGuidance = input.recency === 'any'
    ? 'Use the most current reliable information available.'
    : `Prioritize information from the last ${input.recency}.`

  const raw = await runResponsesModel(env, {
    instructions: [
      'You are the web research tool for a spreadsheet agent.',
      'Search the web before answering. Return concise factual findings that can be transformed into structured spreadsheet rows.',
      'Keep names, dates, quantities, units, currencies, and source attribution explicit.',
      'Do not invent missing values. Clearly distinguish facts from uncertainty.',
      domainGuidance,
      recencyGuidance
    ].join(' '),
    input: `Research objective: ${input.objective}\nSearch queries:\n${input.queries.map((query, index) => `${index + 1}. ${query}`).join('\n')}`,
    tools: [{ type: 'web_search', search_context_size: 'medium' }],
    tool_choice: 'required',
    max_tool_calls: maxToolCalls,
    max_output_tokens: 5000,
    include: ['web_search_call.action.sources', 'web_search_call.results'],
    store: false
  })

  const normalized = normalizeWebSearchResponse(raw, {
    researchId,
    objective: input.objective,
    retrievedAt: new Date().toISOString()
  })
  if (!normalized.summary) throw new Error('Web search returned no readable answer')

  const snapshotR2Key = researchSnapshotR2Key(input.sessionId, researchId)
  await env.WORKBOOKS.put(snapshotR2Key, JSON.stringify(normalized), {
    httpMetadata: { contentType: 'application/json; charset=utf-8' },
    customMetadata: {
      researchId,
      sessionId: input.sessionId,
      model: env.WEB_SEARCH_MODEL || DEFAULT_WEB_SEARCH_MODEL
    }
  })

  return { result: normalized, snapshotR2Key }
}

export function researchSnapshotR2Key(sessionId: string, researchId: string) {
  return `agent-demo/research/${sessionId}/${researchId}.json`
}

export function compactWebSearchResultForAgent(result: WebSearchResult): WebSearchResult {
  return {
    ...result,
    summary: result.summary.slice(0, AGENT_WEB_SEARCH_SUMMARY_CHAR_LIMIT),
    sources: result.sources.slice(0, AGENT_WEB_SEARCH_SOURCE_LIMIT).map(source => ({
      ...source,
      ...(source.snippet
        ? { snippet: source.snippet.slice(0, AGENT_WEB_SEARCH_SNIPPET_CHAR_LIMIT) }
        : {})
    }))
  }
}

export async function runResponsesModel(env: AgentWorkerEnv, input: Record<string, unknown>) {
  if (!env.AI) throw new Error('The AI binding is required for this model request')
  const runner = env.AI as unknown as ResponsesRunner
  const gatewayId = env.AI_GATEWAY_ID?.trim()
  const timeoutMs = clampInteger(env.WEB_SEARCH_TIMEOUT_MS, 3_000, 30_000, DEFAULT_WEB_SEARCH_TIMEOUT_MS)
  return runner.run(
    env.WEB_SEARCH_MODEL?.trim() || DEFAULT_WEB_SEARCH_MODEL,
    input,
    {
      signal: AbortSignal.timeout(timeoutMs),
      ...(gatewayId ? { gateway: { id: gatewayId, metadata: { application: 'excelgen-agent' } } } : {})
    }
  )
}

export function normalizeWebSearchResponse(
  raw: unknown,
  metadata: { researchId: string, objective: string, retrievedAt: string }
): WebSearchResult {
  return {
    ...metadata,
    summary: extractResponseText(raw).trim(),
    sources: extractWebSearchSources(raw)
  }
}

export function extractResponseText(raw: unknown) {
  if (!raw || typeof raw !== 'object') return ''
  const response = raw as Record<string, unknown>
  if (typeof response.output_text === 'string') return response.output_text
  if (typeof response.text === 'string') return response.text
  if (typeof response.response === 'string') return response.response

  const text: string[] = []
  for (const item of Array.isArray(response.output) ? response.output : []) {
    if (!item || typeof item !== 'object') continue
    const output = item as Record<string, unknown>
    if (typeof output.text === 'string') text.push(output.text)
    for (const part of Array.isArray(output.content) ? output.content : []) {
      if (!part || typeof part !== 'object') continue
      const content = part as Record<string, unknown>
      if (typeof content.text === 'string') text.push(content.text)
      else if (typeof content.output_text === 'string') text.push(content.output_text)
    }
  }
  return text.join('\n')
}

export function extractWebSearchSources(raw: unknown): WebSearchSource[] {
  const candidates: Array<{ url: string, title?: string, snippet?: string }> = []
  walk(raw, (value) => {
    if (!value || typeof value !== 'object' || Array.isArray(value)) return
    const item = value as Record<string, unknown>
    const url = typeof item.url === 'string' ? item.url : ''
    if (!/^https?:\/\//i.test(url)) return
    candidates.push({
      url,
      title: typeof item.title === 'string' ? item.title : undefined,
      snippet: typeof item.snippet === 'string'
        ? item.snippet
        : typeof item.text === 'string' && item.text.length < 800
          ? item.text
          : undefined
    })
  })

  const unique = new Map<string, { url: string, title?: string, snippet?: string }>()
  for (const candidate of candidates) {
    const canonical = canonicalizeUrl(candidate.url)
    const existing = unique.get(canonical)
    unique.set(canonical, {
      url: candidate.url,
      title: candidate.title || existing?.title,
      snippet: candidate.snippet || existing?.snippet
    })
  }

  return [...unique.values()].slice(0, 40).map((source, index) => ({
    id: `S${index + 1}`,
    title: (source.title || hostname(source.url)).slice(0, 300),
    url: source.url,
    ...(source.snippet ? { snippet: source.snippet.slice(0, 800) } : {})
  }))
}

function walk(value: unknown, visit: (value: unknown) => void, seen = new WeakSet<object>()) {
  visit(value)
  if (!value || typeof value !== 'object') return
  if (seen.has(value)) return
  seen.add(value)
  for (const child of Array.isArray(value) ? value : Object.values(value)) walk(child, visit, seen)
}

function canonicalizeUrl(value: string) {
  try {
    const url = new URL(value)
    url.hash = ''
    for (const key of [...url.searchParams.keys()]) {
      if (/^(utm_|gclid|fbclid)/i.test(key)) url.searchParams.delete(key)
    }
    return url.toString()
  } catch {
    return value
  }
}

function hostname(value: string) {
  try {
    return new URL(value).hostname.replace(/^www\./, '')
  } catch {
    return value
  }
}

function clampInteger(value: string | undefined, minimum: number, maximum: number, fallback: number) {
  const parsed = Number.parseInt(value || '', 10)
  return Number.isFinite(parsed) ? Math.min(maximum, Math.max(minimum, parsed)) : fallback
}
