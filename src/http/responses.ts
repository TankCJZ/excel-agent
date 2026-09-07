import type { AgentWorkerEnv } from '#agent/types'
import { ConversationStateError } from '#agent/persistence/conversationRepository'
import { WorkbookMutationError } from '#agent/workbook/mutation/errors'


export class HttpError extends Error {
  readonly status: number
  readonly code: string
  readonly details?: unknown

  constructor(
    status: number,
    code: string,
    message: string,
    details?: unknown
  ) {
    super(message)
    this.status = status
    this.code = code
    this.details = details
  }
}

export function jsonResponse(
  request: Request,
  env: AgentWorkerEnv,
  requestId: string,
  data: unknown,
  status = 200
) {
  const headers = corsHeaders(request, env)
  headers.set('content-type', 'application/json; charset=utf-8')
  return new Response(JSON.stringify({ ok: true, data, requestId }), { status, headers })
}

export function errorResponse(request: Request, env: AgentWorkerEnv, requestId: string, error: unknown) {
  const normalized = normalizeHttpError(error)
  if (normalized.status >= 500) {
    console.error('Agent request failed', {
      requestId,
      code: normalized.code,
      error: internalErrorMessage(error)
    })
  }
  const headers = corsHeaders(request, env)
  headers.set('content-type', 'application/json; charset=utf-8')
  return new Response(JSON.stringify({
    ok: false,
    error: {
      code: normalized.code,
      message: normalized.message,
      ...(normalized.details === undefined ? {} : { details: normalized.details })
    },
    requestId
  }), { status: normalized.status, headers })
}

export function normalizeHttpError(error: unknown) {
  if (error instanceof HttpError) return error
  if (error instanceof ConversationStateError) return new HttpError(409, error.code, error.message)
  if (error instanceof WorkbookMutationError) return new HttpError(422, error.code, error.message)
  if (error && typeof error === 'object' && 'issues' in error) {
    return new HttpError(400, 'VALIDATION_ERROR', 'Request validation failed', error.issues)
  }
  return new HttpError(500, 'INTERNAL_ERROR', 'Unexpected server error')
}

export function publicStreamError(error: unknown) {
  if (error instanceof HttpError || error instanceof ConversationStateError || error instanceof WorkbookMutationError) {
    return { code: error.code, message: error.message }
  }
  return { code: 'AGENT_PIPELINE_FAILED', message: 'The agent request could not be completed' }
}

export function internalErrorMessage(error: unknown) {
  return error instanceof Error ? error.message : 'Unexpected server error'
}

export function corsHeaders(request: Request, env: AgentWorkerEnv) {
  const headers = new Headers({
    'access-control-allow-methods': 'GET, POST, PATCH, DELETE, OPTIONS',
    'access-control-allow-headers': 'content-type, authorization',
    'access-control-max-age': '86400',
    'vary': 'Origin',
    'x-content-type-options': 'nosniff'
  })
  const origin = request.headers.get('origin')
  const allowedOrigins = (env.ALLOWED_ORIGINS || '').split(',').map((value) => value.trim()).filter(Boolean)
  if (origin && allowedOrigins.includes(origin)) headers.set('access-control-allow-origin', origin)
  return headers
}
