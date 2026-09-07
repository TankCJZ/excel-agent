import { errors, jwtVerify } from 'jose'
import { consumeAgentRateLimit } from '#agent/persistence/agentRepository'
import { HttpError } from '#agent/http/responses'
import type { AgentWorkerEnv } from '#agent/types'

export type AgentIdentity = {
  userId: string
  authSessionId: string
  scopes: ReadonlySet<string>
  localDemo: boolean
  billingEnabled: boolean
}

export async function authenticateAgentRequest(request: Request, env: AgentWorkerEnv): Promise<AgentIdentity> {
  const secret = env.AGENT_TOKEN_SECRET?.trim()
  // A loopback URL alone is not proof of development (for example behind a proxy).
  const localDevelopment = getAgentEnvironment(env) === 'development' && isLoopbackRequest(request)
  const localBillingEnabled = localDevelopment && isLocalBillingEnabled(env)
  if (!secret) {
    if (localDevelopment && !localBillingEnabled) {
      return {
        userId: 'local-demo',
        authSessionId: 'local-demo',
        scopes: new Set(['agent:use', 'agent:eval']),
        localDemo: true,
        billingEnabled: false
      }
    }
    throw new HttpError(
      503,
      'AUTH_NOT_CONFIGURED',
      localBillingEnabled
        ? 'Agent authentication is required when local billing is enabled'
        : 'Agent authentication is not configured'
    )
  }

  const authorization = request.headers.get('authorization') || ''
  const match = authorization.match(/^Bearer\s+(.+)$/i)
  if (!match) throw new HttpError(401, 'UNAUTHORIZED', 'Authentication required')

  try {
    const verified = await jwtVerify(match[1], new TextEncoder().encode(secret), {
      algorithms: ['HS256'],
      issuer: 'excelgen-app',
      audience: 'excelgen-agent'
    })
    const scopes = new Set(Array.isArray(verified.payload.scope)
      ? verified.payload.scope.filter((value): value is string => typeof value === 'string')
      : [])
    if (!verified.payload.sub || typeof verified.payload.sid !== 'string' || !scopes.has('agent:use')) {
      throw new HttpError(401, 'UNAUTHORIZED', 'Agent token payload is invalid')
    }
    return {
      userId: verified.payload.sub,
      authSessionId: verified.payload.sid,
      scopes,
      localDemo: false,
      billingEnabled: !localDevelopment || localBillingEnabled
    }
  } catch (error) {
    if (error instanceof HttpError) throw error
    if (error instanceof errors.JOSEError) {
      throw new HttpError(401, 'UNAUTHORIZED', 'Agent token is invalid or expired')
    }
    throw error
  }
}

export function getAgentEnvironment(env: Pick<AgentWorkerEnv, 'AGENT_ENV'>): 'development' | 'production' {
  // Missing, unknown or misspelled values fail closed. Never infer this from headers.
  return env.AGENT_ENV === 'development' ? 'development' : 'production'
}

export function isLocalBillingEnabled(env: Pick<AgentWorkerEnv, 'AGENT_ENV' | 'LOCAL_BILLING_ENABLED'>) {
  return getAgentEnvironment(env) === 'development'
    && /^(?:1|true|yes|on)$/i.test(env.LOCAL_BILLING_ENABLED?.trim() || '')
}

export function requireAgentScope(identity: AgentIdentity, scope: string) {
  if (!identity.scopes.has(scope)) throw new HttpError(403, 'FORBIDDEN', 'This operation is not permitted')
}

export async function enforceAgentRateLimit(
  env: AgentWorkerEnv,
  identity: AgentIdentity,
  scope: string,
  limit: number,
  windowSeconds: number
) {
  const now = Date.now()
  const windowStart = Math.floor(now / (windowSeconds * 1000))
  const key = `${identity.userId}:${scope}:${windowStart}`
  const expiresAt = new Date((windowStart + 1) * windowSeconds * 1000).toISOString()
  const updatedAt = new Date(now).toISOString()
  const count = await consumeAgentRateLimit(env, key, expiresAt, updatedAt)
  if (count > limit) {
    const retryAfter = Math.max(1, Math.ceil((Date.parse(expiresAt) - now) / 1000))
    throw new HttpError(429, 'RATE_LIMITED', 'Too many agent requests', { retryAfter })
  }
}

function isLoopbackRequest(request: Request) {
  const hostname = new URL(request.url).hostname
  return hostname === 'localhost' || hostname === '127.0.0.1' || hostname === '[::1]'
}
