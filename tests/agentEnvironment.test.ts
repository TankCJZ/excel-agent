import assert from 'node:assert/strict'
import { readFileSync } from 'node:fs'
import test from 'node:test'
import { SignJWT } from 'jose'
import { authenticateAgentRequest, getAgentEnvironment, isLocalBillingEnabled } from '#agent/platform/auth'
import { HttpError } from '#agent/http/responses'
import type { AgentWorkerEnv } from '#agent/types'

const secret = 'agent-environment-test-only-signing-secret'
const localUrls = ['http://localhost:8788/api/chat', 'http://127.0.0.1:8788/api/chat', 'http://[::1]:8788/api/chat']
const publicUrl = 'https://excel-agent.excelgen.app/api/chat'
const errorCode = (code: string) => (error: unknown) => error instanceof HttpError && error.code === code

function token() {
  return new SignJWT({ sid: 'environment-test-session', scope: ['agent:use'] })
    .setProtectedHeader({ alg: 'HS256' }).setIssuer('excelgen-app').setAudience('excelgen-agent')
    .setSubject('environment-test-user').setExpirationTime('5m').sign(new TextEncoder().encode(secret))
}

test('production and unknown environments never allow anonymous or unbilled loopback requests', async () => {
  const signed = await token()
  for (const environment of ['production', undefined, '', 'staging', 'Development', 'development ']) {
    for (const billing of [undefined, 'false', 'true']) {
      const env = { AGENT_ENV: environment, LOCAL_BILLING_ENABLED: billing } as AgentWorkerEnv
      assert.equal(getAgentEnvironment(env), 'production')
      assert.equal(isLocalBillingEnabled(env), false)
      for (const url of [...localUrls, publicUrl]) {
        await assert.rejects(authenticateAgentRequest(new Request(url), env), errorCode('AUTH_NOT_CONFIGURED'))
        const configured = { ...env, AGENT_TOKEN_SECRET: secret }
        await assert.rejects(authenticateAgentRequest(new Request(url), configured), errorCode('UNAUTHORIZED'))
        const identity = await authenticateAgentRequest(new Request(url, {
          headers: { authorization: `Bearer ${signed}` }
        }), configured)
        assert.equal(identity.userId, 'environment-test-user')
        assert.equal(identity.localDemo, false)
        assert.equal(identity.billingEnabled, true)
        assert.equal(identity.scopes.has('agent:eval'), false)
      }
    }
  }
})

test('development still requires a loopback URL and respects the local billing switch', async () => {
  const signed = await token()
  const development = { AGENT_ENV: 'development' } as AgentWorkerEnv
  for (const url of localUrls) {
    const demo = await authenticateAgentRequest(new Request(url), development)
    assert.equal(demo.localDemo, true)
    assert.equal(demo.billingEnabled, false)
    await assert.rejects(authenticateAgentRequest(new Request(url), {
      ...development, LOCAL_BILLING_ENABLED: 'true'
    }), errorCode('AUTH_NOT_CONFIGURED'))
    for (const billing of ['false', 'true']) {
      const configured = { ...development, AGENT_TOKEN_SECRET: secret, LOCAL_BILLING_ENABLED: billing }
      await assert.rejects(authenticateAgentRequest(new Request(url), configured), errorCode('UNAUTHORIZED'))
      const identity = await authenticateAgentRequest(new Request(url, {
        headers: { authorization: `Bearer ${signed}` }
      }), configured)
      assert.equal(identity.localDemo, false)
      assert.equal(identity.billingEnabled, billing === 'true')
    }
  }
  for (const url of [publicUrl, 'http://localhost.evil.test/api/chat', 'http://127.0.0.1.evil.test/api/chat']) {
    await assert.rejects(authenticateAgentRequest(new Request(url), development), errorCode('AUTH_NOT_CONFIGURED'))
    const identity = await authenticateAgentRequest(new Request(url, {
      headers: { authorization: `Bearer ${signed}` }
    }), { ...development, AGENT_TOKEN_SECRET: secret, LOCAL_BILLING_ENABLED: 'false' })
    assert.equal(identity.billingEnabled, true)
  }
})

test('request headers and JSON cannot select development mode or disable billing', async () => {
  const signed = await token()
  const request = new Request(publicUrl, {
    method: 'POST',
    headers: {
      authorization: `Bearer ${signed}`, origin: 'http://localhost:3000', host: 'localhost',
      'x-forwarded-host': '127.0.0.1', 'x-agent-env': 'development', 'content-type': 'application/json'
    },
    body: JSON.stringify({ AGENT_ENV: 'development', billingEnabled: false, localDemo: true })
  })
  const identity = await authenticateAgentRequest(request, {
    AGENT_ENV: 'production', AGENT_TOKEN_SECRET: secret, LOCAL_BILLING_ENABLED: 'false'
  } as AgentWorkerEnv)
  assert.equal(identity.billingEnabled, true)
  assert.equal(identity.localDemo, false)
})

test('production config is explicit and only dev commands override the environment', () => {
  const config = JSON.parse(readFileSync(new URL('../wrangler.jsonc', import.meta.url), 'utf8'))
  const { scripts } = JSON.parse(readFileSync(new URL('../package.json', import.meta.url), 'utf8'))
  assert.equal(config.vars.AGENT_ENV, 'production')
  for (const name of ['dev', 'dev:mock']) {
    assert.match(scripts[name], /^wrangler dev /)
    assert.match(scripts[name], /--var AGENT_ENV:development(?:\s|$)/)
  }
  assert.doesNotMatch(scripts['deploy'], /AGENT_ENV:development/)
})
