import assert from 'node:assert/strict'
import { readFileSync } from 'node:fs'
import test from 'node:test'
import { corsHeaders } from '../src/http/responses.ts'

const { scripts } = JSON.parse(readFileSync(new URL('../package.json', import.meta.url), 'utf8'))
const productionConfig = JSON.parse(readFileSync(new URL('../wrangler.jsonc', import.meta.url), 'utf8'))
const localOrigins = ['http://localhost:3000', 'http://127.0.0.1:3000']
const untrustedOrigins = ['http://localhost:3001', 'http://localhost.evil.test:3000', 'https://evil.test', 'null']

function allowedOrigin(origin: string, configuredOrigins: string) {
  const request = new Request('http://127.0.0.1:8791/api/conversations', {
    method: 'OPTIONS',
    headers: {
      origin,
      'access-control-request-method': 'GET',
      'access-control-request-headers': 'authorization'
    }
  })
  return corsHeaders(request, { ALLOWED_ORIGINS: configuredOrigins } as never)
    .get('access-control-allow-origin')
}

for (const scriptName of ['dev', 'dev:mock']) {
  test(`${scriptName} allows local workspace origins without allowing arbitrary origins`, () => {
    const configuredOrigins = scripts[scriptName].match(/--var ALLOWED_ORIGINS:([^\s]+)/)?.[1]
    assert.ok(configuredOrigins, 'Local CORS override must be explicit in the dev command')
    assert.match(scripts[scriptName], /^wrangler dev /)
    for (const origin of localOrigins) {
      assert.equal(allowedOrigin(origin, configuredOrigins), origin)
    }
    for (const origin of untrustedOrigins) {
      assert.equal(allowedOrigin(origin, configuredOrigins), null)
    }
  })
}

test('production CORS allows the production website and configured origins', () => {
  const configuredOrigins = productionConfig.vars.ALLOWED_ORIGINS
  assert.match(configuredOrigins, /https:\/\/www\.excelgen\.app/)
  assert.equal(allowedOrigin('https://www.excelgen.app', configuredOrigins), 'https://www.excelgen.app')
  for (const origin of untrustedOrigins) {
    assert.equal(allowedOrigin(origin, configuredOrigins), null)
  }
  assert.match(scripts['deploy'], /^wrangler deploy$/)
})

test('missing CORS configuration does not implicitly allow localhost', () => {
  for (const origin of localOrigins) {
    assert.equal(allowedOrigin(origin, ''), null)
  }
})
