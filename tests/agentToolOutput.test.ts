import assert from 'node:assert/strict'
import test from 'node:test'
import { toolOutputErrorMessage } from '#agent/http/toolOutput'

test('recognizes Mastra validation errors carried by tool-output-available', () => {
  assert.equal(
    toolOutputErrorMessage({ error: true, message: 'Tool input validation failed' }),
    'Tool input validation failed'
  )
  assert.equal(toolOutputErrorMessage({ accepted: true }), null)
  assert.equal(toolOutputErrorMessage(null), null)
})
