import assert from 'node:assert/strict'
import test from 'node:test'
import { workbookHistoryProcessors } from '../src/mastra/historyCompatibility.ts'
import type { ProcessLLMRequestArgs } from '@mastra/core/processors'

test('old provider state and tools are filtered without mutating stored history or current tool calls', async () => {
  const processors = workbookHistoryProcessors()
  const states = processors.map(() => ({}))
  const original = [
    { role: 'user', content: [{ type: 'text', text: 'old question' }] },
    { role: 'assistant', providerOptions: { openai: { itemId: 'old-provider-item' } }, content: [
      { type: 'reasoning', text: '', providerOptions: { openai: { encryptedContent: 'old' } } },
      { type: 'text', text: 'Old answer', providerOptions: { openai: { itemId: 'text-item' } } },
      { type: 'tool-call', toolCallId: 'old-call', toolName: 'removedTool', input: {} }
    ] },
    { role: 'tool', content: [{ type: 'tool-result', toolCallId: 'old-call', toolName: 'removedTool', output: { type: 'json', value: { old: true } } }] },
    { role: 'user', content: [{ type: 'text', text: 'edit the current workbook' }] }
  ] as ProcessLLMRequestArgs['prompt']
  const copy = structuredClone(original)
  async function process(prompt: ProcessLLMRequestArgs['prompt']) {
    for (const [index, processor] of processors.entries()) {
      const result = await processor.processLLMRequest!({ prompt, state: states[index]! } as ProcessLLMRequestArgs)
      prompt = result?.prompt || prompt
    }
    return prompt
  }
  const first = await process(original)
  assert.deepEqual(original, copy)
  assert.ok(JSON.stringify(first).includes('Old answer'))
  assert.ok(!JSON.stringify(first).includes('removedTool'))
  assert.ok(!JSON.stringify(first).includes('old-provider-item'))
  assert.ok(!JSON.stringify(first).includes('encryptedContent'))
  const continuation = [...original,
    { role: 'assistant', content: [{ type: 'tool-call', toolCallId: 'current-call', toolName: 'inspectWorkbookForEditing', input: {} }] },
    { role: 'tool', content: [{ type: 'tool-result', toolCallId: 'current-call', toolName: 'inspectWorkbookForEditing', output: { type: 'json', value: { current: true } } }] }
  ] as ProcessLLMRequestArgs['prompt']
  const next = await process(continuation)
  assert.equal(JSON.stringify(next).match(/current-call/g)?.length, 2)
  assert.deepEqual(original, copy)
})
