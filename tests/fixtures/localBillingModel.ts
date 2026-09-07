import { MockLanguageModelV3 } from 'ai/test'
import { createMockExcelModel } from './mockExcelModel.ts'

// Only substituted by agent:dev:mock; never imported by production Worker code.
export const DEFAULT_MODEL_ID = 'mock/excel-planner'

export function resolveWorkersAIModel() {
  const delegate = createMockExcelModel()
  return new MockLanguageModelV3({
    modelId: DEFAULT_MODEL_ID,
    doGenerate: options => delegate.doGenerate(options),
    doStream: async options => {
      const prompt = JSON.stringify(options.prompt.filter(message => message.role === 'user').at(-1))
      if (prompt.includes('[billing-fail]')) throw new Error('Simulated model failure')
      if (prompt.includes('[billing-slow]')) {
        await new Promise<void>((resolve, reject) => {
          const abort = () => { clearTimeout(timer); reject(new DOMException('Cancelled', 'AbortError')) }
          const timer = setTimeout(() => { options.abortSignal?.removeEventListener('abort', abort); resolve() }, 60_000)
          if (options.abortSignal?.aborted) abort()
          else options.abortSignal?.addEventListener('abort', abort, { once: true })
        })
      }
      return delegate.doStream(options)
    }
  })
}
