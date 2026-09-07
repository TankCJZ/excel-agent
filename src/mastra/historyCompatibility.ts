import { ToolCallFilter, type InputProcessor, type ProcessLLMRequestArgs } from '@mastra/core/processors'

// Provider-only compatibility: do not rewrite stored messages or UI history.
// Old tool IDs/results may belong to removed tools or an older provider. The
// built-in filter preserves all current-run calls and removes history pairs.
export function workbookHistoryProcessors(): InputProcessor[] {
  return [new ToolCallFilter({ preserveModelOutput: false }), {
    id: 'excelgen-history-provider-compatibility',
    processLLMRequest({ prompt }: ProcessLLMRequestArgs) {
      let boundary = -1
      for (let index = prompt.length - 1; index >= 0; index--) if (prompt[index]!.role === 'user') { boundary = index; break }
      if (boundary < 0) return
      return { prompt: prompt.flatMap((message, index) => {
        if (index >= boundary || message.role === 'system' || message.role === 'tool') return [message]
        if (message.role === 'user') return [{ role: 'user' as const, content: message.content.map(part => {
          const { providerOptions: _options, ...portable } = part
          return portable
        }) }]
        // Old encrypted reasoning and response item IDs are provider state, not
        // conversation facts. Retain assistant text, omit provider-only parts.
        const content = message.content.filter(part => part.type === 'text').map(part => ({ type: 'text' as const, text: part.text }))
        return content.length ? [{ role: 'assistant' as const, content }] : []
      }) }
    }
  }]
}
