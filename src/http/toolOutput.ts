export function toolOutputErrorMessage(output: unknown) {
  if (!output || typeof output !== 'object' || !('error' in output) || output.error !== true) return null
  if ('message' in output && typeof output.message === 'string' && output.message.trim()) {
    return output.message.trim()
  }
  return 'Tool execution failed'
}
