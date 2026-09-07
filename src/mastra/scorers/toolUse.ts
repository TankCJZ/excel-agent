import { createScorer } from '@mastra/core/evals'

export const appropriateToolUseScorer = createScorer({
  id: 'excel-appropriate-tool-use',
  name: 'Excel appropriate tool use',
  description: 'Checks that workbook inspection and one action tool form a safe, non-redundant trajectory.',
  type: 'trajectory'
}).generateScore(({ run }) => {
  const toolNames = (run.output?.steps || [])
    .filter((step) => step.stepType === 'tool_call')
    .map((step) => normalizeToolName(step.name))
  const workbookTools = toolNames.filter((name) => ['inspectworkbook', 'analyzeworkbook', 'createchart', 'exportanalysisworkbook'].includes(name))
  const inspectIndexes = workbookTools.flatMap((name, index) => name === 'inspectworkbook' ? [index] : [])
  const actionIndexes = workbookTools.flatMap((name, index) => name !== 'inspectworkbook' ? [index] : [])
  if (inspectIndexes.length > 1 || actionIndexes.length > 1) return 0
  if (inspectIndexes.length && actionIndexes.length && inspectIndexes[0]! > actionIndexes[0]!) return 0
  return 1
})

function normalizeToolName(value: string) {
  return value.replace(/[^a-zA-Z]/g, '').toLowerCase()
}
