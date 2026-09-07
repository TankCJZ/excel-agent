import assert from 'node:assert/strict'
import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import test from 'node:test'
import { buildWorkbookPlan } from '../src/mastra/plans.ts'
import { summarizeWorkbookSheets } from '../src/workbook/analysis.ts'
import { workbookAnalysisSchema } from '../src/contracts.ts'

const rootDir = resolve(import.meta.dirname, '..')
const runtimeFiles = [
  'mastra/runtime.ts',
  'workbook/analysis.ts',
  'workbook/analysisMerge.ts',
  'contracts.ts',
  'mastra/followups.ts',
  'http/agentStream.ts',
  'mastra/instructions/excelAgent.ts',
  'mastra/tools/excelTools.ts',
  'mastra/models/cloudflareResponses.ts',
  'workbook/planExecutor.ts',
  'spreadsheet/generation.ts',
  'workbook/service.ts'
]

test('agent runtime has no binary Chinese-language routing switch', () => {
  const source = runtimeFiles
    .map(file => readFileSync(resolve(rootDir, 'src', file), 'utf8'))
    .join('\n')

  assert.doesNotMatch(source, /\b(?:isChinese|hasChinese)\b/)
  assert.doesNotMatch(source, /buildGroundedAnalysisText/)
  assert.match(source, /respond in the language of the user's latest request/i)
  assert.match(source, /mixed-language input/)
})

test('analysis and plan protocols have the same semantic shape for Japanese and Spanish requests', () => {
  const summary = summarizeWorkbookSheets([{
    name: 'Sales',
    rows: [['Region', 'Revenue'], ['East', 42]]
  }])
  const questions = ['地域別の売上を分析してください', 'Analiza los ingresos por región']
  const shapes = questions.map(question => {
    const analysis = workbookAnalysisSchema.parse({
      kind: 'analysis',
      question,
      metrics: [{
        label: 'Revenue',
        role: 'requested',
        operation: 'sum',
        field: 'Revenue',
        value: 42,
        format: 'currency'
      }],
      table: { columns: ['Region', 'Revenue'], rows: [['East', 42]] },
      chart: null,
      evidence: [{ sheet: 'Sales', range: 'A1:B2', basis: 'validated_plan' }],
      stats: { sourceRowCount: 1, matchingRowCount: 1, outputRowCount: 1, groupCount: 1 },
      summary: 'legacy prose must not cross the contract',
      findings: [{ title: 'legacy', detail: 'legacy' }]
    })
    const plan = buildWorkbookPlan(analysis, summary)
    assert.ok(!('summary' in analysis))
    assert.ok(!('findings' in analysis))
    assert.ok(!('title' in plan))
    assert.ok(!('rationale' in plan))
    return {
      analysisKeys: Object.keys(analysis).filter(key => key !== 'question').sort(),
      planKeys: Object.keys(plan).sort(),
      operations: plan.steps.map(step => step.operation)
    }
  })

  assert.deepEqual(shapes[0], shapes[1])
})
