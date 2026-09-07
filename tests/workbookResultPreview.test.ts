import assert from 'node:assert/strict'
import test from 'node:test'
import { compactResultSummary, assertWorkflowResultBudget } from '../src/workbook/resultPreview.ts'
import { summarizeWorkbookSheets } from '../src/workbook/analysis.ts'

test('long preview text is bounded without changing underlying workbook context values', () => {
  const originalText = '值'.repeat(32767)
  const summary = summarizeWorkbookSheets([{ name: 'Data', rows: [['Title'], ...Array.from({ length: 100 }, () => [originalText])] }])
  const compact = compactResultSummary(summary)
  assert.equal(summary.preview[1]![0], originalText)
  assert.ok(String(compact.preview[1]![0]).endsWith('…'))
  assert.ok(compact.preview.length <= 20)
  assert.doesNotThrow(() => assertWorkflowResultBudget(compact))
  assert.throws(() => assertWorkflowResultBudget({ huge: '值'.repeat(200000) }), /too complex/)
})
