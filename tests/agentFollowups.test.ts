import assert from 'node:assert/strict'
import test from 'node:test'
import { createWorkbookFollowups } from '../src/mastra/followups.ts'
import type { WorkbookAnalysis, WorkbookSummary } from '../src/contracts.ts'

const summary: WorkbookSummary = {
  sheetNames: ['处罚汇总'],
  primarySheet: '处罚汇总',
  rowCount: 68,
  columnCount: 3,
  headers: ['姓名', '处罚金额', '处罚类型'],
  numericColumns: [{ name: '处罚金额', count: 7, sum: 500, average: 71.43, median: 50, min: 20, max: 150 }],
  blankCellCount: 0,
  duplicateRowCount: 0,
  sheets: [{
    name: '处罚汇总',
    rowCount: 68,
    columnCount: 3,
    sourceRange: 'A1:C69',
    headers: ['姓名', '处罚金额', '处罚类型'],
    columns: [],
    blankCellCount: 0,
    duplicateRowCount: 0,
    preview: [['姓名', '处罚金额', '处罚类型'], ['张三', 100, '迟到']]
  }],
  preview: [['姓名', '处罚金额', '处罚类型'], ['张三', 100, '迟到']]
}

const analysis: WorkbookAnalysis = {
  kind: 'analysis',
  question: '处罚总金额和人数是多少？',
  metrics: [{
    label: '处罚金额',
    role: 'requested',
    operation: 'sum',
    field: '处罚金额',
    value: 500,
    format: 'currency'
  }],
  table: { columns: ['姓名', '处罚金额'], rows: [['张三', 100]] },
  chart: null,
  evidence: [{ sheet: '处罚汇总', range: 'A1:C69', basis: 'validated_plan' }],
  stats: { sourceRowCount: 68, matchingRowCount: 7, outputRowCount: 7, groupCount: 7 }
}

test('creates semantic follow-ups from verified workbook fields', () => {
  const suggestions = createWorkbookFollowups({ summary, analysis })

  assert.deepEqual(suggestions.map(item => item.kind), ['create_chart', 'extremes', 'export_analysis'])
  assert.deepEqual(suggestions[0]?.params, { dimension: '姓名', metric: '处罚金额' })
  assert.ok(suggestions.every(item => !('prompt' in item)))
})

test('does not recommend creating the same chart again after chart output', () => {
  const chartAnalysis: WorkbookAnalysis = {
    ...analysis,
    kind: 'chart',
    chart: {
      id: 'chart-1',
      type: 'bar',
      title: '按姓名的处罚金额',
      subtitle: '',
      categories: ['张三'],
      series: [{ name: '处罚金额', data: [100] }],
      xAxisLabel: '姓名',
      yAxisLabel: '处罚金额',
      source: analysis.evidence[0]!
    }
  }
  const suggestions = createWorkbookFollowups({ summary, analysis: chartAnalysis })

  assert.deepEqual(suggestions.map(item => item.kind), ['extremes', 'rank_details', 'export_chart'])
  assert.ok(suggestions.every(item => item.intent !== 'chart'))
})

test('suppresses duplicate export suggestions while a workflow export is requested', () => {
  const suggestions = createWorkbookFollowups({ summary, analysis, workflowRequested: true })

  assert.ok(suggestions.every(item => item.intent !== 'export'))
  assert.ok(suggestions.some(item => item.kind === 'check_quality'))
})

test('uses semantic plan state to avoid suggesting an already completed ranking step', () => {
  const ranked: WorkbookAnalysis = {
    ...analysis,
    plan: {
      goal: '按处罚金额排名',
      sheet: '处罚汇总',
      select: [],
      groupBy: ['姓名'],
      metrics: [{ operation: 'sum', field: '处罚金额', alias: '处罚金额', format: 'currency' }],
      sort: [{ field: '处罚金额', direction: 'desc' }],
      limit: 20
    }
  }
  const suggestions = createWorkbookFollowups({ summary, analysis: ranked })

  assert.ok(suggestions.every(item => item.kind !== 'extremes'))
  assert.ok(suggestions.some(item => item.kind === 'rank_details'))
})

test('returns the same language-neutral protocol for non-Chinese workbook fields', () => {
  const suggestions = createWorkbookFollowups({
    summary: {
      ...summary,
      headers: ['Employee', 'Penalty amount', 'Type'],
      numericColumns: [{ ...summary.numericColumns[0]!, name: 'Penalty amount' }]
    }
  })

  assert.deepEqual(suggestions.map(item => item.kind), ['create_chart', 'extremes', 'export_analysis'])
  assert.deepEqual(suggestions[0]?.params, { dimension: 'Employee', metric: 'Penalty amount' })
})
