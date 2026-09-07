import type {
  SpreadsheetExecutionPlan,
  SpreadsheetPlan,
  WorkbookAnalysis,
  WorkbookPlan,
  WorkbookSummary
} from '#agent/contracts'

export function buildSpreadsheetExecutionPlan(plan: SpreadsheetPlan): SpreadsheetExecutionPlan {
  const steps: SpreadsheetExecutionPlan['steps'] = []
  if (plan.dataMode === 'web_research') {
    steps.push({ id: 'search-sources', operation: 'search' })
  }
  steps.push(
    { id: 'design-columns', operation: 'design' },
    { id: 'extract-records', operation: 'extract' },
    { id: 'validate-records', operation: 'validate' },
    { id: 'generate-workbook', operation: 'generate' }
  )
  return {
    action: 'create_spreadsheet',
    spreadsheet: plan,
    steps
  }
}

export function buildWorkbookPlan(analysis: WorkbookAnalysis, summary: WorkbookSummary): WorkbookPlan {
  const steps: WorkbookPlan['steps'] = [
    { id: 'inspect-source', operation: 'inspect' },
    { id: 'query-data', operation: 'query' },
    { id: 'analyze-result', operation: 'analyze' }
  ]
  if (analysis.chart) {
    steps.push({ id: 'build-chart', operation: 'chart' })
  }
  steps.push({ id: 'export-result', operation: 'export' })

  return {
    action: analysis.chart ? 'chart' : analysis.kind === 'analysis' ? 'analyze' : 'answer',
    source: {
      sheet: summary.primarySheet,
      range: analysis.evidence[0]?.range || summary.sheets[0]?.sourceRange || ''
    },
    steps
  }
}
