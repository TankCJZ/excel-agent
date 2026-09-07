import type { WorkbookAnalysis, WorkbookSummary } from '#agent/contracts'


export type FollowupIntent = 'analyze' | 'chart' | 'export' | 'quality'
export type FollowupKind =
  | 'extremes'
  | 'rank_details'
  | 'create_chart'
  | 'count_by_dimension'
  | 'check_quality'
  | 'export_chart'
  | 'export_analysis'
  | 'summarize_findings'
  | 'review_columns'
  | 'check_sources'
  | 'compare_findings'
  | 'create_research_sheet'
  | 'rank_sources'

export type FollowupSuggestion = {
  id: string
  intent: FollowupIntent
  kind: FollowupKind
  params: Record<string, string>
}

type FollowupInput = {
  summary: WorkbookSummary
  analysis?: WorkbookAnalysis | null
  workflowRequested?: boolean
}

/**
 * Builds safe next-step prompts from verified workbook metadata and analysis output.
 * This deliberately avoids another model call so suggestions do not add latency and
 * never invent fields that are absent from the workbook.
 */
export function createWorkbookFollowups(input: FollowupInput): FollowupSuggestion[] {
  const { summary, analysis, workflowRequested = false } = input
  const numericNames = new Set(summary.numericColumns.map(column => normalize(column.name)))
  const plan = analysis?.plan
  const dimension = firstDefined([
    plan?.groupBy[0],
    analysis?.chart?.xAxisLabel,
    analysis?.table?.columns.find(column => !numericNames.has(normalize(column))),
    summary.headers.find(header => !numericNames.has(normalize(header))),
    summary.headers[0]
  ])
  const metric = firstDefined([
    plan?.metrics[0]?.field,
    plan?.metrics[0]?.alias,
    analysis?.chart?.yAxisLabel,
    analysis?.chart?.series[0]?.name,
    summary.numericColumns[0]?.name
  ])
  const alreadyRanked = Boolean(plan?.sort.length)
  const alreadyCheckedQuality = Boolean(plan?.metrics.some(item => (
    ['blank_count', 'blank_rate', 'duplicate_count'].includes(item.operation)
  )))

  const suggestions: FollowupSuggestion[] = []
  const push = (intent: FollowupIntent, kind: FollowupKind, params: Record<string, string> = {}) => {
    if (suggestions.some(item => item.kind === kind)) return
    suggestions.push({ id: `${intent}-${suggestions.length + 1}`, intent, kind, params })
  }

  if (analysis?.chart) {
    if (dimension && metric) {
      if (!alreadyRanked) push('analyze', 'extremes', { dimension, metric })
      push('analyze', 'rank_details', { dimension, metric })
    }
  } else if (dimension && metric) {
    push('chart', 'create_chart', { dimension, metric })
    push('analyze', alreadyRanked ? 'rank_details' : 'extremes', { dimension, metric })
  } else if (dimension) {
    push('analyze', 'count_by_dimension', { dimension })
    if (!alreadyCheckedQuality) push('quality', 'check_quality')
  }

  if (!workflowRequested) {
    push('export', analysis?.chart ? 'export_chart' : 'export_analysis')
  }

  if (suggestions.length < 3 && !alreadyCheckedQuality) push('quality', 'check_quality')
  if (suggestions.length < 3) push('analyze', 'summarize_findings')

  return suggestions.slice(0, 3)
}

function firstDefined(values: Array<string | undefined>) {
  for (const value of values) {
    const trimmed = value?.trim()
    if (trimmed) return trimmed
  }
  return undefined
}

function normalize(value: string) {
  return value.trim().toLocaleLowerCase().replace(/[\s“”"'「」：:，,。.!！？?、/_-]+/gu, '')
}
