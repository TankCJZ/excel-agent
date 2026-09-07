import type { WorkbookAnalysis } from '#agent/contracts'


export function mergeWorkbookAnalyses(
  current: WorkbookAnalysis | null,
  next: WorkbookAnalysis
): WorkbookAnalysis {
  if (!current) return next

  const question = `${current.question}；${next.question}`
  const metrics = uniqueBy(
    [...current.metrics, ...next.metrics],
    metric => `${metric.role}:${metric.operation}:${metric.field || ''}:${metric.label}`.trim().toLocaleLowerCase()
  ).slice(0, 8)
  const evidence = uniqueBy(
    [...current.evidence, ...next.evidence],
    item => `${item.sheet}:${item.range}`
  ).slice(0, 8)
  const table = tableScore(current.table) >= tableScore(next.table) ? current.table : next.table

  return {
    kind: current.kind === 'chart' || next.kind === 'chart' ? 'chart' : 'analysis',
    question,
    metrics,
    table,
    chart: next.chart || current.chart,
    evidence,
    stats: next.stats,
    plan: next.plan || current.plan
  }
}

function tableScore(table: WorkbookAnalysis['table']) {
  return table ? table.columns.length * Math.max(table.rows.length, 1) : 0
}

function uniqueBy<T>(items: T[], getKey: (item: T) => string) {
  const seen = new Set<string>()
  return items.filter((item) => {
    const key = getKey(item)
    if (seen.has(key)) return false
    seen.add(key)
    return true
  })
}
