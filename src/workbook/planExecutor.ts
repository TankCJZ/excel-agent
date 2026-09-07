import {
  prepareWorkbookDataSheets,
  type CellValue,
  type PreparedWorkbookSheet,
  type WorkbookDataSheet
} from '#agent/workbook/analysis'
import {
  analysisPlanSchema,
  type AnalysisFilterCondition,
  type AnalysisPlan,
  type AnalysisPlanMetric,
  type ChartSpec,
  type WorkbookAnalysis
} from '#agent/contracts'

type OutputRow = Record<string, CellValue>

export class AnalysisPlanError extends Error {
  readonly code: 'invalid_plan' | 'unknown_sheet' | 'unknown_field' | 'invalid_operation'

  constructor(code: AnalysisPlanError['code'], message: string) {
    super(message)
    this.name = 'AnalysisPlanError'
    this.code = code
  }
}

export function executeWorkbookAnalysisPlan(
  sheets: WorkbookDataSheet[],
  inputPlan: AnalysisPlan,
  options: { forceChart?: boolean } = {}
): WorkbookAnalysis {
  const parsed = analysisPlanSchema.safeParse(inputPlan)
  if (!parsed.success) {
    throw new AnalysisPlanError('invalid_plan', `Invalid analysis plan: ${parsed.error.issues.map(issue => issue.message).join('; ')}`)
  }

  const plan = parsed.data
  const preparedSheets = prepareWorkbookDataSheets(sheets)
  const sheet = resolveSheet(preparedSheets, plan.sheet)
  const fields = resolvePlanFields(sheet, plan)
  const filteredRows = sheet.rows.filter(row => matchesWhere(row, fields.where, plan.where?.logic || 'and'))
  const output = buildOutputRows(sheet, filteredRows, plan, fields)
  const sorted = sortOutputRows(output.rows, output.columns, plan)
  const limited = sorted.slice(0, plan.limit)
  const evidence = {
    sheet: sheet.name,
    range: sheet.profile.sourceRange,
    basis: 'validated_plan' as const
  }
  const chartRequest = plan.chart || (options.forceChart ? { type: 'auto' as const } : undefined)
  const chart = chartRequest
    ? buildChart(plan, sheet, limited, output.columns, fields, evidence, chartRequest)
    : null

  return {
    kind: chart ? 'chart' : 'analysis',
    question: plan.goal,
    metrics: buildHeadlineMetrics(plan, filteredRows, fields, output.rows.length),
    table: {
      columns: output.columns,
      rows: limited.map(row => output.columns.map(column => row[column] ?? null))
    },
    chart,
    evidence: [evidence],
    stats: {
      sourceRowCount: sheet.profile.rowCount,
      matchingRowCount: filteredRows.length,
      outputRowCount: limited.length,
      groupCount: plan.groupBy.length ? output.rows.length : 0
    },
    plan
  }
}

function resolvePlanFields(sheet: PreparedWorkbookSheet, plan: AnalysisPlan) {
  const select = plan.select.map(field => resolveField(sheet, field))
  const groupBy = plan.groupBy.map(field => resolveField(sheet, field))
  const metrics = plan.metrics.map(metric => ({
    ...metric,
    field: metric.field ? resolveField(sheet, metric.field) : undefined
  }))
  const where = (plan.where?.conditions || []).map(condition => ({
    condition,
    field: resolveField(sheet, condition.field)
  }))

  if (plan.metrics.length && plan.select.length && !plan.groupBy.length) {
    throw new AnalysisPlanError(
      'invalid_plan',
      'select cannot be combined with aggregate metrics unless groupBy is provided; use select for raw rows or metrics for an aggregate result'
    )
  }

  return { select, groupBy, metrics, where }
}

function resolveSheet(sheets: PreparedWorkbookSheet[], requestedName: string) {
  const exact = sheets.find(sheet => sheet.name === requestedName)
  if (exact) return exact
  const normalized = normalizeIdentifier(requestedName)
  const matches = sheets.filter(sheet => normalizeIdentifier(sheet.name) === normalized)
  if (matches.length === 1) return matches[0]!
  throw new AnalysisPlanError(
    'unknown_sheet',
    `Unknown sheet "${requestedName}". Available sheets: ${sheets.map(sheet => sheet.name).join(', ') || '(none)'}`
  )
}

function resolveField(sheet: PreparedWorkbookSheet, requestedName: string) {
  const exact = sheet.profile.columns.find(column => column.name === requestedName)
  if (exact) return exact
  const normalized = normalizeIdentifier(requestedName)
  const matches = sheet.profile.columns.filter(column => normalizeIdentifier(column.name) === normalized)
  if (matches.length === 1) return matches[0]!
  throw new AnalysisPlanError(
    'unknown_field',
    `Unknown field "${requestedName}" in sheet "${sheet.name}". Available fields: ${sheet.headers.join(', ') || '(none)'}`
  )
}

function matchesWhere(
  row: CellValue[],
  conditions: Array<{ condition: AnalysisFilterCondition, field: PreparedWorkbookSheet['profile']['columns'][number] }>,
  logic: 'and' | 'or'
) {
  if (!conditions.length) return true
  const matches = conditions.map(({ condition, field }) => matchesCondition(row[field.index] ?? null, condition))
  return logic === 'or' ? matches.some(Boolean) : matches.every(Boolean)
}

function matchesCondition(actual: CellValue, condition: AnalysisFilterCondition) {
  const expected = condition.value
  switch (condition.operator) {
    case 'is_blank': return isBlank(actual)
    case 'not_blank': return !isBlank(actual)
    case 'eq': return compareValues(actual, expected ?? null) === 0
    case 'neq': return compareValues(actual, expected ?? null) !== 0
    case 'gt': return compareValues(actual, expected ?? null) > 0
    case 'gte': return compareValues(actual, expected ?? null) >= 0
    case 'lt': return compareValues(actual, expected ?? null) < 0
    case 'lte': return compareValues(actual, expected ?? null) <= 0
    case 'contains': return normalizeText(actual).includes(normalizeText(expected ?? null))
    case 'not_contains': return !normalizeText(actual).includes(normalizeText(expected ?? null))
    case 'starts_with': return normalizeText(actual).startsWith(normalizeText(expected ?? null))
    case 'ends_with': return normalizeText(actual).endsWith(normalizeText(expected ?? null))
    case 'in': return (condition.values || []).some(value => compareValues(actual, value) === 0)
    case 'not_in': return !(condition.values || []).some(value => compareValues(actual, value) === 0)
    case 'between': {
      const [lower, upper] = condition.values || []
      return lower !== undefined && upper !== undefined
        && compareValues(actual, lower) >= 0
        && compareValues(actual, upper) <= 0
    }
  }
}

function buildOutputRows(
  sheet: PreparedWorkbookSheet,
  rows: CellValue[][],
  plan: AnalysisPlan,
  fields: ReturnType<typeof resolvePlanFields>
): { columns: string[], rows: OutputRow[] } {
  if (fields.metrics.length) {
    if (fields.groupBy.length) {
      const nonRedundantMetrics = fields.metrics.filter(metric => !isRedundantGroupedCountMetric(metric, fields.groupBy))
      const visibleMetrics = nonRedundantMetrics.length ? nonRedundantMetrics : fields.metrics
      const groups = new Map<string, { values: CellValue[], rows: CellValue[][] }>()
      for (const row of rows) {
        const values = fields.groupBy.map(field => row[field.index] ?? null)
        const key = JSON.stringify(values.map(stableValueKey))
        const group = groups.get(key) || { values, rows: [] }
        group.rows.push(row)
        groups.set(key, group)
      }
      const columns = [...fields.groupBy.map(field => field.name), ...visibleMetrics.map(metric => metric.alias)]
      return {
        columns,
        rows: [...groups.values()].map(group => {
          const output: OutputRow = {}
          fields.groupBy.forEach((field, index) => {
            output[field.name] = group.values[index] ?? null
          })
          visibleMetrics.forEach(metric => {
            output[metric.alias] = calculateMetric(group.rows, metric)
          })
          return output
        })
      }
    }

    const output: OutputRow = {}
    fields.metrics.forEach(metric => {
      output[metric.alias] = calculateMetric(rows, metric)
    })
    return { columns: fields.metrics.map(metric => metric.alias), rows: [output] }
  }

  if (fields.groupBy.length) {
    const columns = fields.groupBy.map(field => field.name)
    const seen = new Set<string>()
    const outputRows = rows.flatMap(row => {
      const values = fields.groupBy.map(field => row[field.index] ?? null)
      const key = JSON.stringify(values.map(stableValueKey))
      if (seen.has(key)) return []
      seen.add(key)
      return [Object.fromEntries(columns.map((column, index) => [column, values[index] ?? null])) as OutputRow]
    })
    return { columns, rows: outputRows }
  }

  const selected = fields.select.length ? fields.select : sheet.profile.columns.slice(0, 12)
  const columns = selected.map(field => field.name)
  return {
    columns,
    rows: rows.map(row => Object.fromEntries(selected.map(field => [field.name, row[field.index] ?? null])) as OutputRow)
  }
}

function calculateMetric(
  rows: CellValue[][],
  metric: Omit<AnalysisPlanMetric, 'field'> & { field?: PreparedWorkbookSheet['profile']['columns'][number] }
): CellValue {
  const values = metric.field ? rows.map(row => row[metric.field!.index] ?? null) : []
  const present = values.filter(value => !isBlank(value))
  switch (metric.operation) {
    case 'count': return metric.field ? present.length : rows.length
    case 'distinct_count': return new Set(present.map(stableValueKey)).size
    case 'blank_count': return values.filter(isBlank).length
    case 'blank_rate': return rows.length ? roundNumber(values.filter(isBlank).length / rows.length) : 0
    case 'duplicate_count': {
      const keys = metric.field
        ? present.map(stableValueKey)
        : rows.map(row => JSON.stringify(row.map(stableValueKey)))
      return keys.length - new Set(keys).size
    }
    case 'sum': return roundNumber(numericValues(values, metric).reduce((total, value) => total + value, 0))
    case 'average': {
      const numbers = numericValues(values, metric)
      return roundNumber(numbers.reduce((total, value) => total + value, 0) / numbers.length)
    }
    case 'median': {
      const numbers = numericValues(values, metric).sort((left, right) => left - right)
      const middle = Math.floor(numbers.length / 2)
      return roundNumber(numbers.length % 2 ? numbers[middle]! : (numbers[middle - 1]! + numbers[middle]!) / 2)
    }
    case 'min': return minOrMaxValue(present, metric, 'min')
    case 'max': return minOrMaxValue(present, metric, 'max')
  }
}

function numericValues(
  values: CellValue[],
  metric: Omit<AnalysisPlanMetric, 'field'> & { field?: PreparedWorkbookSheet['profile']['columns'][number] }
) {
  const numbers = values.flatMap(value => typeof value === 'number' && Number.isFinite(value) ? [value] : [])
  if (!numbers.length) {
    throw new AnalysisPlanError(
      'invalid_operation',
      `${metric.operation} requires a numeric field; "${metric.field?.name || '(none)'}" has no numeric values`
    )
  }
  return numbers
}

function minOrMaxValue(
  values: CellValue[],
  metric: Omit<AnalysisPlanMetric, 'field'> & { field?: PreparedWorkbookSheet['profile']['columns'][number] },
  direction: 'min' | 'max'
) {
  if (!values.length) return null
  return [...values].sort((left, right) => {
    const comparison = compareValues(left, right)
    return direction === 'min' ? comparison : -comparison
  })[0] ?? null
}

function sortOutputRows(rows: OutputRow[], columns: string[], plan: AnalysisPlan) {
  const sort = plan.sort.length
    ? plan.sort
    : plan.groupBy.length && plan.metrics.length
      ? [{ field: plan.metrics[0]!.alias, direction: 'desc' as const }]
      : []
  if (!sort.length) return rows
  const resolvedSort = sort.map(item => {
    const exact = columns.find(column => column === item.field)
    const normalized = exact || columns.find(column => normalizeIdentifier(column) === normalizeIdentifier(item.field))
    if (!normalized) {
      throw new AnalysisPlanError(
        'unknown_field',
        `Unknown sort field "${item.field}". Sortable output fields: ${columns.join(', ')}`
      )
    }
    return { field: normalized, direction: item.direction }
  })
  return [...rows].sort((left, right) => {
    for (const item of resolvedSort) {
      const comparison = compareValues(left[item.field] ?? null, right[item.field] ?? null)
      if (comparison !== 0) return item.direction === 'asc' ? comparison : -comparison
    }
    return 0
  })
}

function buildHeadlineMetrics(
  plan: AnalysisPlan,
  rows: CellValue[][],
  fields: ReturnType<typeof resolvePlanFields>,
  groupCount: number
): WorkbookAnalysis['metrics'] {
  const metrics: WorkbookAnalysis['metrics'] = []
  for (const metric of fields.metrics.slice(0, 7)) {
    const value = calculateMetric(rows, metric)
    if (typeof value !== 'string' && typeof value !== 'number') continue
    metrics.push({
      label: metric.alias,
      role: 'requested',
      operation: metric.operation,
      field: metric.field?.name || null,
      value,
      format: metric.format
    })
  }
  if (plan.groupBy.length) {
    const groupField = plan.groupBy[0] || ''
    const alreadyCountsGroups = metrics.some(metric => (
      ['count', 'distinct_count'].includes(metric.operation)
      && metric.field
      && normalizeIdentifier(metric.field) === normalizeIdentifier(groupField)
    ))
    if (metrics.length < 8 && !alreadyCountsGroups) {
      metrics.push({
        label: groupField,
        role: 'group_count',
        operation: 'distinct_count',
        field: groupField || null,
        value: groupCount,
        format: 'number'
      })
    }
  } else if (!metrics.length || metrics.length < 8) {
    metrics.push({
      label: '',
      role: 'matching_rows',
      operation: 'count',
      field: null,
      value: rows.length,
      format: 'number'
    })
  }
  return metrics.slice(0, 8)
}

function isRedundantGroupedCountMetric(
  metric: ReturnType<typeof resolvePlanFields>['metrics'][number],
  groupBy: ReturnType<typeof resolvePlanFields>['groupBy']
) {
  if (!['count', 'distinct_count'].includes(metric.operation) || !metric.field) return false
  return groupBy.some(field => normalizeIdentifier(field.name) === normalizeIdentifier(metric.field!.name))
}

function buildChart(
  plan: AnalysisPlan,
  sheet: PreparedWorkbookSheet,
  rows: OutputRow[],
  columns: string[],
  fields: ReturnType<typeof resolvePlanFields>,
  evidence: WorkbookAnalysis['evidence'][number],
  request: NonNullable<AnalysisPlan['chart']>
): ChartSpec {
  if (!fields.groupBy.length || !fields.metrics.length) {
    throw new AnalysisPlanError('invalid_plan', 'A chart requires groupBy and at least one metric')
  }
  const metricColumns = fields.metrics.slice(0, 6).map(metric => metric.alias)
  const series = metricColumns.map(column => ({
    name: column,
    data: rows.map(row => {
      const value = row[column]
      if (typeof value !== 'number' || !Number.isFinite(value)) {
        throw new AnalysisPlanError('invalid_operation', `Chart series "${column}" must contain numeric values`)
      }
      return value
    })
  }))
  const categories = rows.map(row => fields.groupBy.map(field => formatValue(row[field.name] ?? null)).join(' / '))
  const type = request.type === 'auto'
    ? inferChartType(fields.groupBy[0]!, categories.length)
    : request.type
  const title = request.title || plan.goal
  return {
    id: `chart-${stableHash(`${sheet.name}:${plan.goal}:${columns.join('|')}`)}`,
    type,
    title,
    subtitle: '',
    categories,
    series,
    xAxisLabel: fields.groupBy.map(field => field.name).join(' / '),
    yAxisLabel: metricColumns.join(' / '),
    source: evidence
  }
}

function inferChartType(field: PreparedWorkbookSheet['profile']['columns'][number], categoryCount: number): ChartSpec['type'] {
  if (field.type === 'date' || /(date|time|month|year|week|day|日期|时间|月份|月|年份|年|周|日)/i.test(field.name)) return 'line'
  if (categoryCount <= 8) return 'bar'
  return 'bar'
}

function compareValues(left: CellValue, right: CellValue) {
  if (isBlank(left) && isBlank(right)) return 0
  if (isBlank(left)) return -1
  if (isBlank(right)) return 1
  if (typeof left === 'number' && typeof right === 'number') return left - right
  if (typeof left === 'boolean' && typeof right === 'boolean') return Number(left) - Number(right)
  const leftDate = parseDateValue(left)
  const rightDate = parseDateValue(right)
  if (leftDate !== null && rightDate !== null) return leftDate - rightDate
  return normalizeText(left).localeCompare(normalizeText(right), undefined, { numeric: true })
}

function parseDateValue(value: CellValue) {
  if (typeof value !== 'string' || !/[\-/:年月日]/.test(value)) return null
  const time = Date.parse(value.replace(/年|月/g, '-').replace(/日/g, ''))
  return Number.isFinite(time) ? time : null
}

function normalizeIdentifier(value: string) {
  return value.normalize('NFKC').trim().toLocaleLowerCase().replace(/[\s_-]+/g, '')
}

function normalizeText(value: CellValue) {
  return String(value ?? '').normalize('NFKC').trim().toLocaleLowerCase()
}

function stableValueKey(value: CellValue) {
  return `${typeof value}:${normalizeText(value)}`
}

function isBlank(value: CellValue): value is null | string {
  return value === null || (typeof value === 'string' && value.trim() === '')
}

function roundNumber(value: number) {
  if (!Number.isFinite(value)) return 0
  return Math.round((value + Number.EPSILON) * 10000) / 10000
}

function formatValue(value: CellValue) {
  if (value === null || value === '') return ''
  if (typeof value === 'number') return String(roundNumber(value))
  return String(value)
}

function stableHash(value: string) {
  let hash = 2166136261
  for (let index = 0; index < value.length; index += 1) {
    hash ^= value.charCodeAt(index)
    hash = Math.imul(hash, 16777619)
  }
  return (hash >>> 0).toString(36)
}
