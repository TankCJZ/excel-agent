import { z } from 'zod'

export const sessionIdSchema = z.string().uuid()
export const threadIdSchema = z.string().min(8).max(120).regex(/^[a-zA-Z0-9_-]+$/)
export const turnIdSchema = z.string().uuid()
export const taskIdSchema = z.string().uuid()
export const workbookIdSchema = z.string().uuid()
export const adoptTaskWorkbookSchema = z.object({ sessionId: sessionIdSchema, threadId: threadIdSchema, workbookRevision: z.number().int().nonnegative() }).strict()
export const r2KeySchema = z.string().min(1).max(512).startsWith('agent-demo/').refine((key) => !key.includes('..'))

export const cellValueSchema = z.union([z.string(), z.number(), z.boolean(), z.null()])
export const columnTypeSchema = z.enum(['number', 'date', 'text', 'boolean', 'mixed', 'empty'])
export const analysisMetricOperationSchema = z.enum([
  'count',
  'distinct_count',
  'sum',
  'average',
  'median',
  'min',
  'max',
  'blank_count',
  'blank_rate',
  'duplicate_count'
])

export const workbookColumnSchema = z.object({
  name: z.string(),
  index: z.number().int().nonnegative(),
  type: columnTypeSchema,
  nonEmptyCount: z.number().int().nonnegative(),
  nullCount: z.number().int().nonnegative(),
  uniqueCount: z.number().int().nonnegative(),
  numeric: z.object({
    count: z.number().int().nonnegative(),
    sum: z.number(),
    average: z.number(),
    median: z.number(),
    min: z.number(),
    max: z.number()
  }).nullable()
})

export const workbookSheetSchema = z.object({
  name: z.string(),
  rowCount: z.number().int().nonnegative(),
  columnCount: z.number().int().nonnegative(),
  sourceRange: z.string(),
  headers: z.array(z.string()),
  columns: z.array(workbookColumnSchema),
  blankCellCount: z.number().int().nonnegative(),
  duplicateRowCount: z.number().int().nonnegative(),
  preview: z.array(z.array(cellValueSchema))
})

export const workbookSummarySchema = z.object({
  calculation: z.object({ status: z.literal('pending'), formulaCells: z.number().int().nonnegative() }).optional(),
  sheetNames: z.array(z.string()),
  primarySheet: z.string(),
  rowCount: z.number().int().nonnegative(),
  columnCount: z.number().int().nonnegative(),
  headers: z.array(z.string()),
  numericColumns: z.array(z.object({
    name: z.string(),
    count: z.number().int().nonnegative(),
    sum: z.number(),
    average: z.number(),
    median: z.number(),
    min: z.number(),
    max: z.number()
  })),
  blankCellCount: z.number().int().nonnegative(),
  duplicateRowCount: z.number().int().nonnegative(),
  sheets: z.array(workbookSheetSchema),
  preview: z.array(z.array(cellValueSchema))
})

export const analysisEvidenceSchema = z.object({
  sheet: z.string(),
  range: z.string(),
  basis: z.enum(['workbook_structure', 'validated_plan']).default('validated_plan')
})

export const analysisMetricSchema = z.object({
  label: z.string(),
  role: z.enum(['requested', 'group_count', 'matching_rows']).default('requested'),
  operation: analysisMetricOperationSchema,
  field: z.string().nullable().default(null),
  value: z.union([z.string(), z.number()]),
  format: z.enum(['number', 'currency', 'percent', 'text']).default('number')
})

export const analysisTableSchema = z.object({
  columns: z.array(z.string()).max(12),
  rows: z.array(z.array(cellValueSchema)).max(50)
})

export const chartSpecSchema = z.object({
  id: z.string(),
  type: z.enum(['bar', 'line', 'pie', 'scatter']),
  title: z.string(),
  subtitle: z.string().default(''),
  categories: z.array(z.string()).max(50),
  series: z.array(z.object({
    name: z.string(),
    data: z.array(z.number()).max(50)
  })).min(1).max(6),
  xAxisLabel: z.string(),
  yAxisLabel: z.string(),
  source: analysisEvidenceSchema
})

export const dataChartCreationInputSchema = z.object({
  title: z.string().trim().min(2).max(160),
  subtitle: z.string().trim().max(300).default(''),
  type: z.enum(['bar', 'line', 'pie', 'scatter']),
  categoryLabel: z.string().trim().min(1).max(120),
  valueLabel: z.string().trim().min(1).max(120),
  categories: z.array(z.string().trim().min(1).max(160)).min(1).max(50),
  series: z.array(z.object({
    name: z.string().trim().min(1).max(120),
    data: z.array(z.number().finite()).min(1).max(50)
  })).min(1).max(6),
  dataMode: z.enum(['conversation_data', 'model_knowledge', 'web_research'])
}).superRefine((input, context) => {
  for (const [index, series] of input.series.entries()) {
    if (series.data.length !== input.categories.length) {
      context.addIssue({
        code: 'custom',
        path: ['series', index, 'data'],
        message: 'Each chart series must contain one value for every category'
      })
    }
  }
})

export const analysisFilterConditionSchema = z.object({
  field: z.string().trim().min(1).max(120),
  operator: z.enum([
    'eq',
    'neq',
    'gt',
    'gte',
    'lt',
    'lte',
    'contains',
    'not_contains',
    'starts_with',
    'ends_with',
    'in',
    'not_in',
    'is_blank',
    'not_blank',
    'between'
  ]),
  value: cellValueSchema.optional(),
  values: z.array(cellValueSchema).min(1).max(50).optional()
}).superRefine((condition, context) => {
  if (['is_blank', 'not_blank'].includes(condition.operator)) return
  if (['in', 'not_in'].includes(condition.operator) && !condition.values?.length) {
    context.addIssue({
      code: 'custom',
      path: ['values'],
      message: `${condition.operator} requires values`
    })
  }
  if (condition.operator === 'between' && condition.values?.length !== 2) {
    context.addIssue({
      code: 'custom',
      path: ['values'],
      message: 'between requires exactly two values'
    })
  }
  if (!['in', 'not_in', 'between'].includes(condition.operator) && condition.value === undefined) {
    context.addIssue({
      code: 'custom',
      path: ['value'],
      message: `${condition.operator} requires value`
    })
  }
})

export const analysisWhereSchema = z.object({
  logic: z.enum(['and', 'or']).default('and'),
  conditions: z.array(analysisFilterConditionSchema).min(1).max(20)
})

export const analysisPlanMetricSchema = z.object({
  operation: analysisMetricOperationSchema,
  field: z.string().trim().min(1).max(120).optional(),
  alias: z.string().trim().min(1).max(80),
  format: z.enum(['number', 'currency', 'percent', 'text']).default('number')
}).superRefine((metric, context) => {
  if (!['count', 'duplicate_count'].includes(metric.operation) && !metric.field) {
    context.addIssue({
      code: 'custom',
      path: ['field'],
      message: `${metric.operation} requires a field`
    })
  }
})

export const analysisSortSchema = z.object({
  field: z.string().trim().min(1).max(120),
  direction: z.enum(['asc', 'desc']).default('desc')
})

export const analysisChartRequestSchema = z.object({
  type: z.enum(['auto', 'bar', 'line', 'pie', 'scatter']).default('auto'),
  title: z.string().trim().min(1).max(160).optional()
})

export const analysisPlanSchema = z.object({
  goal: z.string().trim().min(2).max(500),
  sheet: z.string().trim().min(1).max(120),
  select: z.array(z.string().trim().min(1).max(120)).max(12).default([]),
  where: analysisWhereSchema.optional(),
  groupBy: z.array(z.string().trim().min(1).max(120)).max(3).default([]),
  metrics: z.array(analysisPlanMetricSchema).max(8).default([]),
  sort: z.array(analysisSortSchema).max(4).default([]),
  limit: z.number().int().min(1).max(50).default(20),
  chart: analysisChartRequestSchema.optional()
}).superRefine((plan, context) => {
  if (!plan.select.length && !plan.metrics.length) {
    context.addIssue({
      code: 'custom',
      path: ['metrics'],
      message: 'A plan requires at least one selected field or metric'
    })
  }
  if (plan.chart && (!plan.groupBy.length || !plan.metrics.length)) {
    context.addIssue({
      code: 'custom',
      path: ['chart'],
      message: 'A chart requires at least one groupBy field and one metric'
    })
  }
  const aliases = new Set<string>()
  for (const [index, metric] of plan.metrics.entries()) {
    const alias = metric.alias.toLowerCase()
    if (aliases.has(alias)) {
      context.addIssue({ code: 'custom', path: ['metrics', index, 'alias'], message: 'Metric aliases must be unique' })
    }
    aliases.add(alias)
  }
})

export const workbookAnalysisSchema = z.object({
  kind: z.enum(['answer', 'analysis', 'chart']),
  question: z.string(),
  metrics: z.array(analysisMetricSchema).max(8),
  table: analysisTableSchema.nullable(),
  chart: chartSpecSchema.nullable(),
  evidence: z.array(analysisEvidenceSchema).min(1).max(8),
  stats: z.object({
    sourceRowCount: z.number().int().nonnegative(),
    matchingRowCount: z.number().int().nonnegative(),
    outputRowCount: z.number().int().nonnegative(),
    groupCount: z.number().int().nonnegative()
  }),
  plan: analysisPlanSchema.optional()
})

export const workbookPlanSchema = z.object({
  action: z.enum(['answer', 'analyze', 'chart']),
  source: z.object({
    sheet: z.string(),
    range: z.string()
  }),
  steps: z.array(z.object({
    id: z.string().min(1).max(60),
    operation: z.enum(['inspect', 'query', 'analyze', 'chart', 'export'])
  })).min(2).max(8)
})

export const webSearchInputSchema = z.object({
  objective: z.string().trim().min(3).max(800),
  queries: z.array(z.string().trim().min(2).max(240)).min(1).max(3),
  recency: z.enum(['day', 'week', 'month', 'year', 'any']).default('any'),
  preferredDomains: z.array(z.string().trim().min(3).max(160)).max(8).default([])
})

export const webSearchSourceSchema = z.object({
  id: z.string().min(1).max(24),
  title: z.string().min(1).max(300),
  url: z.string().url(),
  snippet: z.string().max(800).optional()
})

export const webSearchResultSchema = z.object({
  researchId: z.string().uuid(),
  objective: z.string(),
  summary: z.string(),
  sources: z.array(webSearchSourceSchema).max(40),
  retrievedAt: z.string()
})

export const spreadsheetColumnSchema = z.object({
  key: z.string().trim().min(1).max(80).regex(/^[a-zA-Z][a-zA-Z0-9_]*$/),
  label: z.string().trim().min(1).max(120),
  description: z.string().trim().min(1).max(300),
  type: z.enum(['text', 'number', 'currency', 'percent', 'date', 'url', 'boolean']),
  required: z.boolean().default(false),
  sourceField: z.string().trim().min(1).max(120).optional()
})

const reservedSpreadsheetSheetNames = new Set(['__sources', '__metadata'])
export const spreadsheetSheetNameSchema = z.string()
  .trim()
  .min(1)
  .max(31)
  .refine((name) => !/[:\\/?*\[\]]/.test(name), 'Sheet name contains an invalid Excel character')
  .refine((name) => !name.startsWith("'") && !name.endsWith("'"), 'Sheet name cannot start or end with an apostrophe')
  .refine((name) => !reservedSpreadsheetSheetNames.has(name.toLowerCase()), 'Sheet name is reserved')

export const spreadsheetPlanSchema = z.object({
  title: z.string().trim().min(2).max(160),
  sheetName: spreadsheetSheetNameSchema,
  purpose: z.string().trim().min(3).max(800),
  dataMode: z.enum(['blank_template', 'conversation_data', 'model_knowledge', 'web_research', 'workbook_analysis']),
  columns: z.array(spreadsheetColumnSchema).min(2).max(15),
  maxRows: z.number().int().min(1).max(100).default(50),
  includeSources: z.boolean().default(true)
}).superRefine((plan, context) => {
  const keys = new Set<string>()
  for (const [index, column] of plan.columns.entries()) {
    if (keys.has(column.key)) {
      context.addIssue({ code: 'custom', path: ['columns', index, 'key'], message: 'Column keys must be unique' })
    }
    keys.add(column.key)
    if (plan.dataMode === 'workbook_analysis' && !column.sourceField) {
      context.addIssue({
        code: 'custom',
        path: ['columns', index, 'sourceField'],
        message: 'Workbook-analysis columns require an explicit sourceField'
      })
    }
  }
})

export const spreadsheetFormulaCellSchema = z.object({ type: z.literal('formula'), formula: z.string().min(2).max(2048).startsWith('=') }).strict()
export const spreadsheetCellSchema = z.union([cellValueSchema, spreadsheetFormulaCellSchema])
export const spreadsheetRecordSchema = z.record(z.string().max(80), spreadsheetCellSchema)
export const spreadsheetRecordCellSchema = z.object({
  key: z.string().trim().min(1).max(80).regex(/^[a-zA-Z][a-zA-Z0-9_]*$/),
  value: spreadsheetCellSchema
})
export const spreadsheetRecordInputSchema = z.object({
  cells: z.array(spreadsheetRecordCellSchema).min(1).max(15)
})
export const spreadsheetCreationInputSchema = z.object({
  plan: spreadsheetPlanSchema,
  records: z.array(spreadsheetRecordInputSchema).max(100).default([])
}).superRefine((input, context) => {
  const capturedRecordMode = input.plan.dataMode === 'conversation_data' || input.plan.dataMode === 'model_knowledge'
  if (capturedRecordMode && input.records.length === 0) {
    context.addIssue({ code: 'custom', path: ['records'], message: 'Conversation or model-knowledge data requires at least one captured record' })
  }
  const formulaTemplate = input.plan.dataMode === 'blank_template' && input.records.length > 0 && input.records.every(record => record.cells.every(cell => cell.value === null || (typeof cell.value === 'object' && cell.value.type === 'formula')))
  if (!capturedRecordMode && !formulaTemplate && input.records.length > 0) {
    context.addIssue({ code: 'custom', path: ['records'], message: 'Only conversation or model-knowledge data may include captured records' })
  }
  if (input.records.length > input.plan.maxRows) context.addIssue({ code: 'custom', path: ['records'], message: 'Records exceed the requested row limit' })
  const columns = new Map(input.plan.columns.map(column => [column.key, column]))
  for (const [recordIndex, record] of input.records.entries()) {
    const recordKeys = new Set<string>()
    for (const [cellIndex, cell] of record.cells.entries()) {
      if (recordKeys.has(cell.key)) {
        context.addIssue({
          code: 'custom',
          path: ['records', recordIndex, 'cells', cellIndex, 'key'],
          message: 'Record field is duplicated'
        })
      }
      recordKeys.add(cell.key)
      if (!columns.has(cell.key)) {
        context.addIssue({
          code: 'custom',
          path: ['records', recordIndex, 'cells', cellIndex, 'key'],
          message: 'Record field is not declared in columns'
        })
      }
    }
    for (const column of input.plan.columns) {
      if (column.required && !recordKeys.has(column.key)) {
        context.addIssue({
          code: 'custom',
          path: ['records', recordIndex, 'cells'],
          message: `Required record field is missing: ${column.key}`
        })
      }
    }
  }
})

export const spreadsheetExecutionPlanSchema = z.object({
  action: z.literal('create_spreadsheet'),
  spreadsheet: spreadsheetPlanSchema,
  steps: z.array(z.object({
    id: z.string().min(1).max(60),
    operation: z.enum(['search', 'design', 'extract', 'validate', 'generate'])
  })).min(2).max(8)
})

export const sampleFileRequestSchema = z.object({
  sessionId: sessionIdSchema
})

export const requestedOutcomeSchema = z.enum(['spreadsheet', 'chart', 'analysis'])

export const chatRequestSchema = z.object({
  workbookVersionsV1: z.boolean().optional(),
  workbookMutationV1: z.boolean().optional(),
  workbookRevision: z.number().int().nonnegative().optional(),
  sessionId: sessionIdSchema,
  threadId: threadIdSchema,
  turnId: turnIdSchema,
  retryOfTurnId: turnIdSchema.optional(),
  workbookId: workbookIdSchema.optional(),
  requestedOutcome: requestedOutcomeSchema.optional(),
  message: z.string().trim().min(2).max(4000)
}).refine(input => input.retryOfTurnId !== input.turnId, {
  path: ['retryOfTurnId'],
  message: 'A retry must use a new turn ID'
})

export const conversationCreateSchema = z.object({
  id: threadIdSchema,
  sessionId: sessionIdSchema,
  workbookId: workbookIdSchema.optional()
})

export const conversationUpdateSchema = z.object({
  sessionId: sessionIdSchema,
  workbookRevision: z.number().int().nonnegative().optional(),
  workbookId: workbookIdSchema.nullable().optional(),
  title: z.string().trim().min(1).max(120).optional()
}).refine(input => input.workbookId !== undefined || input.title !== undefined, {
  message: 'At least one conversation field is required'
})

export const conversationListQuerySchema = z.object({
  limit: z.coerce.number().int().min(1).max(50).default(30)
})

export const taskQuerySchema = z.object({
  sessionId: sessionIdSchema
})

export const threadQuerySchema = z.object({
  sessionId: sessionIdSchema
})

export const evalRequestSchema = z.object({
  sessionId: sessionIdSchema
})

export type WorkbookPlan = z.infer<typeof workbookPlanSchema>
export type WorkbookSummary = z.infer<typeof workbookSummarySchema>
export type WorkbookSheet = z.infer<typeof workbookSheetSchema>
export type WorkbookAnalysis = z.infer<typeof workbookAnalysisSchema>
export type ChartSpec = z.infer<typeof chartSpecSchema>
export type DataChartCreationInput = z.infer<typeof dataChartCreationInputSchema>
export type AnalysisPlan = z.infer<typeof analysisPlanSchema>
export type AnalysisPlanMetric = z.infer<typeof analysisPlanMetricSchema>
export type AnalysisFilterCondition = z.infer<typeof analysisFilterConditionSchema>
export type ChatRequest = z.infer<typeof chatRequestSchema>
export type RequestedOutcome = z.infer<typeof requestedOutcomeSchema>
export type ConversationCreate = z.infer<typeof conversationCreateSchema>
export type ConversationUpdate = z.infer<typeof conversationUpdateSchema>
export type WebSearchInput = z.infer<typeof webSearchInputSchema>
export type WebSearchSource = z.infer<typeof webSearchSourceSchema>
export type WebSearchResult = z.infer<typeof webSearchResultSchema>
export type SpreadsheetColumn = z.infer<typeof spreadsheetColumnSchema>
export type SpreadsheetPlan = z.infer<typeof spreadsheetPlanSchema>
export type SpreadsheetRecord = z.infer<typeof spreadsheetRecordSchema>
export type SpreadsheetRecordInput = z.infer<typeof spreadsheetRecordInputSchema>
export type SpreadsheetCreationInput = z.infer<typeof spreadsheetCreationInputSchema>
export type SpreadsheetExecutionPlan = z.infer<typeof spreadsheetExecutionPlanSchema>
