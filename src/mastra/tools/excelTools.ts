import type { ToolsInput } from '@mastra/core/agent'
import { createTool } from '@mastra/core/tools'
import { z } from 'zod'
import {
  analysisPlanSchema,
  chartSpecSchema,
  dataChartCreationInputSchema,
  spreadsheetCreationInputSchema,
  taskIdSchema,
  webSearchInputSchema,
  webSearchResultSchema,
  workbookAnalysisSchema,
  workbookSummarySchema,
  type SpreadsheetPlan,
  type SpreadsheetRecord,
  type WebSearchResult,
  type WorkbookAnalysis
} from '#agent/contracts'
import { mergeWorkbookAnalyses } from '#agent/workbook/analysisMerge'
import { getTaskForSession, updateWorkbookContextCache } from '#agent/persistence/agentRepository'
import type { ExcelAgentRuntimeInput } from '#agent/mastra/runtimeTypes'
import { executeWorkbookAnalysisPlan } from '#agent/workbook/planExecutor'
import type { AgentWorkerEnv } from '#agent/types'
import { loadWorkbookContext, type WorkbookAnalysisContext } from '#agent/workbook/service'
import { compactWebSearchResultForAgent, searchWeb } from '#agent/research/webSearch'
import { assertFreeActionAvailable } from '#agent/billing/billingRepository'
import { createWorkbookEditingTools } from '#agent/mastra/tools/workbookEditing'
import { requireMutation } from '#agent/workbook/mutation/errors'
import { assertWorkbookCalculated } from '#agent/workbook/formulaCalculation'
import { spreadsheetFormulas } from '#agent/spreadsheet/formulas'

export type ExcelAgentToolState = {
  billingError: Error | null
  context: WorkbookAnalysisContext | null
  analysis: WorkbookAnalysis | null
  workflowRequested: boolean
  research: WebSearchResult | null
  researchSnapshotR2Key: string | null
  researchSource: 'initial' | 'external' | null
  webSearchFallback: WebSearchFallback | null
  spreadsheetUsesModelKnowledge: boolean
  spreadsheetPlan: SpreadsheetPlan | null
  spreadsheetRecords: SpreadsheetRecord[]
}

export type WebSearchFallback = {
  status: 'unavailable'
  objective: string
  summary: string
  sources: []
  retrievedAt: string
  continueWithoutWeb: true
  recommendedSpreadsheetMode: 'model_knowledge'
}

export function createExcelAgentTools(env: AgentWorkerEnv, input: ExcelAgentRuntimeInput) {
  const state: ExcelAgentToolState = {
    billingError: null,
    context: input.initialContext || null,
    analysis: null,
    workflowRequested: false,
    research: input.initialResearch || null,
    researchSnapshotR2Key: input.initialResearchSnapshotR2Key || null,
    researchSource: input.initialResearch ? 'initial' : null,
    webSearchFallback: null,
    spreadsheetUsesModelKnowledge: false,
    spreadsheetPlan: null,
    spreadsheetRecords: []
  }
  const editing = createWorkbookEditingTools(env, input, () => {
    requireMutation(!state.workflowRequested && !state.spreadsheetPlan, 'OUTPUT_ALREADY_PLANNED', 'A workbook output is already planned. Request editing in the next turn.')
  })

  const assertActionAvailable = async (action: 'createChart' | 'searchWeb') => {
    if (input.billingEnabled === false || !env.DB || !input.userId || input.userId === 'local-demo') return
    try {
      await assertFreeActionAvailable(env, input.userId, action)
    } catch (error) {
      // Mastra can convert tool exceptions into model-visible results. Billing
      // denials must also stop HTTP settlement, even if the model keeps answering.
      state.billingError = error instanceof Error ? error : new Error('Billing check failed')
      throw error
    }
  }

  const getContext = async () => {
    if (state.context) return state.context
    if (!input.workbookId || !input.fileKey || !input.summary) {
      throw new Error('This conversation does not have an attached workbook')
    }

    const loaded = await loadWorkbookContext(env.WORKBOOKS, {
      workbookId: input.workbookId,
      sessionId: input.sessionId,
      fileKey: input.fileKey,
      contextR2Key: input.contextR2Key || null,
      contextVersion: input.contextVersion ?? null,
      sourceEtag: input.sourceEtag || null
    })
    state.context = loaded.context

    if (loaded.cacheStatus === 'repaired' && input.persistContextMetadata !== false) {
      try {
        await updateWorkbookContextCache(env, {
          workbookId: input.workbookId,
          userId: input.userId || 'local-demo',
          sessionId: input.sessionId,
          contextR2Key: loaded.contextR2Key,
          contextVersion: loaded.contextVersion,
          sourceEtag: loaded.sourceEtag,
          parsedAt: loaded.parsedAt,
          summary: loaded.context.summary
        })
      } catch (error) {
        console.error('Workbook context metadata could not be updated in D1', error)
      }
    }

    return state.context
  }

  const inspectWorkbookTool = createTool({
    id: 'inspect-workbook',
    description: 'Read the current workbook structure, headers, column profiles, and a bounded preview. Use this for questions about what the workbook contains or how it is organized. rowCount is already the number of data records and excludes header and detected aggregate footer rows; preview includes headers as its first row. Do not use this tool for casual conversation.',
    inputSchema: z.object({}),
    outputSchema: workbookSummarySchema,
    execute: async () => {
      if (!input.summary) throw new Error('This conversation does not have an attached workbook')
      return (await getContext()).summary
    }
  })

  const analyzeWorkbookTool = createTool({
    id: 'analyze-workbook',
    description: 'Execute a validated, deterministic analysis plan over the current workbook. First inspect the workbook unless the exact sheet and field names are already available in the conversation. The plan may select rows, filter, group, aggregate, sort, and limit results, but cannot contain code or SQL. Use exact sheet and field names from inspectWorkbook.',
    inputSchema: z.object({ plan: analysisPlanSchema }),
    outputSchema: workbookAnalysisSchema,
    execute: async ({ plan }) => {
      const context = await getContext()
      assertWorkbookCalculated(context.summary)
      const analysis = executeWorkbookAnalysisPlan(context.sheets, plan)
      state.analysis = mergeWorkbookAnalyses(state.analysis, analysis)
      return analysis
    }
  })

  const createChartTool = createTool({
    id: 'create-chart',
    description: 'Execute a validated aggregation plan and build a chart from its deterministic result. Use only when the user explicitly asks for a chart. First inspect the workbook unless exact schema is already in the conversation. The plan must use exact field names, groupBy, at least one numeric metric, and a chart request.',
    inputSchema: z.object({ plan: analysisPlanSchema }),
    outputSchema: chartSpecSchema,
    execute: async ({ plan }) => {
      await assertActionAvailable('createChart')
      const context = await getContext()
      assertWorkbookCalculated(context.summary)
      const analysis = executeWorkbookAnalysisPlan(context.sheets, plan, { forceChart: true })
      if (!analysis.chart) throw new Error('The workbook does not contain enough structured data to create a chart')
      state.analysis = mergeWorkbookAnalyses(state.analysis, analysis)
      return analysis.chart
    }
  })

  const createDataChartTool = createTool({
    id: 'create-data-chart',
    description: 'Build a safe chart from bounded structured data that is already present in the conversation, returned by searchWeb, or responsibly recalled as stable model knowledge. Use this when the requested chart does not depend on an attached workbook. Never invent current facts, URLs, or unsupported numeric values. This tool accepts data only and cannot execute ECharts options, code, formulas, or SQL.',
    inputSchema: dataChartCreationInputSchema,
    outputSchema: chartSpecSchema,
    execute: async (chartInput) => {
      await assertActionAvailable('createChart')
      if (chartInput.dataMode === 'web_research' && !state.research) {
        throw new Error('Run searchWeb successfully before creating a web-research chart')
      }
      const sourceSheet = chartInput.dataMode === 'web_research'
        ? 'Web research'
        : chartInput.dataMode === 'model_knowledge'
          ? 'Model knowledge'
          : 'Conversation data'
      const source = {
        sheet: sourceSheet,
        range: `A1:${spreadsheetColumnLetter(chartInput.series.length + 1)}${chartInput.categories.length + 1}`,
        basis: 'validated_plan' as const
      }
      const chart = chartSpecSchema.parse({
        id: crypto.randomUUID(),
        type: chartInput.type,
        title: chartInput.title,
        subtitle: chartInput.subtitle,
        categories: chartInput.categories,
        series: chartInput.series,
        xAxisLabel: chartInput.categoryLabel,
        yAxisLabel: chartInput.valueLabel,
        source
      })
      const table = {
        columns: [chartInput.categoryLabel, ...chartInput.series.map(series => series.name)],
        rows: chartInput.categories.map((category, categoryIndex) => [
          category,
          ...chartInput.series.map(series => series.data[categoryIndex] ?? null)
        ])
      }
      const analysis = workbookAnalysisSchema.parse({
        kind: 'chart',
        question: chartInput.title,
        metrics: [],
        table,
        chart,
        evidence: [source],
        stats: {
          sourceRowCount: chartInput.categories.length,
          matchingRowCount: chartInput.categories.length,
          outputRowCount: chartInput.categories.length,
          groupCount: chartInput.categories.length
        }
      })
      state.analysis = mergeWorkbookAnalyses(state.analysis, analysis)
      return chart
    }
  })

  const exportAnalysisWorkbookTool = createTool({
    id: 'export-analysis-workbook',
    description: 'Execute a validated analysis plan, then generate a new workbook copy containing result and optional chart-data sheets named from the model-authored analysis content. First inspect the workbook unless exact schema is already in the conversation. This never edits the source workbook and cannot apply arbitrary cell changes, cleanup operations, or formulas.',
    inputSchema: z.object({ plan: analysisPlanSchema }),
    outputSchema: z.object({
      accepted: z.literal(true),
      analysis: workbookAnalysisSchema
    }),
    execute: async ({ plan }) => {
      requireMutation(!editing.getPlan(), 'OUTPUT_ALREADY_PLANNED', 'Workbook editing is already planned for this turn.')
      const context = await getContext()
      assertWorkbookCalculated(context.summary)
      const analysis = executeWorkbookAnalysisPlan(context.sheets, plan)
      state.analysis = mergeWorkbookAnalyses(state.analysis, analysis)
      state.workflowRequested = true
      return { accepted: true as const, analysis }
    }
  })

  const queryTaskTool = createTool({
    id: 'query-workbook-task',
    description: 'Query the latest status and events for an Excel Agent task stored in D1.',
    inputSchema: z.object({ taskId: taskIdSchema }),
    outputSchema: z.object({
      found: z.boolean(),
      status: z.string().nullable(),
      eventCount: z.number().int().nonnegative()
    }),
    execute: async ({ taskId }) => {
      const record = await getTaskForSession(env, taskId, input.userId || 'local-demo', input.sessionId)
      return {
        found: Boolean(record),
        status: record?.task.status || null,
        eventCount: record?.events.length || 0
      }
    }
  })

  const searchWebTool = createTool({
    id: 'search-web',
    description: 'Search the live web using the configured Responses API model with web_search. Use this only when the answer depends on current or external information, source verification, market research, recent facts, or data that is not contained in the attached workbook or conversation. Design up to three focused queries. Do not use it for casual conversation or calculations fully answerable from an attached workbook.',
    inputSchema: webSearchInputSchema,
    outputSchema: z.union([
      webSearchResultSchema,
      z.object({
        status: z.literal('unavailable'),
        objective: z.string(),
        summary: z.string(),
        sources: z.tuple([]),
        retrievedAt: z.string(),
        continueWithoutWeb: z.literal(true),
        recommendedSpreadsheetMode: z.literal('model_knowledge')
      })
    ]),
    execute: async (searchInput) => {
      if (state.research) return compactWebSearchResultForAgent(state.research)
      if (state.webSearchFallback) return state.webSearchFallback
      await assertActionAvailable('searchWeb')
      try {
        const searched = await searchWeb(env, { ...searchInput, sessionId: input.sessionId })
        state.research = searched.result
        state.researchSnapshotR2Key = searched.snapshotR2Key
        state.researchSource = 'external'
        state.webSearchFallback = null
        return compactWebSearchResultForAgent(searched.result)
      } catch (error) {
        console.error('Web search is unavailable; continuing without live web evidence', {
          message: error instanceof Error ? error.message : 'Unknown web-search error'
        })
        const fallback = createWebSearchFallback(searchInput.objective)
        state.webSearchFallback = fallback
        state.research = null
        state.researchSnapshotR2Key = null
        state.researchSource = null
        return fallback
      }
    }
  })

  const createSpreadsheetTool = createTool({
    id: 'create-spreadsheet',
    description: 'Approve a structured plan for creating a new downloadable Excel workbook. Use only when the user asks to create, generate, export, or build a table/spreadsheet. You must autonomously design useful generic columns from the user goal. Choose web_research only after searchWeb has returned evidence, workbook_analysis only after a workbook analysis, conversation_data when the conversation already contains the data, model_knowledge with non-empty records when live search was unavailable, and blank_template only when the user actually requested an empty reusable template. Represent each captured row as cells containing one key and value for every declared required column. This tool queues generation; it does not answer ordinary research questions by itself.',
    inputSchema: spreadsheetCreationInputSchema,
    outputSchema: z.object({
      accepted: z.literal(true),
      title: z.string(),
      dataMode: spreadsheetCreationInputSchema.shape.plan.shape.dataMode,
      columnCount: z.number().int(),
      maxRows: z.number().int(),
      degraded: z.boolean().optional()
    }),
    execute: async ({ plan, records: recordInputs }) => {
      requireMutation(!editing.getPlan(), 'OUTPUT_ALREADY_PLANNED', 'Workbook editing is already planned for this turn.')
      if (plan.dataMode === 'web_research' && !state.research) {
        if (state.webSearchFallback) {
          throw new Error('Live search is unavailable. Retry createSpreadsheet with dataMode model_knowledge and at least one responsibly recalled record; do not create an empty template.')
        }
        throw new Error('Run searchWeb before creating a web-research spreadsheet')
      }
      if (plan.dataMode === 'workbook_analysis' && !state.analysis) {
        throw new Error('Run a workbook analysis before creating a workbook-analysis spreadsheet')
      }
      const effectivePlan = plan.dataMode === 'model_knowledge'
        ? { ...plan, includeSources: false }
        : plan
      const records = recordInputs.map(record => Object.fromEntries(
        record.cells.map(cell => [cell.key, cell.value])
      ))
      spreadsheetFormulas(effectivePlan, records)
      state.spreadsheetPlan = effectivePlan
      state.spreadsheetRecords = records
      state.spreadsheetUsesModelKnowledge = effectivePlan.dataMode === 'model_knowledge'
      return {
        accepted: true as const,
        title: effectivePlan.title,
        dataMode: effectivePlan.dataMode,
        columnCount: effectivePlan.columns.length,
        maxRows: effectivePlan.maxRows,
        ...(effectivePlan.dataMode === 'model_knowledge' ? { degraded: true } : {})
      }
    }
  })

  const tools: ToolsInput = {
    searchWeb: searchWebTool,
    createSpreadsheet: createSpreadsheetTool,
    createDataChart: createDataChartTool,
    queryTask: queryTaskTool
  }
  if (input.summary && input.workbookId && input.fileKey) {
    Object.assign(tools, {
      inspectWorkbook: inspectWorkbookTool,
      analyzeWorkbook: analyzeWorkbookTool,
      createChart: createChartTool,
      exportAnalysisWorkbook: exportAnalysisWorkbookTool
    })
    if (input.supportsWorkbookMutation && env.WORKBOOK_EDITING_ENABLED === 'true') Object.assign(tools, editing.tools)
  }

  return {
    tools,
    state,
    getMutationPlan: editing.getPlan,
    getToolDetail: (toolName: string) => toolName === 'editWorkbook' ? editing.getChanges() || undefined : getToolDetail(toolName, state)
  }
}

function getToolDetail(toolName: string, state: ExcelAgentToolState) {
  if (toolName === 'searchWeb' && state.webSearchFallback) {
    return {
      skipped: true,
      fallback: 'continue_without_web',
      objective: state.webSearchFallback.objective
    }
  }
  if (toolName === 'searchWeb' && state.research) {
    return {
      sourceCount: state.research.sources.length,
      objective: state.research.objective,
      reused: state.researchSource === 'initial'
    }
  }
  if (toolName === 'createSpreadsheet' && state.spreadsheetPlan) {
    return {
      title: state.spreadsheetPlan.title,
      dataMode: state.spreadsheetPlan.dataMode,
      columnCount: state.spreadsheetPlan.columns.length,
      rowCount: state.spreadsheetPlan.maxRows,
      ...(state.spreadsheetUsesModelKnowledge ? { degraded: true, fallback: 'model_knowledge', liveVerified: false } : {})
    }
  }
  if (toolName === 'createDataChart' && state.analysis?.chart) {
    return {
      title: state.analysis.chart.title,
      rowCount: state.analysis.chart.categories.length,
      seriesCount: state.analysis.chart.series.length,
      sheet: state.analysis.chart.source.sheet,
      range: state.analysis.chart.source.range
    }
  }
  return undefined
}

function spreadsheetColumnLetter(columnCount: number) {
  let value = Math.max(1, columnCount)
  let result = ''
  while (value > 0) {
    value -= 1
    result = String.fromCharCode(65 + (value % 26)) + result
    value = Math.floor(value / 26)
  }
  return result
}

export function createWebSearchFallback(objective: string, retrievedAt = new Date().toISOString()): WebSearchFallback {
  return {
    status: 'unavailable',
    objective,
    summary: 'Live web search is temporarily unavailable. Continue with stable existing knowledge and clearly state that current facts could not be verified.',
    sources: [],
    retrievedAt,
    continueWithoutWeb: true,
    recommendedSpreadsheetMode: 'model_knowledge'
  }
}
