import assert from 'node:assert/strict'
import test from 'node:test'
import * as XLSX from 'xlsx'
import { chatRequestSchema, dataChartCreationInputSchema, spreadsheetPlanSchema } from '../src/contracts.ts'
import { buildExcelAgentInstructions } from '../src/mastra/instructions/excelAgent.ts'
import { buildSpreadsheetExecutionPlan } from '../src/mastra/plans.ts'
import { createExcelAgentTools } from '../src/mastra/tools/excelTools.ts'
import { buildSpreadsheetWorkbook, sanitizeSpreadsheetCell } from '../src/spreadsheet/generation.ts'
import {
  AGENT_WEB_SEARCH_SNIPPET_CHAR_LIMIT,
  AGENT_WEB_SEARCH_SOURCE_LIMIT,
  AGENT_WEB_SEARCH_SUMMARY_CHAR_LIMIT,
  compactWebSearchResultForAgent,
  normalizeWebSearchResponse,
  runResponsesModel
} from '../src/research/webSearch.ts'
import { createSampleWorkbook, parseWorkbookBytes } from '../src/workbook/service.ts'
import { attachmentContentDisposition } from '../src/http/fileDownload.ts'

const spreadsheetPlan = {
  title: 'AI tools research',
  sheetName: 'Research',
  purpose: 'Compare current AI spreadsheet tools',
  dataMode: 'web_research' as const,
  columns: [
    { key: 'name', label: 'Name', description: 'Product name', type: 'text' as const, required: true },
    { key: 'website', label: 'Website', description: 'Official product URL', type: 'url' as const, required: false }
  ],
  maxRows: 20,
  includeSources: true
}

test('creates the sample workbook with English-only sheet content', async () => {
  const sample = await createSampleWorkbook()
  const parsed = await parseWorkbookBytes(sample.bytes)

  assert.equal(parsed.summary.primarySheet, 'Sales Data')
  assert.deepEqual(parsed.summary.headers, ['Month', 'Region', 'Sales', 'Cost', 'Orders'])
  assert.deepEqual(sample.preview[1], ['January', 'North', 12800, 7800, 84])
  assert.doesNotMatch(JSON.stringify({ preview: sample.preview, sheets: parsed.summary.sheets }), /\p{Script=Han}/u)
})

test('chat accepts a question without a workbook and validates autonomous spreadsheet plans', () => {
  const chat = chatRequestSchema.parse({
    sessionId: '0e2e9e96-b44f-4aee-a4fd-d33be07383d2',
    threadId: 'thread_research_01',
    turnId: '11111111-2222-4333-8444-555555555555',
    message: 'Research current AI spreadsheet tools'
  })
  assert.equal(chat.workbookId, undefined)

  const chartRequest = chatRequestSchema.parse({
    sessionId: '0e2e9e96-b44f-4aee-a4fd-d33be07383d2',
    threadId: 'thread_research_01',
    turnId: '21111111-2222-4333-8444-555555555555',
    requestedOutcome: 'chart',
    message: 'Compare the major categories'
  })
  assert.equal(chartRequest.requestedOutcome, 'chart')
  assert.match(buildExcelAgentInstructions(false, 'chart'), /Deliver a chart when the required data is available/)
  assert.match(buildExcelAgentInstructions(false, 'chart'), /use createDataChart/)
  assert.match(buildExcelAgentInstructions(true, 'spreadsheet'), /appropriate creation, editing, or export tool/)
  assert.match(buildExcelAgentInstructions(true, 'spreadsheet'), /Never invent or print download links/)
  assert.match(buildExcelAgentInstructions(true, 'spreadsheet'), /sandbox:/)
  assert.match(buildExcelAgentInstructions(true, 'analysis'), /Do not create or export a spreadsheet/)

  const parsed = spreadsheetPlanSchema.parse(spreadsheetPlan)
  assert.equal(parsed.columns.length, 2)
  assert.throws(() => spreadsheetPlanSchema.parse({
    ...spreadsheetPlan,
    columns: [spreadsheetPlan.columns[0], { ...spreadsheetPlan.columns[1], key: 'name' }]
  }))
})

test('creates a validated chart without requiring an uploaded workbook', async () => {
  const runtime = createExcelAgentTools({} as never, {
    userId: 'test-user',
    sessionId: '49e6551b-5d94-4b58-b52b-d93e48b5ae89',
    billingEnabled: false,
    requestedOutcome: 'chart'
  })
  const chartInput = dataChartCreationInputSchema.parse({
    title: 'Quarterly revenue',
    subtitle: 'Conversation-provided figures',
    type: 'bar',
    categoryLabel: 'Quarter',
    valueLabel: 'Revenue',
    categories: ['Q1', 'Q2', 'Q3'],
    series: [{ name: 'Revenue', data: [120, 145, 168] }],
    dataMode: 'conversation_data'
  })

  const result = await runtime.tools.createDataChart!.execute!(chartInput, {} as never)
  assert.equal((result as { title: string }).title, 'Quarterly revenue')
  assert.deepEqual(runtime.state.analysis?.table?.rows, [
    ['Q1', 120],
    ['Q2', 145],
    ['Q3', 168]
  ])
  assert.equal(runtime.state.analysis?.chart?.source.sheet, 'Conversation data')
  assert.deepEqual(runtime.getToolDetail('createDataChart'), {
    title: 'Quarterly revenue',
    rowCount: 3,
    seriesCount: 1,
    sheet: 'Conversation data',
    range: 'A1:B4'
  })

  assert.throws(() => dataChartCreationInputSchema.parse({
    ...chartInput,
    series: [{ name: 'Revenue', data: [120] }]
  }), /one value for every category/)
})

test('normalizes OpenAI Responses web-search output and deduplicates source URLs', () => {
  const result = normalizeWebSearchResponse({
    output_text: 'Two products were verified from their official websites.',
    output: [
      {
        type: 'web_search_call',
        action: {
          sources: [
            { title: 'Product A', url: 'https://example.com/a?utm_source=test' },
            { title: 'Product A duplicate', url: 'https://example.com/a' }
          ]
        }
      },
      {
        type: 'message',
        content: [{ type: 'output_text', text: 'Answer', annotations: [{ title: 'Product B', url: 'https://example.org/b' }] }]
      }
    ]
  }, {
    researchId: '49e6551b-5d94-4b58-b52b-d93e48b5ae89',
    objective: 'Compare products',
    retrievedAt: '2026-08-25T00:00:00.000Z'
  })

  assert.equal(result.summary, 'Two products were verified from their official websites.')
  assert.deepEqual(result.sources.map(source => source.id), ['S1', 'S2'])
  assert.deepEqual(result.sources.map(source => source.url), ['https://example.com/a', 'https://example.org/b'])
})

test('keeps the full research snapshot while compacting the result returned to the Agent', () => {
  const research = {
    researchId: '49e6551b-5d94-4b58-b52b-d93e48b5ae89',
    objective: 'Compare products',
    summary: 'A'.repeat(AGENT_WEB_SEARCH_SUMMARY_CHAR_LIMIT + 100),
    sources: Array.from({ length: 20 }, (_, index) => ({
      id: `S${index + 1}`,
      title: `Source ${index + 1}`,
      url: `https://example.com/${index + 1}`,
      snippet: 'B'.repeat(AGENT_WEB_SEARCH_SNIPPET_CHAR_LIMIT + 100)
    })),
    retrievedAt: '2026-08-25T00:00:00.000Z'
  }

  const compact = compactWebSearchResultForAgent(research)

  assert.equal(compact.summary.length, AGENT_WEB_SEARCH_SUMMARY_CHAR_LIMIT)
  assert.equal(compact.sources.length, AGENT_WEB_SEARCH_SOURCE_LIMIT)
  assert.ok(compact.sources.every(source => (source.snippet?.length || 0) <= AGENT_WEB_SEARCH_SNIPPET_CHAR_LIMIT))
  assert.equal(research.sources.length, 20)
  assert.equal(research.summary.length, AGENT_WEB_SEARCH_SUMMARY_CHAR_LIMIT + 100)
})

test('uses one external web search per turn and reuses completed research on retry', async () => {
  let calls = 0
  let requestInput: Record<string, unknown> | undefined
  const env = {
    AI: {
      async run(_model: string, input: Record<string, unknown>) {
        calls += 1
        requestInput = input
        return {
          output_text: 'Verified current product information.',
          output: [{ action: { sources: [{ title: 'Product', url: 'https://example.com/product' }] } }]
        }
      }
    },
    WORKBOOKS: { async put() {} }
  }
  const searchInput = {
    objective: 'Research current AI spreadsheet tools',
    queries: ['current AI spreadsheet tools'],
    recency: 'month' as const,
    preferredDomains: []
  }
  const runtime = createExcelAgentTools(env as never, {
    userId: 'test-user',
    sessionId: '49e6551b-5d94-4b58-b52b-d93e48b5ae89',
    billingEnabled: false
  })

  const first = await runtime.tools.searchWeb!.execute!(searchInput, {} as never)
  const second = await runtime.tools.searchWeb!.execute!(searchInput, {} as never)

  assert.equal(calls, 1)
  assert.equal(requestInput?.max_output_tokens, 5000)
  assert.deepEqual(second, first)
  assert.equal(runtime.state.researchSource, 'external')

  const retryRuntime = createExcelAgentTools({} as never, {
    userId: 'test-user',
    sessionId: '49e6551b-5d94-4b58-b52b-d93e48b5ae89',
    billingEnabled: false,
    initialResearch: runtime.state.research,
    initialResearchSnapshotR2Key: 'agent-demo/research/session/research.json'
  })
  const reused = await retryRuntime.tools.searchWeb!.execute!(searchInput, {} as never)
  assert.deepEqual(reused, first)
  assert.equal(retryRuntime.state.researchSource, 'initial')
  assert.equal(retryRuntime.getToolDetail('searchWeb')?.reused, true)
})

test('gives web search its own deadline before the overall agent timeout', async () => {
  let options: Record<string, unknown> | undefined
  const env = {
    AI: {
      async run(_model: string, _input: Record<string, unknown>, nextOptions?: Record<string, unknown>) {
        options = nextOptions
        return { output_text: 'ok' }
      }
    },
    WEB_SEARCH_TIMEOUT_MS: '15000'
  }

  await runResponsesModel(env as never, { input: 'test' })
  assert.ok(options?.signal instanceof AbortSignal)
  assert.equal((options?.signal as AbortSignal).aborted, false)
})

test('uses non-empty model knowledge records when live web search is unavailable', async () => {
  const originalConsoleError = console.error
  console.error = () => {}
  try {
    const runtime = createExcelAgentTools({} as never, {
      userId: 'test-user',
      sessionId: '49e6551b-5d94-4b58-b52b-d93e48b5ae89'
    })
    const searchResult = await runtime.tools.searchWeb!.execute!({
      objective: 'Research current AI spreadsheet tools',
      queries: ['current AI spreadsheet tools'],
      recency: 'month',
      preferredDomains: []
    }, {} as never)

    assert.equal((searchResult as { status?: string }).status, 'unavailable')
    assert.deepEqual(runtime.getToolDetail('searchWeb'), {
      skipped: true,
      fallback: 'continue_without_web',
      objective: 'Research current AI spreadsheet tools'
    })

    await assert.rejects(() => runtime.tools.createSpreadsheet!.execute!({
      plan: spreadsheetPlan,
      records: []
    }, {} as never), /dataMode model_knowledge/)

    const modelKnowledgePlan = {
      ...spreadsheetPlan,
      dataMode: 'model_knowledge' as const,
      includeSources: false
    }
    const records = [
      { name: 'Formula Bot', website: 'https://www.formulabot.com/' },
      { name: 'GPTExcel', website: 'https://gptexcel.uk/' }
    ]
    const spreadsheetResult = await runtime.tools.createSpreadsheet!.execute!({
      plan: modelKnowledgePlan,
      records: records.map(record => ({
        cells: Object.entries(record).map(([key, value]) => ({ key, value }))
      }))
    }, {} as never)
    assert.deepEqual(spreadsheetResult, {
      accepted: true,
      title: spreadsheetPlan.title,
      dataMode: 'model_knowledge',
      columnCount: 2,
      maxRows: 20,
      degraded: true
    })
    assert.equal(runtime.state.spreadsheetPlan?.dataMode, 'model_knowledge')
    assert.equal(runtime.state.spreadsheetPlan?.includeSources, false)
    assert.deepEqual(runtime.state.spreadsheetRecords, records)

    const generated = await buildSpreadsheetWorkbook(modelKnowledgePlan, runtime.state.spreadsheetRecords, [])
    assert.equal(generated.summary.rowCount, 2)
    assert.equal(generated.preview.length, 3)
  } finally {
    console.error = originalConsoleError
  }
})

test('builds a safe spreadsheet with language-neutral auxiliary sheets', async () => {
  assert.equal(sanitizeSpreadsheetCell('=HYPERLINK("bad")'), '\'=HYPERLINK("bad")')
  const generated = await buildSpreadsheetWorkbook(spreadsheetPlan, [
    { name: '=unsafe', website: 'https://example.com/a' },
    { name: 'Safe product', website: 'https://example.org/b' }
  ], [
    { id: 'S1', title: 'Product A', url: 'https://example.com/a' }
  ])

  const workbook = XLSX.read(generated.bytes, { type: 'array', cellNF: true, cellStyles: true })
  assert.deepEqual(workbook.SheetNames, ['Research', '__sources', '__metadata'])
  assert.equal(workbook.Workbook?.Sheets?.find(sheet => sheet.name === '__sources')?.Hidden, 1)
  assert.equal(workbook.Workbook?.Sheets?.find(sheet => sheet.name === '__metadata')?.Hidden, 1)
  const rows = XLSX.utils.sheet_to_json<unknown[]>(workbook.Sheets.Research!, { header: 1, raw: true })
  assert.deepEqual(rows[0], ['Name', 'Website'])
  assert.equal(rows[1]?.[0], "'=unsafe")
  assert.equal(workbook.Sheets.Research?.['!autofilter']?.ref, 'A1:B3')
  assert.equal(workbook.Sheets.Research?.B2?.l?.Target, 'https://example.com/a')
  assert.equal(generated.summary.rowCount, 2)
  assert.deepEqual(generated.summary.sheetNames, ['Research'])
  assert.equal(generated.preview.length, 3)
})

test('formats semantic columns and excludes hidden system sheets from workbook context', async () => {
  const generated = await buildSpreadsheetWorkbook({
    title: 'Quarterly sales',
    sheetName: 'Quarterly Sales',
    purpose: 'Review quarterly sales and growth',
    dataMode: 'conversation_data',
    columns: [
      { key: 'quarter', label: 'Quarter', description: 'Quarter name', type: 'text', required: true },
      { key: 'sales', label: 'Sales', description: 'Quarter sales', type: 'currency', required: true },
      { key: 'growth', label: 'Growth', description: 'Quarter growth rate', type: 'percent', required: true }
    ],
    maxRows: 10,
    includeSources: false
  }, [
    { quarter: 'Q1', sales: 120000, growth: 0.125 },
    { quarter: 'Q2', sales: 135000, growth: 0.052 }
  ])

  const workbook = XLSX.read(generated.bytes, { type: 'array', cellNF: true, cellStyles: true })
  assert.equal(workbook.Sheets['Quarterly Sales']?.B2?.z, '#,##0.00')
  assert.equal(workbook.Sheets['Quarterly Sales']?.C2?.z, '0.0%')
  assert.equal(workbook.Sheets['Quarterly Sales']?.A1?.s?.fgColor?.rgb, '132A3A')
  assert.equal(workbook.Sheets['Quarterly Sales']?.A3?.s?.fgColor?.rgb, 'F1F8F4')

  const parsed = await parseWorkbookBytes(generated.bytes)
  assert.deepEqual(parsed.summary.sheetNames, ['Quarterly Sales'])
  assert.equal(parsed.summary.rowCount, 2)
})

test('creates a generic execution plan without a scenario-specific template', () => {
  const plan = buildSpreadsheetExecutionPlan(spreadsheetPlanSchema.parse(spreadsheetPlan))
  assert.equal(plan.action, 'create_spreadsheet')
  assert.equal(plan.spreadsheet.columns[0]?.key, 'name')
  assert.deepEqual(plan.steps.map(step => step.operation), ['search', 'design', 'extract', 'validate', 'generate'])
})

test('encodes Unicode workbook names in an ASCII-safe download header', () => {
  const header = attachmentContentDisposition('季度销售分析示例表.xlsx')
  assert.equal(/^[\x20-\x7E]+$/.test(header), true)
  assert.match(header, /filename="excelgen-result\.xlsx"/)
  assert.match(header, /filename\*=UTF-8''%E5%AD%A3%E5%BA%A6/)
})
