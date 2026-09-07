import { MockLanguageModelV3 } from 'ai/test'

// Test-only scripted model. Production Worker modules must never import this fixture.
export const MOCK_MODEL_ID = 'mock/excel-planner'

export function createMockExcelModel() {
  const generate = async (options: Parameters<MockLanguageModelV3['doGenerate']>[0]) => {
      const { prompt } = options
      const userPrompt = extractUserText(prompt)
      const availableTools = new Set((options.tools || []).map(tool => 'name' in tool ? tool.name : ''))
      const hasExplicitTools = Boolean(options.tools?.length)
      const toolAvailable = (name: string) => !hasExplicitTools || availableTools.has(name)
      const hasWorkbookTools = toolAvailable('inspectWorkbook')
      const wantsChart = /(chart|graph|visuali[sz]e|plot|bar|line|pie|scatter|图表|柱状图|折线图|饼图|散点图|趋势图|走势图|分布图|可视化)/i.test(userPrompt)
      const wantsAnalysis = /(analy[sz]e|statistics?|summary|insight|trend|compare|top|highest|lowest|average|total|amount|how\s+many|who|quality|分析|统计|总结|趋势|对比|比较|最高|最低|平均|合计|总额|金额|多少|人数|哪些人|谁|排名|异常|数据质量)/i.test(userPrompt)
      const wantsDescription = isWorkbookDescriptionPrompt(userPrompt)
      const wantsExport = /(?:导出|下载|保存|另存).{0,12}(?:分析结果|工作簿|excel|xlsx|文件|报告)|(?:生成|创建).{0,8}(?:新的?|分析|结果|报告)?(?:工作簿|excel\s*文件|xlsx\s*文件|报告文件)|(?:export|download|save|create|generate).{0,24}(?:analysis|result|report)?.{0,8}(?:workbook|spreadsheet|xlsx|file|report)/i.test(userPrompt)
      const wantsSpreadsheet = /(?:创建|生成|制作|导出|整理成).{0,48}(?:excel|xlsx|电子表格|工作簿|表格|数据表|模板|清单)|(?:create|generate|build|export|make).{0,80}(?:excel|xlsx|spreadsheet|workbook|data\s*table|template|dataset|directory)/i.test(userPrompt)
      const wantsWebSearch = /(?:最新|当前|今天|近期|联网|网上|网页|市场调研|行业调研|检索|查找资料|资料来源|新闻)|(?:latest|current|today|recent|web|online|market\s+research|industry\s+research|search|sources?|news|pricing|funding)/i.test(userPrompt)
      const wantsUnsupportedEdit = /(?:修改|更新|编辑|写回|删除|去重|清洗|添加|写入).{0,12}(?:单元格|列|行|公式|工作簿|表格|excel)|(?:edit|modify|update|write\s*back|delete|deduplicate|clean|add|insert).{0,24}(?:cell|column|row|formula|workbook|spreadsheet|excel)/i.test(userPrompt)
      const wantsWorkbookExport = wantsExport && hasWorkbookTools
        && /(?:导出|下载|保存|另存).{0,32}(?:工作簿|excel|xlsx|文件|报告)|(?:export|download|save).{0,48}(?:workbook|spreadsheet|xlsx|file|report)/i.test(userPrompt)
      const wantsWorkbookAction = hasWorkbookTools
        && (wantsWorkbookExport || (!wantsSpreadsheet && (wantsChart || wantsAnalysis)))
      const workbookSummary = extractToolResult(prompt, 'inspectWorkbook')
      const webResearch = extractToolResult(prompt, 'searchWeb')

      if (wantsWebSearch && !webResearch && toolAvailable('searchWeb')) {
        return mockModelResult({
          content: [{
            type: 'tool-call',
            toolCallId: crypto.randomUUID(),
            toolName: 'searchWeb',
            input: JSON.stringify({
              objective: userPrompt,
              queries: [userPrompt],
              recency: 'any',
              preferredDomains: []
            })
          }],
          finishReason: 'tool-calls'
        })
      }

      if (!wantsUnsupportedEdit && (wantsWorkbookAction || (wantsDescription && hasWorkbookTools)) && !workbookSummary) {
        return mockModelResult({
          content: [{
            type: 'tool-call',
            toolCallId: crypto.randomUUID(),
            toolName: 'inspectWorkbook',
            input: JSON.stringify({})
          }],
          finishReason: 'tool-calls'
        })
      }

      const plan = buildMockAnalysisPlan(userPrompt, workbookSummary, { chart: wantsChart })

      if (!wantsUnsupportedEdit && wantsWorkbookExport && !hasToolResult(prompt, 'exportAnalysisWorkbook')) {
        return mockModelResult({
          content: [{
            type: 'tool-call',
            toolCallId: crypto.randomUUID(),
            toolName: 'exportAnalysisWorkbook',
            input: JSON.stringify({ plan })
          }],
          finishReason: 'tool-calls'
        })
      }

      if (!wantsUnsupportedEdit && !wantsWorkbookExport && !wantsSpreadsheet && wantsChart && hasWorkbookTools && !hasToolResult(prompt, 'createChart')) {
        return mockModelResult({
          content: [{
            type: 'tool-call',
            toolCallId: crypto.randomUUID(),
            toolName: 'createChart',
            input: JSON.stringify({ plan })
          }],
          finishReason: 'tool-calls'
        })
      }

      if (!wantsUnsupportedEdit && !wantsWorkbookExport && !wantsSpreadsheet && !wantsChart && wantsAnalysis && hasWorkbookTools && !hasToolResult(prompt, 'analyzeWorkbook')) {
        return mockModelResult({
          content: [{
            type: 'tool-call',
            toolCallId: crypto.randomUUID(),
            toolName: 'analyzeWorkbook',
            input: JSON.stringify({ plan })
          }],
          finishReason: 'tool-calls'
        })
      }

      if (wantsSpreadsheet && !wantsWorkbookExport && !hasToolResult(prompt, 'createSpreadsheet') && toolAvailable('createSpreadsheet')) {
        const columns = [
          { key: 'item', label: 'item', description: 'entity_or_observation', type: 'text', required: true },
          { key: 'details', label: 'details', description: 'structured_details', type: 'text', required: true },
          { key: 'value', label: 'value', description: 'primary_numeric_value', type: 'number', required: false },
          { key: 'source_url', label: 'source_url', description: 'supporting_source_url', type: 'url', required: false }
        ]
        const searchUnavailable = getStringProperty(webResearch, 'status') === 'unavailable'
        const dataMode = searchUnavailable
          ? 'model_knowledge'
          : webResearch
            ? 'web_research'
            : /(?:空白|模板|blank|template)/i.test(userPrompt)
              ? 'blank_template'
              : 'conversation_data'
        return mockModelResult({
          content: [{
            type: 'tool-call',
            toolCallId: crypto.randomUUID(),
            toolName: 'createSpreadsheet',
            input: JSON.stringify({
              plan: {
                title: 'mock-spreadsheet',
                sheetName: 'data',
                purpose: userPrompt,
                dataMode,
                columns,
                maxRows: 30,
                includeSources: dataMode !== 'model_knowledge'
              },
              records: dataMode === 'conversation_data' || dataMode === 'model_knowledge'
                ? [{ cells: [
                    { key: 'item', value: userPrompt },
                    { key: 'details', value: dataMode === 'model_knowledge' ? 'Not live-verified; generated from the model\'s stable knowledge.' : userPrompt },
                    { key: 'value', value: null },
                    { key: 'source_url', value: null }
                  ] }]
                : []
            })
          }],
          finishReason: 'tool-calls'
        })
      }

      const exportResult = extractToolResult(prompt, 'exportAnalysisWorkbook')
      const analysis = extractToolResult(prompt, 'analyzeWorkbook') || getObjectProperty(exportResult, 'analysis')
      const searchSummary = getStringProperty(webResearch, 'summary')
      const description = formatMockWorkbookDescription(extractToolResult(prompt, 'inspectWorkbook'))
      const text = searchSummary
        || formatMockResult(analysis)
        || (wantsDescription ? description : '')
        || userPrompt
      return mockModelResult({
        content: [{
          type: 'text',
          text
        }],
        finishReason: 'stop'
      })
  }

  return new MockLanguageModelV3({
    provider: 'excelgen-mock',
    modelId: MOCK_MODEL_ID,
    doGenerate: generate,
    doStream: async (options) => createMockModelStream(await generate(options))
  })
}

function isWorkbookDescriptionPrompt(question: string) {
  return /(?:这|此)(?:是|份|个)?.{0,8}(?:什么|哪类|哪种).{0,8}(?:数据)?表|(?:这|此)份表.{0,8}(?:是做什么|记录什么|有什么|包含什么|主要内容)|(?:介绍|说明|描述|概括).{0,8}(?:工作簿|表格|数据表|这份表)|(?:表格|数据表|工作簿).{0,8}(?:是什么|记录什么|包含什么|主要内容)|what\s+(?:kind|type)\s+of\s+(?:data|table)|what\s+is\s+(?:this|the)\s+(?:data|sheet|table|workbook)|what\s+does\s+(?:this|the)\s+(?:sheet|table|workbook)\s+(?:contain|show|represent)|describe\s+(?:this|the)\s+(?:data|sheet|table|workbook)/i.test(question)
}

function buildMockAnalysisPlan(question: string, summary: unknown, options: { chart: boolean }) {
  const workbook = summary && typeof summary === 'object' ? summary : {}
  const sheet = getStringProperty(workbook, 'primarySheet') || 'Sheet1'
  const headers = getStringArrayProperty(workbook, 'headers')
  const numericColumns = getObjectArrayProperty(workbook, 'numericColumns')
    .map(column => getStringProperty(column, 'name'))
    .filter(Boolean)
  const requestedHeader = headers.find(header => normalizeIdentifier(question).includes(normalizeIdentifier(header)))
  const metricField = requestedHeader && numericColumns.includes(requestedHeader)
    ? requestedHeader
    : findHeader(headers, /(销售额|营收|金额|收入|利润|成本|revenue|sales|amount|income|profit|cost)/i, numericColumns)
      || numericColumns[0]
  const requestedDimension = headers.find(header => !numericColumns.includes(header)
    && normalizeIdentifier(question).includes(normalizeIdentifier(header)))
  const personRequested = /(哪些人|谁|人员|员工|人数|姓名|who|which\s+(?:people|persons|employees|users)|how\s+many\s+(?:people|persons|employees|users)|names?)/i.test(question)
  const personField = findHeader(
    headers,
    /^(姓名|人员|员工|人名|用户|账户|账号|name|person|employee|user|account)$/i,
    headers.filter(header => !numericColumns.includes(header))
  )
  const dimensionField = requestedDimension || (personRequested ? personField : undefined) || findHeader(
    headers,
    options.chart
      ? /(月份|月|日期|时间|区域|地区|分类|产品|month|date|time|region|category|product)/i
      : /(区域|地区|分类|产品|月份|月|日期|时间|region|category|product|month|date|time)/i,
    headers.filter(header => !numericColumns.includes(header))
  )
  const qualityRequested = /(quality|blank|missing|duplicate|数据质量|空白|缺失|重复)/i.test(question)
  const wantsAverage = /(average|mean|avg|平均|均值)/i.test(question)
  const wantsCount = /(count|how many|number of|数量|多少|几条|几笔)/i.test(question)
  const wantsTotal = /(total|sum|合计|总计|总额|总金额|金额.{0,8}(?:多少|几)|(?:多少|几).{0,8}金额)/i.test(question)
  const operation = wantsAverage ? 'average' as const : wantsTotal ? 'sum' as const : wantsCount ? 'count' as const : 'sum' as const

  if (qualityRequested) {
    const metrics: Array<{
      operation: 'blank_count' | 'duplicate_count'
      field?: string
      alias: string
      format: 'number'
    }> = headers.slice(0, 6).map(field => ({
      operation: 'blank_count' as const,
      field,
      alias: `${field}_blank_count`,
      format: 'number' as const
    }))
    metrics.push({
      operation: 'duplicate_count',
      alias: 'duplicate_count',
      format: 'number'
    })
    return {
      goal: question,
      sheet,
      select: [],
      groupBy: [],
      metrics,
      sort: [],
      limit: 20
    }
  }

  if (metricField && dimensionField) {
    const alias = `${metricField}_${operation}`
    const timeLike = /(月份|月|日期|时间|month|date|time)/i.test(dimensionField)
    return {
      goal: question,
      sheet,
      select: [],
      ...(personRequested && metricField
        ? {
            where: {
              logic: 'and' as const,
              conditions: [{ field: metricField, operator: 'gt' as const, value: 0 }]
            }
          }
        : {}),
      groupBy: [dimensionField],
      metrics: [{ operation, field: metricField, alias, format: 'number' as const }],
      sort: [{ field: options.chart && timeLike ? dimensionField : alias, direction: options.chart && timeLike ? 'asc' as const : 'desc' as const }],
      limit: 20,
      ...(options.chart
        ? {
            chart: {
              type: /(line|折线|趋势)/i.test(question) || timeLike ? 'line' as const : /(pie|饼图)/i.test(question) ? 'pie' as const : 'bar' as const,
              title: question
            }
          }
        : {})
    }
  }

  return {
    goal: question,
    sheet,
    select: [],
    groupBy: [],
    metrics: [
      { operation: 'count' as const, alias: 'row_count', format: 'number' as const },
      { operation: 'duplicate_count' as const, alias: 'duplicate_count', format: 'number' as const }
    ],
    sort: [],
    limit: 20
  }
}

function findHeader(headers: string[], pattern: RegExp, candidates: string[]) {
  return candidates.find(header => pattern.test(header)) || headers.find(header => pattern.test(header))
}

function getStringArrayProperty(value: object, property: string) {
  return property in value && Array.isArray(value[property as keyof typeof value])
    ? (value[property as keyof typeof value] as unknown[]).map(String)
    : []
}

function getObjectArrayProperty(value: object, property: string) {
  return property in value && Array.isArray(value[property as keyof typeof value])
    ? (value[property as keyof typeof value] as unknown[]).filter((item): item is object => Boolean(item) && typeof item === 'object')
    : []
}

function normalizeIdentifier(value: string) {
  return value.normalize('NFKC').toLocaleLowerCase().replace(/[\s_-]+/g, '')
}

function extractToolResult(prompt: unknown, toolName: string): unknown {
  if (!Array.isArray(prompt)) return null
  for (let messageIndex = prompt.length - 1; messageIndex >= 0; messageIndex -= 1) {
    const message = prompt[messageIndex]
    if (!message || typeof message !== 'object' || !('content' in message) || !Array.isArray(message.content)) continue
    for (const part of message.content) {
      if (!part || typeof part !== 'object' || !('type' in part) || part.type !== 'tool-result') continue
      if (!('toolName' in part) || part.toolName !== toolName) continue
      if ('output' in part) return unwrapToolValue(part.output)
      if ('result' in part) return unwrapToolValue(part.result)
    }
  }
  return null
}

function unwrapToolValue(value: unknown, depth = 0): unknown {
  if (depth > 3) return value
  if (typeof value === 'string') {
    try {
      return unwrapToolValue(JSON.parse(value), depth + 1)
    } catch {
      return value
    }
  }
  if (!value || typeof value !== 'object') return value
  if ('value' in value) return unwrapToolValue(value.value, depth + 1)
  if ('output' in value) return unwrapToolValue(value.output, depth + 1)
  return value
}

function getStringProperty(value: unknown, property: string) {
  return value && typeof value === 'object' && property in value && typeof value[property as keyof typeof value] === 'string'
    ? String(value[property as keyof typeof value])
    : ''
}

function getObjectProperty(value: unknown, property: string): object | null {
  if (!value || typeof value !== 'object' || !(property in value)) return null
  const propertyValue = value[property as keyof typeof value]
  return propertyValue && typeof propertyValue === 'object' ? propertyValue : null
}

function formatMockWorkbookDescription(value: unknown) {
  if (!value || typeof value !== 'object') return ''
  const headers = 'headers' in value && Array.isArray(value.headers) ? value.headers.map(String) : []
  const rowCount = 'rowCount' in value && typeof value.rowCount === 'number' ? value.rowCount : 0
  return JSON.stringify({ rowCount, headers: headers.slice(0, 12) })
}

function formatMockResult(value: unknown) {
  if (!value || typeof value !== 'object') return ''
  return JSON.stringify(value)
}

function mockModelResult(input: {
  content: Array<
    | { type: 'text', text: string }
    | { type: 'tool-call', toolCallId: string, toolName: string, input: string }
  >
  finishReason: 'stop' | 'tool-calls'
}) {
  return {
    content: input.content,
    finishReason: { unified: input.finishReason, raw: undefined },
    usage: {
      inputTokens: { total: 48, noCache: 48, cacheRead: undefined, cacheWrite: undefined },
      outputTokens: { total: 24, text: 24, reasoning: undefined }
    },
    warnings: []
  }
}

type MockStreamPart =
  | { type: 'stream-start', warnings: [] }
  | { type: 'text-start', id: string }
  | { type: 'text-delta', id: string, delta: string }
  | { type: 'text-end', id: string }
  | { type: 'tool-call', toolCallId: string, toolName: string, input: string }
  | {
      type: 'finish'
      usage: ReturnType<typeof mockModelResult>['usage']
      finishReason: ReturnType<typeof mockModelResult>['finishReason']
    }

function createMockModelStream(result: ReturnType<typeof mockModelResult>) {
  const chunks: MockStreamPart[] = [{ type: 'stream-start', warnings: [] }]
  for (const part of result.content) {
    if (part.type === 'tool-call') {
      chunks.push({
        type: 'tool-call',
        toolCallId: part.toolCallId,
        toolName: part.toolName,
        input: part.input
      })
      continue
    }

    const textId = crypto.randomUUID()
    chunks.push({ type: 'text-start', id: textId })
    const characters = Array.from(part.text)
    for (let index = 0; index < characters.length; index += 12) {
      chunks.push({
        type: 'text-delta',
        id: textId,
        delta: characters.slice(index, index + 12).join('')
      })
    }
    chunks.push({ type: 'text-end', id: textId })
  }
  chunks.push({
    type: 'finish',
    usage: result.usage,
    finishReason: result.finishReason
  })

  let index = 0
  return {
    stream: new ReadableStream<MockStreamPart>({
      pull(controller) {
        const chunk = chunks[index]
        if (chunk) {
          index += 1
          controller.enqueue(chunk)
        } else {
          controller.close()
        }
      }
    })
  }
}

function hasToolResult(prompt: unknown, toolName: string) {
  if (!Array.isArray(prompt)) {
    return false
  }

  let latestUserIndex = -1
  for (let index = prompt.length - 1; index >= 0; index -= 1) {
    const message: unknown = prompt[index]
    if (message && typeof message === 'object' && 'role' in message && message.role === 'user') {
      latestUserIndex = index
      break
    }
  }
  return prompt.slice(latestUserIndex + 1).some((message) => {
    if (!message || typeof message !== 'object' || !('role' in message) || message.role !== 'tool') {
      return false
    }
    const content = 'content' in message && Array.isArray(message.content) ? message.content : []
    return content.some((part: unknown) => (
      part
      && typeof part === 'object'
      && 'type' in part
      && part.type === 'tool-result'
      && 'toolName' in part
      && part.toolName === toolName
    ))
  })
}

function extractUserText(prompt: unknown) {
  if (!Array.isArray(prompt)) {
    return 'Analyze this workbook and add a summary sheet.'
  }

  const userMessages = prompt.filter((message) => (
    message && typeof message === 'object' && 'role' in message && message.role === 'user'
  ))
  const latest = userMessages.at(-1)
  if (!latest || typeof latest !== 'object' || !('content' in latest)) {
    return 'Analyze this workbook and add a summary sheet.'
  }
  if (typeof latest.content === 'string') {
    return latest.content
  }
  if (!Array.isArray(latest.content)) {
    return 'Analyze this workbook and add a summary sheet.'
  }

  return latest.content
    .filter((part: unknown) => part && typeof part === 'object' && 'type' in part && part.type === 'text' && 'text' in part)
    .map((part: unknown) => String((part as { text: unknown }).text))
    .join('\n') || 'Analyze this workbook and add a summary sheet.'
}
