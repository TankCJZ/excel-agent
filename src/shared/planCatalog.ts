export type PlanKey = 'free' | 'pro' | 'max'

export type PlanEntitlements = {
  key: PlanKey
  monthlyCredits: number
  maxUploadBytes: number
  maxConcurrentTasks: number
  artifactRetentionDays: number
  maxWorkbookSheets: number
  maxWorkbookCells: number
  freeResearchRuns: number | null
  freeChartRuns: number | null
}

const MEBIBYTE = 1024 * 1024

export const PLAN_ENTITLEMENTS: Record<PlanKey, PlanEntitlements> = {
  free: {
    key: 'free',
    monthlyCredits: 0,
    maxUploadBytes: 5 * MEBIBYTE,
    maxConcurrentTasks: 1,
    artifactRetentionDays: 7,
    maxWorkbookSheets: 10,
    maxWorkbookCells: 50_000,
    freeResearchRuns: 1,
    freeChartRuns: 1
  },
  pro: {
    key: 'pro',
    monthlyCredits: 1_200,
    maxUploadBytes: 10 * MEBIBYTE,
    maxConcurrentTasks: 2,
    artifactRetentionDays: 90,
    maxWorkbookSheets: 30,
    maxWorkbookCells: 200_000,
    freeResearchRuns: null,
    freeChartRuns: null
  },
  max: {
    key: 'max',
    monthlyCredits: 4_000,
    maxUploadBytes: 20 * MEBIBYTE,
    maxConcurrentTasks: 5,
    artifactRetentionDays: 365,
    maxWorkbookSheets: 30,
    maxWorkbookCells: 200_000,
    freeResearchRuns: null,
    freeChartRuns: null
  }
}

export const CREDIT_POLICY_VERSION = 1
export const FREE_SIGNUP_CREDITS = 30

export const CREDIT_COSTS = {
  response: 2,
  inspectWorkbook: 2,
  editWorkbook: 8,
  analyzeWorkbook: 6,
  createSpreadsheet: 8,
  createChart: 4,
  exportAnalysisWorkbook: 2,
  searchWeb: 12,
  workbookArtifact: 2
} as const

export type BillableAgentTool = keyof Omit<typeof CREDIT_COSTS, 'response' | 'workbookArtifact'>

export type CreditUsageAction = { action: string, credits: number, idempotencyKey: string }

// Preserve every performed action (including waived actions for free-plan limits),
// but make the charged amounts add up to the actual, capped settlement.
export function allocateCreditUsage(actions: CreditUsageAction[], actualCredits: number) {
  let remaining = actualCredits
  const seen = new Set<string>()
  return actions.filter(action => {
    if (seen.has(action.idempotencyKey)) return false
    seen.add(action.idempotencyKey)
    return true
  }).map(action => {
    const nominalCredits = Math.max(0, Math.floor(action.credits))
    const credits = Math.min(remaining, nominalCredits)
    remaining -= credits
    return { ...action, credits, nominalCredits }
  })
}

const TOOL_COSTS: Partial<Record<string, number>> = {
  inspectWorkbook: CREDIT_COSTS.inspectWorkbook,
  inspectWorkbookForEditing: CREDIT_COSTS.inspectWorkbook,
  editWorkbook: CREDIT_COSTS.editWorkbook,
  analyzeWorkbook: CREDIT_COSTS.analyzeWorkbook,
  createSpreadsheet: CREDIT_COSTS.createSpreadsheet,
  createChart: CREDIT_COSTS.createChart,
  exportAnalysisWorkbook: CREDIT_COSTS.exportAnalysisWorkbook,
  searchWeb: CREDIT_COSTS.searchWeb
}

export function normalizePlanKey(value: unknown): PlanKey {
  return value === 'pro' || value === 'max' ? value : 'free'
}

export function getPlanEntitlements(value: unknown) {
  return PLAN_ENTITLEMENTS[normalizePlanKey(value)]
}

export function calculateActualTurnCredits(toolNames: Iterable<string>, includeWorkbookArtifact = false) {
  let credits = CREDIT_COSTS.response
  for (const toolName of new Set(toolNames)) credits += TOOL_COSTS[toolName] || 0
  if (includeWorkbookArtifact) credits += CREDIT_COSTS.workbookArtifact
  return credits
}

export function buildCreditUsageActions(
  turnId: string,
  toolNames: Iterable<string>,
  includeWorkbookArtifact = false
) {
  const actions: Array<{ action: string, credits: number, idempotencyKey: string }> = [{
    action: 'response',
    credits: CREDIT_COSTS.response,
    idempotencyKey: `credit-usage:${turnId}:response`
  }]
  for (const toolName of new Set(toolNames)) {
    const credits = TOOL_COSTS[toolName]
    if (!credits) continue
    actions.push({
      action: toolName,
      credits,
      idempotencyKey: `credit-usage:${turnId}:${toolName}`
    })
  }
  if (includeWorkbookArtifact) {
    actions.push({
      action: 'workbookArtifact',
      credits: CREDIT_COSTS.workbookArtifact,
      idempotencyKey: `credit-usage:${turnId}:workbookArtifact`
    })
  }
  return actions
}

export function estimateTurnCredits(
  message: string,
  hasWorkbook: boolean,
  requestedOutcome?: 'spreadsheet' | 'chart' | 'analysis',
  options: { reuseWebSearch?: boolean } = {}
) {
  const normalized = message.toLocaleLowerCase()
  const hasResearchIntent = matches(normalized, [
    'research', 'web search', 'market', 'industry', 'competitor', 'current data',
    '调研', '搜索', '行业', '市场', '竞品', '公司名单', '最新数据'
  ])
  const hasChartIntent = requestedOutcome === 'chart' || matches(normalized, [
    'chart', 'graph', 'dashboard', 'visualize', '图表', '仪表盘', '可视化', '趋势图'
  ])
  const hasAnalysisIntent = requestedOutcome === 'analysis' || matches(normalized, [
    'analyze', 'analysis', 'summary', 'clean', 'trend', 'outlier', 'compare',
    '分析', '清洗', '总结', '趋势', '异常', '比较', '统计'
  ])
  const hasSpreadsheetIntent = requestedOutcome === 'spreadsheet' || (!hasWorkbook && matches(normalized, [
    'spreadsheet', 'workbook', 'excel', 'xlsx', 'tracker', 'table', 'report',
    '表格', '工作簿', '数据表', '跟踪表', '报表', '清单', '创建', '生成'
  ]))
  const hasEditIntent = hasWorkbook && matches(normalized, ['edit', 'change', 'update', 'replace', 'append', 'formula', 'fill', 'add column', 'add sheet', '修改', '更新', '替换', '增加', '新增', '添加', '公式', '填充', '写入'])

  let credits = CREDIT_COSTS.response
  if (hasWorkbook) credits += CREDIT_COSTS.inspectWorkbook
  if (hasAnalysisIntent && hasWorkbook) credits += CREDIT_COSTS.analyzeWorkbook
  if (hasSpreadsheetIntent) credits += CREDIT_COSTS.createSpreadsheet + CREDIT_COSTS.workbookArtifact
  if (hasEditIntent) credits += CREDIT_COSTS.editWorkbook + CREDIT_COSTS.workbookArtifact
  if (hasChartIntent) credits += CREDIT_COSTS.createChart
  if (hasResearchIntent && !options.reuseWebSearch) credits += CREDIT_COSTS.searchWeb

  // Keep the public estimate predictable. Actual settlement never exceeds the held amount.
  return Math.max(CREDIT_COSTS.response, Math.min(30, credits))
}

function matches(value: string, needles: string[]) {
  return needles.some(needle => value.includes(needle))
}
