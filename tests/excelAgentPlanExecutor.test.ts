import assert from 'node:assert/strict'
import test from 'node:test'
import type { WorkbookDataSheet } from '../src/workbook/analysis.ts'
import { mergeWorkbookAnalyses } from '../src/workbook/analysisMerge.ts'
import { analysisPlanSchema } from '../src/contracts.ts'
import {
  AnalysisPlanError,
  executeWorkbookAnalysisPlan
} from '../src/workbook/planExecutor.ts'

const salesWorkbook: WorkbookDataSheet[] = [{
  name: '销售数据',
  rows: [
    ['月份', '区域', '销售额', '成本', '订单数'],
    ['1月', '华北', 12800, 7800, 84],
    ['1月', '华南', 10400, 6500, 72],
    ['2月', '华北', 14600, 8200, 91],
    ['2月', '华南', 11900, 7100, 79],
    ['3月', '华北', 16300, 8800, 104],
    ['3月', '华南', 13200, 7600, 88]
  ]
}]

const penaltyNames = ['Zed', 'Ethan', 'DY', 'Dorain', 'Nolan', 'Snowman', 'Spike']
const penaltyValues = [100, 50, 50, 100, 50, 50, 100]
const attendanceWorkbook: WorkbookDataSheet[] = [{
  name: '12月考勤汇总表',
  rows: [
    ['2025年12月考勤汇总表', null, null, null],
    ['序号', '姓名', '实薪（天）', '处罚'],
    ...Array.from({ length: 68 }, (_, index) => [
      index + 1,
      penaltyNames[index] || `员工${index + 1}`,
      23,
      penaltyValues[index] ?? null
    ]),
    [null, null, 1564, 500]
  ]
}]

test('executes an LLM-authored grouping plan with deterministic values', () => {
  const plan = analysisPlanSchema.parse({
    goal: '哪个区域的销售额最高？',
    sheet: '销售数据',
    groupBy: ['区域'],
    metrics: [{ operation: 'sum', field: '销售额', alias: '销售额合计' }],
    sort: [{ field: '销售额合计', direction: 'desc' }],
    limit: 10
  })

  const result = executeWorkbookAnalysisPlan(salesWorkbook, plan)

  assert.equal(result.kind, 'analysis')
  assert.deepEqual(result.table?.columns, ['区域', '销售额合计'])
  assert.deepEqual(result.table?.rows, [
    ['华北', 43700],
    ['华南', 35500]
  ])
  assert.equal(result.metrics.find(metric => metric.role === 'requested')?.value, 79200)
  assert.equal(result.metrics.find(metric => metric.role === 'group_count')?.value, 2)
  assert.deepEqual(result.stats, { sourceRowCount: 6, matchingRowCount: 6, outputRowCount: 2, groupCount: 2 })
  assert.deepEqual(result.evidence, [{ sheet: '销售数据', range: 'A1:E7', basis: 'validated_plan' }])
  assert.equal(result.plan?.metrics[0]?.operation, 'sum')
})

test('applies validated filters before grouping and aggregation', () => {
  const plan = analysisPlanSchema.parse({
    goal: '只看华北，按月份统计销售额',
    sheet: '销售数据',
    where: {
      logic: 'and',
      conditions: [{ field: '区域', operator: 'eq', value: '华北' }]
    },
    groupBy: ['月份'],
    metrics: [{ operation: 'sum', field: '销售额', alias: '销售额合计' }],
    sort: [{ field: '月份', direction: 'asc' }],
    limit: 20
  })

  const result = executeWorkbookAnalysisPlan(salesWorkbook, plan)

  assert.deepEqual(result.table?.rows, [
    ['1月', 12800],
    ['2月', 14600],
    ['3月', 16300]
  ])
  assert.equal(result.metrics.find(metric => metric.label === '销售额合计')?.value, 43700)
  assert.equal(result.stats.matchingRowCount, 3)
  assert.equal(result.stats.outputRowCount, 3)
})

test('creates chart data only from deterministic grouped results', () => {
  const plan = analysisPlanSchema.parse({
    goal: '生成一张按月份比较销售额的折线图',
    sheet: '销售数据',
    groupBy: ['月份'],
    metrics: [{ operation: 'sum', field: '销售额', alias: '销售额合计' }],
    sort: [{ field: '月份', direction: 'asc' }],
    limit: 20,
    chart: { type: 'line', title: '月度销售额' }
  })

  const result = executeWorkbookAnalysisPlan(salesWorkbook, plan)

  assert.equal(result.kind, 'chart')
  assert.equal(result.chart?.type, 'line')
  assert.deepEqual(result.chart?.categories, ['1月', '2月', '3月'])
  assert.deepEqual(result.chart?.series[0]?.data, [23200, 26500, 29500])
  assert.equal(result.chart?.source.range, 'A1:E7')
})

test('supports reusable data-quality metrics without a hardcoded business scenario', () => {
  const workbook: WorkbookDataSheet[] = [{
    name: '客户',
    rows: [
      ['账户', '城市'],
      ['A', '上海'],
      ['A', '上海'],
      ['B', null],
      ['C', '北京']
    ]
  }]
  const plan = analysisPlanSchema.parse({
    goal: '检查账户数量、空白城市和重复行',
    sheet: '客户',
    metrics: [
      { operation: 'distinct_count', field: '账户', alias: '账户数' },
      { operation: 'blank_count', field: '城市', alias: '空白城市' },
      { operation: 'duplicate_count', alias: '重复行' }
    ]
  })

  const result = executeWorkbookAnalysisPlan(workbook, plan)

  assert.deepEqual(result.table?.rows, [[3, 1, 1]])
  assert.deepEqual(result.metrics.slice(0, 3).map(metric => metric.value), [3, 1, 1])
})

test('rejects hallucinated fields with correctable schema evidence', () => {
  const plan = analysisPlanSchema.parse({
    goal: '统计不存在的利润字段',
    sheet: '销售数据',
    metrics: [{ operation: 'sum', field: '净利润', alias: '净利润合计' }]
  })

  assert.throws(
    () => executeWorkbookAnalysisPlan(salesWorkbook, plan),
    (error: unknown) => {
      assert.ok(error instanceof AnalysisPlanError)
      assert.equal(error.code, 'unknown_field')
      assert.match(error.message, /可用|Available fields/i)
      assert.match(error.message, /销售额/)
      return true
    }
  )
})

test('schema blocks unbounded output and invalid chart plans', () => {
  const tooLarge = analysisPlanSchema.safeParse({
    goal: '列出所有数据',
    sheet: '销售数据',
    select: ['月份'],
    limit: 5000
  })
  const chartWithoutAggregation = analysisPlanSchema.safeParse({
    goal: '生成图表',
    sheet: '销售数据',
    select: ['月份'],
    chart: { type: 'bar' }
  })

  assert.equal(tooLarge.success, false)
  assert.equal(chartWithoutAggregation.success, false)
})

test('answers a people-and-total question in one deterministic plan', () => {
  const plan = analysisPlanSchema.parse({
    goal: '这个表中有哪些人收到了处罚？处罚总金额是多少？',
    sheet: '12月考勤汇总表',
    where: {
      conditions: [{ field: '处罚', operator: 'gt', value: 0 }]
    },
    groupBy: ['姓名'],
    metrics: [
      { operation: 'sum', field: '处罚', alias: '处罚金额' },
      { operation: 'count', field: '姓名', alias: '人数' }
    ],
    sort: [{ field: '处罚金额', direction: 'desc' }],
    limit: 20
  })

  const result = executeWorkbookAnalysisPlan(attendanceWorkbook, plan)

  assert.equal(result.table?.rows.length, 7)
  assert.deepEqual(result.table?.columns, ['姓名', '处罚金额'])
  assert.equal(result.metrics.find(metric => metric.label === '处罚金额')?.value, 500)
  assert.equal(result.metrics.find(metric => metric.label === '人数')?.value, 7)
  assert.equal(result.metrics.filter(metric => metric.label === '人数').length, 1)
  assert.equal(result.metrics.length, 2)
  assert.deepEqual(result.table?.rows.map(row => row[0]).sort(), [...penaltyNames].sort())
  assert.deepEqual(result.stats, { sourceRowCount: 68, matchingRowCount: 7, outputRowCount: 7, groupCount: 7 })
})

test('merges multiple model tool calls without losing names or the correct total', () => {
  const listResult = executeWorkbookAnalysisPlan(attendanceWorkbook, analysisPlanSchema.parse({
    goal: '列出所有受到处罚的人员姓名及其处罚金额',
    sheet: '12月考勤汇总表',
    select: ['姓名', '处罚'],
    where: { conditions: [{ field: '处罚', operator: 'gt', value: 0 }] },
    sort: [{ field: '处罚', direction: 'desc' }]
  }))
  const totalResult = executeWorkbookAnalysisPlan(attendanceWorkbook, analysisPlanSchema.parse({
    goal: '计算处罚金额的总和',
    sheet: '12月考勤汇总表',
    metrics: [{ operation: 'sum', field: '处罚', alias: '处罚总金额' }]
  }))

  const merged = mergeWorkbookAnalyses(listResult, totalResult)

  assert.equal(merged.metrics.find(metric => metric.label === '处罚总金额')?.value, 500)
  assert.equal(merged.table?.rows.length, 7)
  assert.deepEqual(merged.table?.rows.map(row => row[0]).sort(), [...penaltyNames].sort())
  assert.deepEqual(merged.stats, { sourceRowCount: 68, matchingRowCount: 68, outputRowCount: 1, groupCount: 0 })
})
