import assert from 'node:assert/strict'
import test from 'node:test'
import {
  prepareWorkbookDataSheets,
  summarizeWorkbookSheets,
  type WorkbookDataSheet
} from '../src/workbook/analysis.ts'

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

test('profiles workbook columns and numeric statistics deterministically', () => {
  const summary = summarizeWorkbookSheets(salesWorkbook)

  assert.equal(summary.primarySheet, '销售数据')
  assert.equal(summary.rowCount, 6)
  assert.equal(summary.columnCount, 5)
  assert.equal(summary.sheets[0]?.sourceRange, 'A1:E7')
  assert.equal(summary.numericColumns.find(column => column.name === '销售额')?.sum, 79200)
  assert.equal(summary.blankCellCount, 0)
  assert.equal(summary.duplicateRowCount, 0)
})

test('keeps workbook parsing semantic and independent of conversation language', () => {
  const prepared = prepareWorkbookDataSheets(salesWorkbook)[0]!

  assert.deepEqual(prepared.headers, ['月份', '区域', '销售额', '成本', '订单数'])
  assert.deepEqual(prepared.rows[0], ['1月', '华北', 12800, 7800, 84])
  assert.equal(prepared.profile.columns.find(column => column.name === '销售额')?.type, 'number')
})

test('excludes an aggregate footer while previewing every row in a 68-record workbook', () => {
  const summary = summarizeWorkbookSheets(attendanceWorkbook)

  assert.equal(summary.rowCount, 68)
  assert.equal(summary.sheets[0]?.sourceRange, 'A2:D70')
  assert.equal(summary.numericColumns.find(column => column.name === '处罚')?.sum, 500)
  assert.equal(summary.preview.length, 69)
  assert.equal(summary.preview.at(-1)?.[0], 68)
})
