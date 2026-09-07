import assert from 'node:assert/strict'
import test from 'node:test'
import * as XLSX from 'xlsx'
import { unzipSync, strFromU8 } from 'fflate'
import { spreadsheetCreationInputSchema, type SpreadsheetPlan } from '../src/shared/contracts.ts'
import { buildSpreadsheetWorkbook } from '../src/spreadsheet/generation.ts'
import { parseWorkbookBytes } from '../src/workbook/service.ts'
import { createExcelAgentTools } from '../src/mastra/tools/excelTools.ts'

const plan: SpreadsheetPlan = {
  title: 'Order tracker', purpose: 'Calculate order totals', sheetName: 'Orders', dataMode: 'conversation_data', maxRows: 20, includeSources: false,
  columns: ['item', 'quantity', 'price', 'total', 'example'].map(key => ({ key, label: key, description: key, type: key === 'item' || key === 'example' ? 'text' : 'number', required: false }))
}

test('new workbooks preserve explicit formulas and keep formula examples as escaped text', async () => {
  const generated = await buildSpreadsheetWorkbook(plan, [
    { item: 'Alpha', quantity: 3, price: 20, total: { type: 'formula', formula: '=B2*C2' }, example: '=SUM(B2:C2)' },
    { item: 'Beta', quantity: 5, price: 12, total: { type: 'formula', formula: '=B3*C3' } }
  ])
  const book = XLSX.read(generated.bytes, { cellFormula: true, sheetStubs: true })
  assert.equal(book.Sheets.Orders!.D2.f, 'B2*C2')
  assert.equal(book.Sheets.Orders!.D3.f, 'B3*C3')
  assert.equal(book.Sheets.Orders!.E2.f, undefined)
  assert.equal(book.Sheets.Orders!.E2.v, "'=SUM(B2:C2)")
  const xml = strFromU8(unzipSync(generated.bytes)['xl/worksheets/sheet1.xml']!)
  assert.match(xml, /<c r="D2"[^>]*><f>B2\*C2<\/f><\/c>/)
  assert.equal(generated.summary.calculation?.status, 'pending')
  assert.equal(generated.preview[1]![3], '=B2*C2')
  assert.equal((await parseWorkbookBytes(generated.bytes)).summary.calculation?.formulaCells, 2)
})

test('typed formula validation rejects unsafe and cyclic cells before the tool accepts a plan', async () => {
  for (const formula of ['=WEBSERVICE("https://example.invalid")', '=D2+1', '=SUM([external.xlsx]Data!A1)']) {
    await assert.rejects(buildSpreadsheetWorkbook(plan, [{ total: { type: 'formula', formula } }]))
  }
  const runtime = createExcelAgentTools({} as never, { userId: 'test', sessionId: 'test', billingEnabled: false })
  await assert.rejects(runtime.tools.createSpreadsheet!.execute!({ plan, records: [{ cells: [{ key: 'total', value: { type: 'formula', formula: '=D2+1' } }] }] } as never, {} as never))
  assert.equal(runtime.state.spreadsheetPlan, null)
})

test('blank templates can contain formulas and null inputs, but cannot smuggle sample data', () => {
  const templatePlan = { ...plan, dataMode: 'blank_template' }
  const records = [{ cells: [{ key: 'quantity', value: null }, { key: 'price', value: null }, { key: 'total', value: { type: 'formula', formula: '=B2*C2' } }] }]
  assert.equal(spreadsheetCreationInputSchema.safeParse({ plan: templatePlan, records }).success, true)
  assert.equal(spreadsheetCreationInputSchema.safeParse({ plan: templatePlan, records: [{ cells: [{ key: 'quantity', value: 100 }] }] }).success, false)
  assert.equal(spreadsheetCreationInputSchema.safeParse({ plan, records: [{ cells: [{ key: 'total', value: { type: 'formula', formula: '=B2*C2', cachedValue: 999 } }] }] }).success, false)
})
