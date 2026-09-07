import assert from 'node:assert/strict'
import test from 'node:test'
import * as XLSX from 'xlsx'
import { strToU8, strFromU8, unzipSync, zipSync } from 'fflate'
import { parseWorkbookBytes } from '../src/workbook/service.ts'
import { hasUncalculatedXlsxFormulas } from '../src/workbook/formulaCalculation.ts'
import { applyWorkbookMutation } from '../src/workbook/mutation/engine.ts'
import { createExcelAgentTools } from '../src/mastra/tools/excelTools.ts'
import { analysisPlanSchema } from '../src/shared/contracts.ts'
import { encodeWorkbookContext, decodeWorkbookContext, WORKBOOK_CONTEXT_VERSION } from '../src/workbook/context.ts'

function fixture(cached = true) {
  const book = XLSX.utils.book_new()
  const sheet = XLSX.utils.aoa_to_sheet([['Item', 'Price', 'Quantity', 'Total'], ['Alpha', 12, 3, null]])
  sheet.D2 = { t: 'n', f: 'B2*C2', ...(cached ? { v: 36 } : {}) }
  sheet['!ref'] = 'A1:D2'
  XLSX.utils.book_append_sheet(book, sheet, 'Sales')
  return new Uint8Array(XLSX.write(book, { type: 'array', bookType: 'xlsx' }))
}

test('uncalculated formulas are never dropped, synthesized as zero or exposed as old results', async () => {
  const input = fixture(false)
  assert.equal(hasUncalculatedXlsxFormulas(input), true)
  const context = await parseWorkbookBytes(input)
  assert.deepEqual(context.summary.calculation, { status: 'pending', formulaCells: 1 })
  assert.deepEqual(context.sheets[0]!.rows[1], ['Alpha', 12, 3, '=B2*C2'])
  assert.equal(context.summary.numericColumns.some(column => column.name === 'Total'), false)
  const encoded = await encodeWorkbookContext({ version: WORKBOOK_CONTEXT_VERSION, sourceEtag: 'etag', ...context })
  const decoded = await decodeWorkbookContext(encoded)
  assert.deepEqual(decoded.summary.calculation, context.summary.calculation)
  assert.deepEqual(decoded.sheets, context.sheets)
})

test('a normal Excel cached result stays readable, but mutation invalidates the result across subsequent edits', async () => {
  const input = fixture()
  assert.equal(hasUncalculatedXlsxFormulas(input), false)
  const original = await parseWorkbookBytes(input)
  assert.equal(original.summary.calculation, undefined)
  assert.equal(original.sheets[0]!.rows[1]![3], 36)
  const v2 = applyWorkbookMutation(input, { title: 'Change price', operations: [{ kind: 'set_values', sheet: 'Sales', start: 'B2', values: [[20]], overwrite: true }] })
  const v3 = applyWorkbookMutation(v2.bytes, { title: 'Change quantity', operations: [{ kind: 'set_values', sheet: 'Sales', start: 'C2', values: [[4]], overwrite: true }] })
  const context = await parseWorkbookBytes(v3.bytes)
  assert.deepEqual(context.sheets[0]!.rows[1], ['Alpha', 20, 4, '=B2*C2'])
  assert.equal(context.summary.calculation?.status, 'pending')
  // Simulate the saved Excel result. Real Excel opening is separately required.
  const parts = unzipSync(v3.bytes)
  parts['xl/worksheets/sheet1.xml'] = strToU8(strFromU8(parts['xl/worksheets/sheet1.xml']!).replace('<f>B2*C2</f>', '<f>B2*C2</f><v>80</v>'))
  const recalculated = await parseWorkbookBytes(zipSync(parts))
  assert.equal(recalculated.summary.calculation, undefined)
  assert.equal(recalculated.sheets[0]!.rows[1]![3], 80)
})

test('analysis, chart and export tools reject pending calculation while inspection stays available', async () => {
  const context = await parseWorkbookBytes(fixture(false))
  const runtime = createExcelAgentTools({} as never, { workbookId: 'workbook-1', fileKey: 'test.xlsx', sessionId: 'calculation-test', summary: context.summary, initialContext: context, billingEnabled: false })
  const summary = await runtime.tools.inspectWorkbook!.execute!({}, {} as never)
  assert.equal((summary as typeof context.summary).calculation?.status, 'pending')
  const plan = analysisPlanSchema.parse({ goal: 'Calculate total', sheet: 'Sales', metrics: [{ operation: 'sum', field: 'Total', alias: 'total' }] })
  for (const tool of ['analyzeWorkbook', 'createChart', 'exportAnalysisWorkbook']) {
    await assert.rejects(async () => runtime.tools[tool]!.execute!({ plan }, {} as never), { code: 'WORKBOOK_CALCULATION_PENDING' })
    assert.equal(runtime.state.analysis, null)
    assert.equal(runtime.state.workflowRequested, false)
  }
})

test('explicit mutation state overrides even plausible stale cached values', async () => {
  const context = await parseWorkbookBytes(fixture(), { pendingCalculation: true })
  assert.equal(context.summary.calculation?.status, 'pending')
  assert.equal(context.sheets[0]!.rows[1]![3], '=B2*C2')
})

test('Excel shared-formula saves re-upload as calculated and remain editable with every follower preserved', async () => {
  const book = XLSX.utils.book_new()
  const sheet = XLSX.utils.aoa_to_sheet([['Item', 'Price', 'Quantity', 'Total'], ['Alpha', 12, 3, 36], ['Beta', 2, 5, 10]])
  sheet.D2 = { t: 'n', f: 'B2*C2', v: 36 }
  sheet.D3 = { t: 'n', f: 'B3*C3', v: 10 }
  XLSX.utils.book_append_sheet(book, sheet, 'Sales')
  const parts = unzipSync(new Uint8Array(XLSX.write(book, { type: 'array', bookType: 'xlsx' })))
  parts['xl/worksheets/sheet1.xml'] = strToU8(strFromU8(parts['xl/worksheets/sheet1.xml']!)
    .replace('<f>B2*C2</f>', '<f t="shared" ref="D2:D3" si="0">B2*C2</f>')
    .replace('<f>B3*C3</f>', '<f t="shared" si="0"/>'))
  const saved = zipSync(parts)
  const context = await parseWorkbookBytes(saved)
  assert.equal(context.summary.calculation, undefined)
  assert.equal(context.sheets[0]!.rows[2]![3], 10)
  const edited = applyWorkbookMutation(saved, { title: 'Keep shared formulas', operations: [{ kind: 'set_values', sheet: 'Sales', start: 'B2', values: [[20]], overwrite: true }] })
  const checked = XLSX.read(edited.bytes.slice(), { cellFormula: true, sheetStubs: true }).Sheets.Sales!
  assert.equal(checked.D2.f, 'B2*C2')
  assert.equal(checked.D3.f, 'B3*C3')
  assert.equal(checked.B2.v, 20)
  assert.equal((await parseWorkbookBytes(edited.bytes)).summary.calculation?.status, 'pending')
})
