import assert from 'node:assert/strict'
import test from 'node:test'
import * as XLSX from 'xlsx'
import { strToU8, strFromU8, unzipSync, zipSync } from 'fflate'
import { workbookMutationInputSchema, MUTATION_LIMITS } from '../src/shared/workbookMutation.ts'
import { applyWorkbookMutation, inspectWorkbookForEditing } from '../src/workbook/mutation/engine.ts'
import { parseSafeFormula, fillFormula, assertAcyclicFormulas } from '../src/workbook/mutation/formulas.ts'
import { parseRange } from '../src/workbook/mutation/coordinates.ts'
import { XlsxPackage } from '../src/workbook/mutation/package.ts'
import { readBoundedZip } from '../src/workbook/mutation/boundedZip.ts'

function sample() {
  const book = XLSX.utils.book_new()
  const sheet = XLSX.utils.aoa_to_sheet([
    ['Quarterly report'], [], ['Item', 'Sales', 'Cost'], ['Alpha', 120, 80], [], ['Beta', 200, 150]
  ])
  sheet['!merges'] = [XLSX.utils.decode_range('A1:C1')]
  sheet['!rows'] = [{}, {}, {}, {}, { hidden: true }]
  sheet['!cols'] = [{ wch: 24 }, { wch: 14 }, { wch: 14 }]
  XLSX.utils.book_append_sheet(book, sheet, 'Sales Data')
  XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet([['Tax'], [0.2]]), '税率')
  return new Uint8Array(XLSX.write(book, { type: 'array', bookType: 'xlsx' }))
}
function mutate(bytes: Uint8Array, operations: unknown[]) {
  return applyWorkbookMutation(bytes, workbookMutationInputSchema.parse({ title: 'Updated workbook', operations }))
}
function contents(bytes: Uint8Array, path: string) { return strFromU8(unzipSync(bytes)[path]!) }

test('dense XML is rejected before DOM allocation even when ZIP sizes and physical cell counts are small', () => {
  const parts = unzipSync(sample())
  parts['xl/worksheets/sheet1.xml'] = strToU8(`<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${'<row r="1"/>'.repeat(MUTATION_LIMITS.xmlComplexity)}</sheetData></worksheet>`)
  assert.throws(() => inspectWorkbookForEditing(zipSync(parts)), /too complex to edit safely/)
})

test('edit plan budgets count UTF-8 bytes and all targeted cells before execution', () => {
  assert.equal(workbookMutationInputSchema.safeParse({ title: 'Large edit', operations: [
    { kind: 'set_values', sheet: 'Data', start: 'A1', overwrite: false, values: [Array.from({ length: 10 }, () => '文'.repeat(10000))] }
  ] }).success, false)
  assert.equal(workbookMutationInputSchema.safeParse({ title: 'Too many cells', operations: [
    { kind: 'set_formula', sheet: 'Data', range: 'A1:A5001', formula: '=1+1', overwrite: false },
    { kind: 'set_number_format', sheet: 'Data', range: 'B1:B5001', format: 'integer' }
  ] }).success, false)
  assert.equal(workbookMutationInputSchema.safeParse({ title: 'Invalid range', operations: [
    { kind: 'set_formula', sheet: 'Data', range: '???', formula: '=1+1', overwrite: false }
  ] }).success, false)
})

test('streamed ZIP reads verify actual sizes and reject forged expansion declarations', () => {
  const input = zipSync({ 'large.txt': new Uint8Array(2_000_000).fill(65) })
  const data = new DataView(input.buffer, input.byteOffset, input.byteLength)
  for (let offset = 0; offset < input.length - 30; offset++) {
    if (data.getUint32(offset, true) === 0x04034b50) data.setUint32(offset + 22, 10, true)
    if (data.getUint32(offset, true) === 0x02014b50) data.setUint32(offset + 24, 10, true)
  }
  assert.throws(() => readBoundedZip(input), /larger than its declared size/)
  const normal = zipSync({ 'one.xml': strToU8('one'), 'two.xml': strToU8('two') })
  const result = readBoundedZip(normal, name => name === 'two.xml')
  assert.deepEqual(Object.keys(result), ['two.xml'])
  assert.equal(strFromU8(result['two.xml']!), 'two')
})

test('physical coordinates survive blank/header/hidden rows; source and unrelated parts are preserved', () => {
  const source = sample()
  const copy = source.slice()
  const parts = unzipSync(source)
  parts['xl/media/image1.png'] = new Uint8Array([1, 2, 3, 4])
  parts['xl/charts/chart1.xml'] = strToU8('<chart>untouched fixture</chart>')
  const input = zipSync(parts)
  const result = mutate(input, [{ kind: 'set_values', sheet: 'Sales Data', start: 'B6', values: [[250]], overwrite: true }])
  const book = XLSX.read(result.bytes)
  assert.equal(book.Sheets['Sales Data']!.B6.v, 250)
  assert.equal(book.Sheets['Sales Data']!.B4.v, 120)
  assert.equal(book.Sheets['Sales Data']!.B5, undefined)
  assert.deepEqual(source, copy)
  const output = unzipSync(result.bytes)
  for (const path of Object.keys(parts).filter(path => path !== 'xl/worksheets/sheet1.xml')) assert.deepEqual(output[path], parts[path], path)
  assert.match(contents(result.bytes, 'xl/worksheets/sheet1.xml'), /hidden="1"/)
  assert.equal(result.changes.changedCells, 1)
  assert.equal(result.changes.sheets[0]!.rows[0]!.row, 5)
})

test('writes actual formulas, retains relative/absolute references, and requests recalculation without false caches', () => {
  const result = mutate(sample(), [
    { kind: 'append_column', sheet: 'Sales Data', title: 'Profit', headerRow: 3, values: [] },
    { kind: 'set_formula', sheet: 'Sales Data', range: 'D4:D6', formula: '=B4-C4', overwrite: false },
    { kind: 'set_formula', sheet: 'Sales Data', range: 'E4:E6', formula: '=D4*(1-\'税率\'!$A$2)', overwrite: false }
  ])
  const book = XLSX.read(result.bytes, { cellFormula: true, sheetStubs: true })
  assert.equal(book.Sheets['Sales Data']!.D6.f, 'B6-C6')
  assert.equal(book.Sheets['Sales Data']!.E6.f, "D6*(1-'税率'!$A$2)")
  assert.match(contents(result.bytes, 'xl/worksheets/sheet1.xml'), /<c r="D6"><f>B6-C6<\/f><\/c>/)
  assert.match(contents(result.bytes, 'xl/workbook.xml'), /fullCalcOnLoad="1"/)
  assert.match(contents(result.bytes, 'xl/workbook.xml'), /forceFullCalc="1"/)
  assert.equal(result.changes.pendingCalculation, true)
  assert.equal(result.changes.formulaCells, 6)
  const cell = result.changes.sheets[0]!.rows.flatMap(row => row.cells).find(cell => cell.address === 'D4')!
  assert.equal(cell.value, null)
  assert.equal(cell.formula, '=B4-C4')
})

test('a second edit retains the first and invalidates existing dependent caches', () => {
  const book = XLSX.read(sample())
  book.Sheets['Sales Data']!.D4 = { t: 'n', f: 'B4-C4', v: 40 }
  book.Sheets['Sales Data']!['!ref'] = 'A1:D6'
  const source = new Uint8Array(XLSX.write(book, { type: 'array', bookType: 'xlsx' }))
  const v2 = mutate(source, [{ kind: 'set_values', sheet: 'Sales Data', start: 'B4', values: [[300]], overwrite: true }])
  const v3 = mutate(v2.bytes, [{ kind: 'set_values', sheet: 'Sales Data', start: 'C4', values: [[100]], overwrite: true }])
  const result = XLSX.read(v3.bytes, { cellFormula: true, sheetStubs: true }).Sheets['Sales Data']!
  assert.equal(result.B4.v, 300)
  assert.equal(result.C4.v, 100)
  assert.equal(result.D4.f, 'B4-C4')
  assert.match(contents(v3.bytes, 'xl/worksheets/sheet1.xml'), /<c r="D4"><f>B4-C4<\/f><\/c>/)
})

test('adds worksheets and treats formula-like strings as literal text', () => {
  const result = mutate(sample(), [{ kind: 'add_sheet', sheet: 'Notes & examples', values: [['=SUM(A1:A3)', '+12', '-12', '@SUM', true, 12, null, '<xml> & "quoted"']] }])
  const book = XLSX.read(result.bytes)
  assert.equal(book.Sheets['Notes & examples']!.A1.t, 's')
  assert.equal(book.Sheets['Notes & examples']!.A1.f, undefined)
  assert.equal(book.Sheets['Notes & examples']!.A1.v, '=SUM(A1:A3)')
  assert.equal(book.Sheets['Notes & examples']!.H1.v, '<xml> & "quoted"')
  assert.deepEqual(result.changes.addedSheets, ['Notes & examples'])
})

test('date writes use the workbook date system and preserve number formats', () => {
  for (const date1904 of [false, true]) {
    const book = XLSX.read(sample())
    book.Workbook = { ...book.Workbook, WBProps: { date1904 } }
    const input = new Uint8Array(XLSX.write(book, { type: 'array', bookType: 'xlsx' }))
    const result = mutate(input, [
      { kind: 'set_values', sheet: 'Sales Data', start: 'F4', values: [[{ type: 'date', value: '2026-09-05' }, 0.125]], overwrite: false },
      { kind: 'set_number_format', sheet: 'Sales Data', range: 'G4', format: 'percent' }
    ])
    const read = XLSX.read(result.bytes, { cellNF: true }).Sheets['Sales Data']!
    assert.equal(read.F4.v, date1904 ? 44808 : 46270)
    assert.equal(read.F4.z, 'm/d/yy')
    assert.equal(read.G4.z, '0.00%')
  }
})

test('editing context returns physical coordinates instead of normalized analysis indices', () => {
  const context = inspectWorkbookForEditing(sample())
  assert.equal(context.sheets[0]!.preview.find(cell => cell.address === 'A6')!.value, 'Beta')
  assert.deepEqual(context.sheets[0]!.mergedRanges, ['A1:C1'])
  assert.equal(context.sheets[0]!.lastColumn, 'C')
})

test('no-op returns byte-identical source and no changed-cell charge signal', () => {
  const source = sample()
  const result = mutate(source, [{ kind: 'set_values', sheet: 'Sales Data', start: 'B4', values: [[120]], overwrite: true }])
  assert.equal(result.changes.changedCells, 0)
  assert.deepEqual(result.bytes, source)
})

test('rejects missing sheets, overwrites without consent, merged children, duplicate writes and invalid ranges', () => {
  const cases = [
    { kind: 'set_values', sheet: 'Missing', start: 'A1', values: [[1]], overwrite: true },
    { kind: 'set_values', sheet: 'Sales Data', start: 'B4', values: [[1]], overwrite: false },
    { kind: 'set_values', sheet: 'Sales Data', start: 'B1', values: [[1]], overwrite: true },
    { kind: 'append_column', sheet: 'Sales Data', title: 'Sales', headerRow: 3, values: [] },
    { kind: 'add_sheet', sheet: 'sales data', values: [] },
    { kind: 'set_values', sheet: 'Sales Data', start: 'XFE1', values: [[1]], overwrite: true },
    { kind: 'set_values', sheet: 'Sales Data', start: 'B1048577', values: [[1]], overwrite: true },
    { kind: 'set_values', sheet: 'Sales Data', start: 'F4', values: [[{ type: 'date', value: '2026-02-30' }]], overwrite: true },
    { kind: 'set_formula', sheet: 'Sales Data', range: 'F1:F20000', formula: '=1', overwrite: true }
  ]
  for (const operation of cases) assert.throws(() => mutate(sample(), [operation]))
  assert.throws(() => mutate(sample(), [cases[1], { ...cases[1], overwrite: true }]))
  const op = { kind: 'set_values', sheet: 'Sales Data', start: 'F1', values: [[1]], overwrite: true }
  assert.throws(() => mutate(sample(), [op, op]), /Multiple operations/)
  assert.throws(() => parseRange('B5:A1'), /top-left/)
})

test('rejects unsafe/partial formulas and detects cyclic dependencies including existing formulas', () => {
  for (const formula of ['=WEBSERVICE("https://example.com")', '=A1 garbage', '=SUM(A1', '=SUM()', '=SUM(A1,,A2)', '=INDIRECT("A1")', '=cmd|\' /C calc\'!A0', '=[book.xlsx]Sheet!A1', '=UnknownName+1', '=A1#', '=Sheet1:Sheet2!A1', '=SUMIFS(A1,A2)', '=COUNTIFS(A1)', '=A1+']) {
    assert.throws(() => parseSafeFormula(formula, 'Sheet1', ['Sheet1', 'Sheet2']), formula)
  }
  assert.throws(() => mutate(sample(), [{ kind: 'set_formula', sheet: 'Sales Data', range: 'F4', formula: '=F4+1', overwrite: false }]), /circular/)
  assert.throws(() => assertAcyclicFormulas([
    { sheet: 'A', address: 'A1', formula: "'B'!A1" }, { sheet: 'B', address: 'A1', formula: "'A'!A1" }
  ], ['A', 'B']), /circular/)
})

test('formula fill changes reference tokens only, handles mixed references and escaped quotes', () => {
  assert.equal(fillFormula('=IF(A1="A1",$B1+C$2+$D$4,\'Bob\'\'s Data\'!E1)', 'Sheet', ['Sheet', "Bob's Data"], 2, 1), 'IF(B3="A1",$B3+D$2+$D$4,\'Bob\'\'s Data\'!F3)')
  assert.equal(parseSafeFormula('=XLOOKUP(A1,B1:B10,C1:C10)', 'Sheet', ['Sheet']).formula, '_xlfn.XLOOKUP(A1,B1:B10,C1:C10)')
  assert.doesNotThrow(() => parseSafeFormula('=SUM(A1:B2)+IF(TRUE,1,0)', 'Sheet', ['Sheet']))
  assert.throws(() => fillFormula('=A1', 'Sheet', ['Sheet'], -1, 0))
})

test('rejects protected sheets, unsafe package contents and ZIP expansion before editing', () => {
  const source = sample()
  const protectedParts = unzipSync(source)
  protectedParts['xl/worksheets/sheet1.xml'] = strToU8(strFromU8(protectedParts['xl/worksheets/sheet1.xml']!).replace('</worksheet>', '<sheetProtection sheet="1"/></worksheet>'))
  assert.throws(() => mutate(zipSync(protectedParts), [{ kind: 'set_values', sheet: 'Sales Data', start: 'B4', values: [[1]], overwrite: true }]), /protected/)
  const macroParts = unzipSync(source)
  macroParts['xl/vbaProject.bin'] = new Uint8Array([1])
  assert.throws(() => new XlsxPackage(zipSync(macroParts)), /macros/)
  const entityParts = unzipSync(source)
  entityParts['xl/workbook.xml'] = strToU8('<!DOCTYPE workbook [<!ENTITY x "attack">]><workbook/>')
  assert.throws(() => new XlsxPackage(zipSync(entityParts)), /entities/)
  const bomb = zipSync({ ...unzipSync(source), 'huge.xml': new Uint8Array(MUTATION_LIMITS.partBytes + 1) })
  assert.throws(() => new XlsxPackage(bomb), /safe processing limits/)
})
