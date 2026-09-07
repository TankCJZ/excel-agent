import assert from 'node:assert/strict'
import test from 'node:test'
import * as XLSX from 'xlsx'
import { assertReadableWorkbook } from '../src/workbook/readSafety.ts'
import { prepareEditableBytes } from '../src/workbook/mutation/source.ts'

test('huge sparse dimensions are rejected before sheet_to_json allocates a rectangle', () => {
  const book = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(book, { A1: { t: 's', v: 'Title' }, XFD1048576: { t: 'n', v: 1 }, '!ref': 'A1:XFD1048576' }, 'Sheet1')
  assert.throws(() => assertReadableWorkbook(book), /too sparse/)
})

test('CSV editing preserves IDs, date-like strings and formula-like strings as text', () => {
  const csv = new TextEncoder().encode('ID,Date,Formula,Name\n00123,1-2,=SUM(A1:A3),测试\n')
  const converted = prepareEditableBytes(csv, 'input.csv')
  const sheet = XLSX.read(converted.bytes).Sheets.Sheet1!
  for (const [address, value] of Object.entries({ A2: '00123', B2: '1-2', C2: '=SUM(A1:A3)', D2: '测试' })) {
    assert.equal(sheet[address]!.v, value)
    assert.equal(sheet[address]!.t, 's')
    assert.equal(sheet[address]!.f, undefined)
  }
  assert.equal(converted.convertedFrom, 'csv')
})

test('XLS conversion never silently drops unsupported features', () => {
  assert.throws(() => prepareEditableBytes(new Uint8Array([0xd0, 0xcf]), 'legacy.xls'), /Save As XLSX/)
})
