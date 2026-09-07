import type { WorkBook, WorkSheet } from 'xlsx'
import * as XLSX from 'xlsx'

// Check the rectangle before sheet_to_json allocates empty cells. A workbook
// with two populated cells at A1 and XFD1048576 must not expand into billions.
export function assertReadableWorkbook(book: WorkBook) {
  let slots = 0
  let populated = 0
  if (book.SheetNames.length > 64) throw new Error('Workbook has too many worksheets to preview safely')
  for (const name of book.SheetNames) {
    const sheet = book.Sheets[name] as WorkSheet | undefined
    if (!sheet) throw new Error('Worksheet data is missing')
    populated += Object.keys(sheet).filter(key => !key.startsWith('!')).length
    if (populated > 150_000) throw new Error('Workbook contains too many populated cells to preview safely')
    if (!sheet['!ref']) continue
    const range = XLSX.utils.decode_range(sheet['!ref'])
    const count = (range.e.r - range.s.r + 1) * (range.e.c - range.s.c + 1)
    if (!Number.isSafeInteger(count) || count < 1 || range.e.r > 1048575 || range.e.c > 16383) throw new Error('Invalid worksheet dimensions')
    slots += count
    if (slots > 500_000 || populated > 150_000) throw new Error('Worksheet range is too large or too sparse to preview safely. Remove unused rows and columns, then upload it again.')
  }
}

export function readCsvAsText(bytes: Uint8Array) {
  // CSV has no cell types. Do not infer dates, numbers, leading-zero IDs or
  // executable formulas from uploaded text; edits may explicitly add types.
  const text = new TextDecoder('utf-8', { fatal: true, ignoreBOM: false }).decode(bytes)
  const book = XLSX.read(text, { type: 'string', raw: true, cellFormula: false })
  assertReadableWorkbook(book)
  return book
}
