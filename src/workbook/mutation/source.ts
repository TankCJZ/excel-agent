import * as XLSX from 'xlsx'
import { MUTATION_LIMITS } from '#shared/agent/workbookMutation'
import { requireMutation } from '#agent/workbook/mutation/errors'
import { inspectWorkbookForEditing } from '#agent/workbook/mutation/engine'
import { readCsvAsText } from '#agent/workbook/readSafety'

export function prepareEditableBytes(source: Uint8Array, fileName: string) {
  requireMutation(source.length <= MUTATION_LIMITS.sourceBytes, 'FILE_LIMIT', 'Workbook exceeds the editing size limit.')
  const extension = fileName.split('.').at(-1)?.toLowerCase()
  requireMutation(['xlsx', 'xls', 'csv'].includes(extension || ''), 'UNSUPPORTED_WORKBOOK', 'Edit an XLSX workbook or convert an XLS/CSV file to XLSX.')
  if (extension === 'xlsx') return { bytes: source, convertedFrom: null }
  requireMutation(extension !== 'xls', 'UNSUPPORTED_WORKBOOK', 'For safe editing, open this XLS file in Excel and Save As XLSX, then upload that copy. Automatic conversion could lose formatting or embedded objects. The original is unchanged.')
  if (extension === 'csv') requireMutation(!source.includes(0) && !(source[0] === 0x50 && source[1] === 0x4b), 'UNSUPPORTED_WORKBOOK', 'Upload a UTF-8 CSV file or a correctly named XLSX workbook.')
  const book = readCsvAsText(source)
  requireMutation(book.SheetNames.length <= MUTATION_LIMITS.sheets && book.SheetNames.reduce((count, name) => count + Object.keys(book.Sheets[name]!).filter(key => !key.startsWith('!')).length, 0) <= MUTATION_LIMITS.physicalCells, 'FILE_LIMIT', 'Converted workbook exceeds safe editing limits.')
  const bytes = new Uint8Array(XLSX.write(book, { type: 'array', bookType: 'xlsx', cellStyles: true }))
  inspectWorkbookForEditing(bytes)
  return { bytes, convertedFrom: extension as 'xls' | 'csv' }
}
