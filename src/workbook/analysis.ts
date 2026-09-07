import type {
  WorkbookSheet,
  WorkbookSummary
} from '#agent/contracts'

export type CellValue = string | number | boolean | null

export type WorkbookDataSheet = {
  name: string
  rows: CellValue[][]
}

export type PreparedWorkbookSheet = {
  name: string
  headerRowIndex: number
  headers: string[]
  rows: CellValue[][]
  profile: WorkbookSheet
}

const PREVIEW_DATA_ROW_LIMIT = 100

export function summarizeWorkbookSheets(sheets: WorkbookDataSheet[]): WorkbookSummary {
  const prepared = prepareWorkbookDataSheets(sheets)
  const primary = prepared.find((sheet) => sheet.rows.length > 0) || prepared[0] || prepareSheet({ name: 'Sheet1', rows: [] })

  return {
    sheetNames: prepared.map((sheet) => sheet.name),
    primarySheet: primary.name,
    rowCount: primary.profile.rowCount,
    columnCount: primary.profile.columnCount,
    headers: primary.headers,
    numericColumns: primary.profile.columns.flatMap((column) => column.numeric
      ? [{ name: column.name, ...column.numeric }]
      : []),
    blankCellCount: primary.profile.blankCellCount,
    duplicateRowCount: primary.profile.duplicateRowCount,
    sheets: prepared.map((sheet) => sheet.profile),
    preview: primary.profile.preview
  }
}

export function prepareWorkbookDataSheets(sheets: WorkbookDataSheet[]): PreparedWorkbookSheet[] {
  return sheets.map(prepareSheet)
}

function prepareSheet(sheet: WorkbookDataSheet): PreparedWorkbookSheet {
  const headerRowIndex = findHeaderRow(sheet.rows)
  const rawHeaders = sheet.rows[headerRowIndex] || []
  const columnCount = Math.max(rawHeaders.length, ...sheet.rows.slice(headerRowIndex + 1).map((row) => row.length), 0)
  const headers = makeUniqueHeaders(Array.from({ length: columnCount }, (_, index) => normalizeHeader(rawHeaders[index], index)))
  const candidateRows = sheet.rows.slice(headerRowIndex + 1)
    .filter((row) => row.some((cell) => !isEmpty(cell)))
    .map((row) => Array.from({ length: columnCount }, (_, index) => row[index] ?? null))
  const rows = candidateRows.filter((row, index) => !isAggregateFooterRow(row, index, candidateRows, headers))
  const columns = headers.map((name, index) => profileColumn(name, index, rows.map((row) => row[index] ?? null)))
  const blankCellCount = rows.reduce((total, row) => total + row.filter(isEmpty).length, 0)
  const duplicateRowCount = countDuplicates(rows)
  const finalRow = Math.max(headerRowIndex + rows.length + 1, 1)
  const finalColumn = Math.max(columnCount, 1)
  const profile: WorkbookSheet = {
    name: sheet.name,
    rowCount: rows.length,
    columnCount,
    sourceRange: `A${headerRowIndex + 1}:${columnLetter(finalColumn - 1)}${finalRow}`,
    headers,
    columns,
    blankCellCount,
    duplicateRowCount,
    preview: [headers, ...rows.slice(0, PREVIEW_DATA_ROW_LIMIT)]
  }

  return { name: sheet.name, headerRowIndex, headers, rows, profile }
}

function isAggregateFooterRow(
  row: CellValue[],
  rowIndex: number,
  rows: CellValue[][],
  headers: string[]
) {
  // Summary rows are normally placed at the end. Keeping the check local to the
  // footer avoids dropping incomplete records in the middle of a dataset.
  if (rowIndex < Math.max(1, rows.length - 3)) return false

  const textCells = row
    .filter((value): value is string => typeof value === 'string' && value.trim().length > 0)
    .map((value) => normalizeText(value))
  if (textCells.some((value) => /^(?:合计|总计|小计|汇总|grandtotal|subtotal|total)$/.test(value))) {
    return true
  }

  const identifierIndexes = headers.flatMap((header, index) => (
    /^(?:序号|编号|行号|姓名|人员|员工|用户|账户|账号|id|no|number|name|person|employee|user|account)$/i.test(normalizeText(header))
      ? [index]
      : []
  ))
  if (!identifierIndexes.length || identifierIndexes.some((index) => !isEmpty(row[index]))) return false

  const priorRows = rows.slice(0, rowIndex)
  let matchingAggregateColumns = 0
  for (let columnIndex = 0; columnIndex < row.length; columnIndex += 1) {
    const footerValue = row[columnIndex]
    if (typeof footerValue !== 'number' || !Number.isFinite(footerValue) || Math.abs(footerValue) < 1e-9) continue
    const priorValues = priorRows.flatMap((priorRow) => {
      const value = priorRow[columnIndex]
      return typeof value === 'number' && Number.isFinite(value) ? [value] : []
    })
    if (!priorValues.length) continue
    const priorSum = priorValues.reduce((total, value) => total + value, 0)
    if (Math.abs(priorSum - footerValue) <= Math.max(1e-6, Math.abs(footerValue) * 1e-9)) {
      matchingAggregateColumns += 1
    }
  }

  return matchingAggregateColumns >= 2
}

function findHeaderRow(rows: CellValue[][]) {
  const searchLimit = Math.min(rows.length, 10)
  let bestIndex = 0
  let bestScore = -1
  for (let index = 0; index < searchLimit; index += 1) {
    const row = rows[index] || []
    const nonEmpty = row.filter((cell) => !isEmpty(cell)).length
    const stringValues = row.filter((cell) => typeof cell === 'string' && cell.trim()).length
    const score = nonEmpty * 2 + stringValues
    if (nonEmpty >= 2 && score > bestScore) {
      bestIndex = index
      bestScore = score
    }
  }
  return bestIndex
}

function profileColumn(name: string, index: number, values: CellValue[]): WorkbookSheet['columns'][number] {
  const present = values.filter((value) => !isEmpty(value))
  const kinds = present.map(inferCellType)
  const typeCounts = new Map<string, number>()
  for (const kind of kinds) typeCounts.set(kind, (typeCounts.get(kind) || 0) + 1)
  const dominant = [...typeCounts.entries()].sort((left, right) => right[1] - left[1])[0]
  const type = present.length === 0
    ? 'empty'
    : dominant && dominant[1] / present.length >= 0.8
      ? dominant[0] as WorkbookSheet['columns'][number]['type']
      : 'mixed'
  const numbers = present.filter((value): value is number => typeof value === 'number' && Number.isFinite(value))
  const sortedNumbers = [...numbers].sort((left, right) => left - right)
  const sum = numbers.reduce((total, value) => total + value, 0)

  return {
    name,
    index,
    type,
    nonEmptyCount: present.length,
    nullCount: values.length - present.length,
    uniqueCount: new Set(present.map((value) => String(value))).size,
    numeric: numbers.length
      ? {
          count: numbers.length,
          sum: roundNumber(sum),
          average: roundNumber(sum / numbers.length),
          median: roundNumber(median(sortedNumbers)),
          min: sortedNumbers[0] || 0,
          max: sortedNumbers.at(-1) || 0
        }
      : null
  }
}

function inferCellType(value: CellValue): WorkbookSheet['columns'][number]['type'] {
  if (typeof value === 'number') return 'number'
  if (typeof value === 'boolean') return 'boolean'
  if (typeof value === 'string' && isDateLike(value)) return 'date'
  return 'text'
}

function isDateLike(value: string) {
  const trimmed = value.trim()
  return /^\d{4}[-/]\d{1,2}[-/]\d{1,2}(?:[ T].*)?$/.test(trimmed)
    || /^\d{1,2}[-/]\d{1,2}[-/]\d{2,4}$/.test(trimmed)
}

function countDuplicates(rows: CellValue[][]) {
  const seen = new Set<string>()
  let duplicates = 0
  for (const row of rows) {
    const signature = JSON.stringify(row.map((cell) => typeof cell === 'string' ? cell.trim() : cell))
    if (seen.has(signature)) duplicates += 1
    else seen.add(signature)
  }
  return duplicates
}

function normalizeHeader(value: CellValue | undefined, index: number) {
  const normalized = typeof value === 'string' ? value.trim() : value
  return normalized === null || normalized === undefined || normalized === ''
    ? `Column ${index + 1}`
    : String(normalized)
}

function makeUniqueHeaders(headers: string[]) {
  const counts = new Map<string, number>()
  return headers.map((header) => {
    const count = (counts.get(header) || 0) + 1
    counts.set(header, count)
    return count === 1 ? header : `${header} (${count})`
  })
}

function median(values: number[]) {
  if (!values.length) return 0
  const middle = Math.floor(values.length / 2)
  return values.length % 2 === 0
    ? ((values[middle - 1] || 0) + (values[middle] || 0)) / 2
    : values[middle] || 0
}

function isEmpty(value: CellValue | undefined) {
  return value === null || value === undefined || (typeof value === 'string' && value.trim() === '')
}

function normalizeText(value: string) {
  return value.toLowerCase().replace(/[\s_\-./\\()[\]{}:：,，]+/g, '')
}

function roundNumber(value: number) {
  return Math.round((value + Number.EPSILON) * 100) / 100
}

function columnLetter(index: number) {
  let value = index + 1
  let result = ''
  while (value > 0) {
    value -= 1
    result = String.fromCharCode(65 + (value % 26)) + result
    value = Math.floor(value / 26)
  }
  return result || 'A'
}
