import { requireMutation } from '#agent/workbook/mutation/errors'

export type Coordinate = { row: number; col: number }
export type CellRange = { start: Coordinate; end: Coordinate }

export function columnName(col: number): string {
  requireMutation(Number.isInteger(col) && col >= 1 && col <= 16384, 'INVALID_RANGE', 'Column is outside Excel limits.')
  let name = ''
  for (let value = col; value > 0; value = Math.floor((value - 1) / 26)) name = String.fromCharCode(65 + (value - 1) % 26) + name
  return name
}

export function cellAddress({ row, col }: Coordinate): string {
  requireMutation(Number.isInteger(row) && row >= 1 && row <= 1048576, 'INVALID_RANGE', 'Row is outside Excel limits.')
  return `${columnName(col)}${row}`
}

export function parseAddress(address: string): Coordinate {
  const match = /^\$?([A-Z]{1,3})\$?([1-9]\d{0,6})$/i.exec(address)
  requireMutation(match, 'INVALID_RANGE', `Invalid cell address: ${address}`)
  const col = [...match[1]!.toUpperCase()].reduce((value, char) => value * 26 + char.charCodeAt(0) - 64, 0)
  const result = { col, row: Number(match[2]) }
  cellAddress(result)
  return result
}

export function parseRange(value: string): CellRange {
  const parts = value.split(':')
  requireMutation(parts.length <= 2, 'INVALID_RANGE', 'Expected a single rectangular range.')
  const start = parseAddress(parts[0]!)
  const end = parseAddress(parts[1] || parts[0]!)
  requireMutation(start.row <= end.row && start.col <= end.col, 'INVALID_RANGE', 'Range must run from top-left to bottom-right.')
  return { start, end }
}

export function containsCell(range: CellRange, cell: Coordinate): boolean {
  return cell.row >= range.start.row && cell.row <= range.end.row && cell.col >= range.start.col && cell.col <= range.end.col
}

export function rangeSize(range: CellRange): number {
  return (range.end.row - range.start.row + 1) * (range.end.col - range.start.col + 1)
}
