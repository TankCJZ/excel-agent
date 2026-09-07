import {
  MUTATION_LIMITS, workbookMutationInputSchema,
  type MutationChanges, type MutationPreviewCell, type MutationValue, type WorkbookMutationInput
} from '#shared/agent/workbookMutation'
import { cellAddress, columnName, containsCell, parseAddress, parseRange, rangeSize, type CellRange } from '#agent/workbook/mutation/coordinates'
import { assertAcyclicFormulas, fillFormula, parseSafeFormula } from '#agent/workbook/mutation/formulas'
import { requireMutation } from '#agent/workbook/mutation/errors'
import {
  XlsxPackage, NS, REL_NS, OFFICE_REL_NS, CONTENT_NS,
  child, children, descendants, element, parseXml, relationshipPath,
  type XmlDocument, type XmlElement
} from '#agent/workbook/mutation/package'
import { dateSerial, styleWithFormat } from '#agent/workbook/mutation/styles'
import { strToU8 } from 'fflate'

class Worksheet {
  readonly doc: XmlDocument
  readonly cells = new Map<string, XmlElement>()
  readonly rows = new Map<number, XmlElement>()
  readonly data: XmlElement
  readonly merges: CellRange[]
  lastColumn = 0
  lastRow = 0
  changed = false
  readonly pkg: XlsxPackage
  readonly name: string
  readonly path: string
  readonly hidden: boolean

  constructor(pkg: XlsxPackage, name: string, path: string, hidden: boolean) {
    this.pkg = pkg
    this.name = name
    this.path = path
    this.hidden = hidden
    this.doc = pkg.xml(path)
    requireMutation(this.doc.documentElement?.namespaceURI === NS && this.doc.documentElement?.localName === 'worksheet', 'UNSUPPORTED_WORKBOOK', 'Only standard worksheets can be edited.')
    const data = child(this.doc.documentElement!, 'sheetData')
    requireMutation(data, 'INVALID_WORKBOOK', 'Worksheet data is missing.')
    this.data = data
    for (const row of children(data, 'row')) {
      const index = Number(row.getAttribute('r'))
      requireMutation(Number.isInteger(index) && index > 0 && index <= 1048576 && !this.rows.has(index), 'INVALID_WORKBOOK', 'Worksheet contains invalid or duplicate row coordinates.')
      this.rows.set(index, row)
      for (const cell of children(row, 'c')) {
        requireMutation(!cell.hasAttribute('cm') && !cell.hasAttribute('vm'), 'UNSUPPORTED_WORKBOOK', 'Cells with dynamic-array or rich-value metadata cannot be safely edited yet.')
        const address = cell.getAttribute('r') || ''
        const point = parseAddress(address)
        requireMutation(point.row === index && !this.cells.has(address), 'INVALID_WORKBOOK', 'Worksheet contains invalid or duplicate cell coordinates.')
        this.cells.set(address, cell)
        this.lastColumn = Math.max(point.col, this.lastColumn)
        this.lastRow = Math.max(point.row, this.lastRow)
      }
    }
    this.merges = descendants(this.doc, 'mergeCell').map(node => parseRange(node.getAttribute('ref') || ''))
    for (const range of this.merges) {
      this.lastColumn = Math.max(this.lastColumn, range.end.col)
      this.lastRow = Math.max(this.lastRow, range.end.row)
    }
    requireMutation(!descendants(this.doc, 'f').some(node => ['array', 'dataTable'].includes(node.getAttribute('t') || '')), 'UNSUPPORTED_WORKBOOK', 'Array and data-table formulas are not supported for safe editing yet.')
  }

  assertEditable(address: string) {
    requireMutation(!descendants(this.doc, 'sheetProtection').length, 'PROTECTED_SHEET', `Worksheet ${this.name} is protected.`)
    const point = parseAddress(address)
    for (const range of this.merges) {
      requireMutation(!containsCell(range, point) || cellAddress(range.start) === address, 'MERGED_CELL', `Only the top-left cell of a merged range can be edited: ${this.name}!${address}`)
    }
  }

  getCell(address: string): XmlElement {
    this.assertEditable(address)
    const existing = this.cells.get(address)
    if (existing) return existing
    const point = parseAddress(address)
    let row = this.rows.get(point.row)
    if (!row) {
      row = element(this.doc, 'row', { r: String(point.row) })
      this.rows.set(point.row, row)
      this.data.appendChild(row)
    }
    const cell = element(this.doc, 'c', { r: address })
    row.appendChild(cell)
    this.cells.set(address, cell)
    this.lastColumn = Math.max(point.col, this.lastColumn)
    this.lastRow = Math.max(point.row, this.lastRow)
    return cell
  }

  flush() {
    if (!this.changed) return
    // Sort once, after all operations; do not scan the entire worksheet per cell.
    for (const [index, row] of [...this.rows].sort(([a], [b]) => a - b)) {
      row.removeAttribute('spans')
      for (const cell of children(row, 'c').sort((a, b) => parseAddress(a.getAttribute('r')!).col - parseAddress(b.getAttribute('r')!).col)) row.appendChild(cell)
      this.data.appendChild(row)
      this.lastRow = Math.max(this.lastRow, index)
    }
    const root = this.doc.documentElement!
    let dimension = child(root, 'dimension')
    if (!dimension) {
      dimension = element(this.doc, 'dimension')
      const first = children(root).find(node => node.localName !== 'sheetPr')
      root.insertBefore(dimension, first || null)
    }
    dimension.setAttribute('ref', `A1:${cellAddress({ col: Math.max(1, this.lastColumn), row: Math.max(1, this.lastRow) })}`)
    this.pkg.update(this.path, this.doc)
  }
}

export type WorkbookEditContext = {
  schemaVersion: 1
  sheets: Array<{
    name: string; hidden: boolean; lastColumn: string | null; lastRow: number; mergedRanges: string[]
    preview: Array<{ address: string; value: string | number | boolean | null; formula?: string }>
  }>
}

function workbookSheets(pkg: XlsxPackage): Worksheet[] {
  const workbook = pkg.xml('xl/workbook.xml')
  const rels = descendants(pkg.xml('xl/_rels/workbook.xml.rels'), 'Relationship', REL_NS)
  const entries = descendants(workbook, 'sheet')
  requireMutation(entries.length > 0 && entries.length <= MUTATION_LIMITS.sheets, 'FILE_LIMIT', 'Workbook worksheet count exceeds safe editing limits.')
  const names = new Set<string>()
  const paths = new Set<string>()
  let count = 0
  return entries.map(entry => {
    const name = entry.getAttribute('name') || ''
    const id = entry.getAttributeNS(OFFICE_REL_NS, 'id')
    const rel = rels.find(node => node.getAttribute('Id') === id)
    requireMutation(name && rel?.getAttribute('Type') === `${OFFICE_REL_NS}/worksheet` && !names.has(name.toLowerCase()), 'INVALID_WORKBOOK', 'Workbook contains duplicate or unsupported sheets.')
    const path = relationshipPath(rel!.getAttribute('Target') || '')
    requireMutation(!paths.has(path), 'INVALID_WORKBOOK', 'Worksheet relationship is duplicated.')
    paths.add(path)
    names.add(name.toLowerCase())
    const sheet = new Worksheet(pkg, name, path, ['hidden', 'veryHidden'].includes(entry.getAttribute('state') || ''))
    count += sheet.cells.size
    requireMutation(count <= MUTATION_LIMITS.physicalCells, 'FILE_LIMIT', 'Workbook contains too many physical cells for safe editing.')
    return sheet
  })
}

function strings(pkg: XlsxPackage): string[] {
  return pkg.has('xl/sharedStrings.xml') ? descendants(pkg.xml('xl/sharedStrings.xml'), 'si').map(node => descendants(node, 't').map(item => item.textContent || '').join('')) : []
}

function valueOf(cell: XmlElement | undefined, shared: string[], formula?: string): MutationPreviewCell['value'] {
  if (!cell || formula !== undefined) return null
  const value = child(cell, 'v')?.textContent || ''
  switch (cell.getAttribute('t')) {
    case 's': return shared[Number(value)] ?? null
    case 'inlineStr': return descendants(cell, 't').map(node => node.textContent || '').join('')
    case 'b': return value === '1'
    case 'str': case 'e': case 'd': return value
    default: return value !== '' && Number.isFinite(Number(value)) ? Number(value) : null
  }
}

function formulaMap(sheets: Worksheet[]): Map<string, string> {
  const result = new Map<string, string>()
  const names = sheets.map(sheet => sheet.name)
  for (const sheet of sheets) {
    const masters = new Map<string, { address: string; formula: string }>()
    for (const [address, cell] of sheet.cells) {
      const f = child(cell, 'f')
      if (f?.getAttribute('t') === 'shared' && f.textContent) masters.set(f.getAttribute('si') || '', { address, formula: f.textContent })
    }
    for (const [address, cell] of sheet.cells) {
      const f = child(cell, 'f')
      if (!f) continue
      let formula = f.textContent || ''
      if (f.getAttribute('t') === 'shared' && !formula) {
        const master = masters.get(f.getAttribute('si') || '')
        requireMutation(master, 'INVALID_WORKBOOK', 'Shared formula has no master cell.')
        const from = parseAddress(master.address)
        const to = parseAddress(address)
        formula = fillFormula(`=${master.formula}`, sheet.name, names, to.row - from.row, to.col - from.col)
      }
      parseSafeFormula(`=${formula}`, sheet.name, names)
      result.set(`${sheet.name}!${address}`, formula)
    }
  }
  return result
}

function preflightTables(pkg: XlsxPackage) {
  for (const name of pkg.names().filter(name => /^xl\/tables\/[^/]+\.xml$/.test(name))) {
    const doc = pkg.xml(name)
    requireMutation(!descendants(doc, 'table').length, 'UNSUPPORTED_WORKBOOK', 'Excel Table objects cannot be safely expanded or edited yet. Convert the table to a range in an XLSX copy before editing; the original remains unchanged.')
  }
}

export function inspectWorkbookForEditing(bytes: Uint8Array): WorkbookEditContext {
  const pkg = new XlsxPackage(bytes)
  preflightTables(pkg)
  const sheets = workbookSheets(pkg)
  const shared = strings(pkg)
  const formulas = formulaMap(sheets)
  return {
    schemaVersion: 1,
    sheets: sheets.map(sheet => ({
      name: sheet.name, hidden: sheet.hidden, lastColumn: sheet.lastColumn ? columnName(sheet.lastColumn) : null, lastRow: sheet.lastRow,
      mergedRanges: sheet.merges.map(range => `${cellAddress(range.start)}:${cellAddress(range.end)}`),
      preview: [...sheet.cells].sort(([a], [b]) => parseAddress(a).row - parseAddress(b).row || parseAddress(a).col - parseAddress(b).col).slice(0, 100).map(([address, cell]) => {
        const formula = formulas.get(`${sheet.name}!${address}`)
        return { address, value: previewValue(valueOf(cell, shared, formula)), ...(formula !== undefined ? { formula: `=${formula}` } : {}) }
      })
    }))
  }
}

function addSheet(pkg: XlsxPackage, name: string): Worksheet {
  requireMutation(!name.startsWith("'") && !name.endsWith("'"), 'INVALID_SHEET', 'Worksheet names cannot start or end with an apostrophe.')
  const workbook = pkg.xml('xl/workbook.xml')
  const existingSheets = descendants(workbook, 'sheet')
  requireMutation(existingSheets.length < MUTATION_LIMITS.sheets && !existingSheets.some(node => node.getAttribute('name')?.toLowerCase() === name.toLowerCase()), 'INVALID_SHEET', 'Worksheet name already exists or the sheet limit was reached.')
  let suffix = Math.max(0, ...existingSheets.map(node => Number(node.getAttribute('sheetId')))) + 1
  while (pkg.has(`xl/worksheets/sheet${suffix}.xml`)) suffix++
  const path = `xl/worksheets/sheet${suffix}.xml`
  const relationships = pkg.xml('xl/_rels/workbook.xml.rels')
  const ids = new Set(descendants(relationships, 'Relationship', REL_NS).map(node => node.getAttribute('Id')))
  let id = `rIdMutation${suffix}`
  while (ids.has(id)) id += '_'
  relationships.documentElement!.appendChild(element(relationships, 'Relationship', { Id: id, Type: `${OFFICE_REL_NS}/worksheet`, Target: `worksheets/sheet${suffix}.xml` }, undefined, REL_NS))
  pkg.update('xl/_rels/workbook.xml.rels', relationships)
  const entry = element(workbook, 'sheet', { name, sheetId: String(suffix) })
  entry.setAttributeNS(OFFICE_REL_NS, 'r:id', id)
  child(workbook.documentElement!, 'sheets')!.appendChild(entry)
  pkg.update('xl/workbook.xml', workbook)
  const contents = pkg.xml('[Content_Types].xml')
  contents.documentElement!.appendChild(element(contents, 'Override', { PartName: `/${path}`, ContentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml' }, undefined, CONTENT_NS))
  pkg.update('[Content_Types].xml', contents)
  pkg.update(path, parseXml(strToU8(`<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="${NS}"><dimension ref="A1"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/><sheetData/></worksheet>`)))
  return new Worksheet(pkg, name, path, false)
}

function clearCellValue(cell: XmlElement) {
  for (const node of children(cell)) if (['v', 'is', 'f'].includes(node.localName!)) cell.removeChild(node)
  cell.removeAttribute('t')
}

function writeValue(sheet: Worksheet, address: string, value: MutationValue, date1904: boolean) {
  const cell = sheet.getCell(address)
  clearCellValue(cell)
  if (typeof value === 'string') {
    requireMutation(!/[\x00-\x08\x0b\x0c\x0e-\x1f]/.test(value), 'INVALID_VALUE', 'Cell text contains unsupported control characters.')
    cell.setAttribute('t', 'inlineStr')
    const inline = element(sheet.doc, 'is')
    const text = element(sheet.doc, 't', {}, value)
    text.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve')
    inline.appendChild(text)
    cell.appendChild(inline)
  } else if (value !== null) {
    const numeric = typeof value === 'object' ? dateSerial(value, date1904) : value
    cell.setAttribute('t', typeof numeric === 'boolean' ? 'b' : 'n')
    cell.appendChild(element(sheet.doc, 'v', {}, String(typeof numeric === 'boolean' ? Number(numeric) : numeric)))
    if (typeof value === 'object') cell.setAttribute('s', styleWithFormat(sheet.pkg, cell.getAttribute('s'), 'date'))
  }
  sheet.changed = true
}

function invalidateCalculation(pkg: XlsxPackage, sheets: Worksheet[], formulas: Map<string, string>) {
  for (const sheet of sheets) {
    for (const [address, cell] of sheet.cells) {
      const formula = formulas.get(`${sheet.name}!${address}`)
      if (formula === undefined) continue
      clearCellValue(cell)
      cell.appendChild(element(sheet.doc, 'f', {}, formula))
      // Deliberately no fabricated <v>0</v>. Excel calculates when opened.
      sheet.changed = true
    }
  }
  const workbook = pkg.xml('xl/workbook.xml')
  let calc = child(workbook.documentElement!, 'calcPr')
  if (!calc) {
    calc = element(workbook, 'calcPr')
    const after = new Set(['oleSize', 'customWorkbookViews', 'pivotCaches', 'smartTagPr', 'smartTagTypes', 'webPublishing', 'fileRecoveryPr', 'webPublishObjects', 'extLst'])
    workbook.documentElement!.insertBefore(calc, children(workbook.documentElement!).find(node => after.has(node.localName!)) || null)
  }
  for (const [key, value] of Object.entries({ calcId: '0', calcMode: 'auto', fullCalcOnLoad: '1', forceFullCalc: '1', calcOnSave: '1' })) calc.setAttribute(key, value)
  pkg.update('xl/workbook.xml', workbook)
  const rels = pkg.xml('xl/_rels/workbook.xml.rels')
  for (const rel of descendants(rels, 'Relationship', REL_NS)) {
    if (rel.getAttribute('Type') !== `${OFFICE_REL_NS}/calcChain`) continue
    const path = relationshipPath(rel.getAttribute('Target')!)
    pkg.remove(path)
    rel.parentNode!.removeChild(rel)
    const types = pkg.xml('[Content_Types].xml')
    for (const type of descendants(types, 'Override', CONTENT_NS)) if (type.getAttribute('PartName') === `/${path}`) type.parentNode!.removeChild(type)
    pkg.update('[Content_Types].xml', types)
    pkg.update('xl/_rels/workbook.xml.rels', rels)
  }
}

// New-generation path uses the same XML cache invalidation and dependency
// validation as uploaded-file editing, without changing unrelated styles.
export function finalizeGeneratedWorkbookFormulas(bytes: Uint8Array) {
  const pkg = new XlsxPackage(bytes)
  const sheets = workbookSheets(pkg)
  const formulas = formulaMap(sheets)
  assertAcyclicFormulas(sheets.flatMap(sheet => [...sheet.cells.keys()].flatMap(address => {
    const formula = formulas.get(`${sheet.name}!${address}`)
    return formula === undefined ? [] : [{ sheet: sheet.name, address, formula }]
  })), sheets.map(sheet => sheet.name))
  if (!formulas.size) return bytes
  invalidateCalculation(pkg, sheets, formulas)
  for (const sheet of sheets) sheet.flush()
  return pkg.write()
}

export function applyWorkbookMutation(bytes: Uint8Array, untrustedInput: WorkbookMutationInput): { bytes: Uint8Array; changes: MutationChanges } {
  const input = workbookMutationInputSchema.parse(untrustedInput)
  const pkg = new XlsxPackage(bytes)
  preflightTables(pkg)
  const sheets = workbookSheets(pkg)
  const shared = strings(pkg)
  const formulas = formulaMap(sheets)
  const changes: MutationChanges = { changedCells: 0, formulaCells: 0, addedSheets: [], addedColumns: [], pendingCalculation: false, convertedFrom: null, sheets: [] }
  const written = new Set<string>()
  const formatted = new Set<string>()
  const touched = new Map<string, Set<string>>()
  const date1904 = ['true', '1'].includes(child(pkg.xml('xl/workbook.xml').documentElement!, 'workbookPr')?.getAttribute('date1904') || '')
  let targeted = 0
  const track = (sheet: Worksheet, address: string, overwrite: boolean, formatOnly = false) => {
    const key = `${sheet.name}!${address}`
    const collection = formatOnly ? formatted : written
    requireMutation(!collection.has(key), 'CONFLICTING_OPERATIONS', `Multiple operations target ${key}. Combine the changes before retrying.`)
    collection.add(key)
    requireMutation(++targeted <= MUTATION_LIMITS.changedCells, 'EDIT_LIMIT', 'Too many cells are targeted in one edit.')
    sheet.assertEditable(address)
    if (!formatOnly) {
      const cell = sheet.cells.get(address)
      requireMutation(overwrite || !cell || (!child(cell, 'f') && valueOf(cell, shared) === null), 'OVERWRITE_REQUIRED', `${key} already contains data. Ask the user to confirm replacing it before setting overwrite=true.`)
    }
  }
  const changed = (sheet: Worksheet, address: string) => {
    const group = touched.get(sheet.name) || new Set<string>()
    group.add(address)
    touched.set(sheet.name, group)
  }
  const writeMatrix = (sheet: Worksheet, start: string, values: MutationValue[][], overwrite: boolean) => {
    const origin = parseAddress(start)
    const width = values[0]?.length || 0
    requireMutation(values.every(row => row.length === width), 'INVALID_VALUES', 'Each input row must have the same number of cells.')
    for (let row = 0; row < values.length; row++) {
      for (let col = 0; col < width; col++) {
        const address = cellAddress({ row: origin.row + row, col: origin.col + col })
        track(sheet, address, overwrite)
        const value = values[row]![col]!
        const key = `${sheet.name}!${address}`
        if (!formulas.has(key) && (typeof value !== 'object' || value === null) && valueOf(sheet.cells.get(address), shared) === value) continue
        formulas.delete(key)
        writeValue(sheet, address, value, date1904)
        changed(sheet, address)
      }
    }
  }
  for (const operation of input.operations) {
    if (operation.kind === 'add_sheet') {
      const sheet = addSheet(pkg, operation.sheet)
      sheets.push(sheet)
      changes.addedSheets.push(sheet.name)
      writeMatrix(sheet, 'A1', operation.values, false)
      continue
    }
    const sheet = sheets.find(sheet => sheet.name === operation.sheet)
    requireMutation(sheet, 'SHEET_NOT_FOUND', `Worksheet not found: ${operation.sheet}. Inspect the workbook and use its exact sheet name.`)
    if (operation.kind === 'set_values') writeMatrix(sheet, operation.start, operation.values, operation.overwrite)
    else if (operation.kind === 'append_column') {
      for (const [address, cell] of sheet.cells) {
        if (parseAddress(address).row !== operation.headerRow) continue
        requireMutation(String(valueOf(cell, shared) ?? '').trim().toLowerCase() !== operation.title.toLowerCase(), 'DUPLICATE_COLUMN', `Column ${operation.title} already exists on the specified header row.`)
      }
      const column = columnName(sheet.lastColumn + 1)
      writeMatrix(sheet, `${column}${operation.headerRow}`, [[operation.title], ...operation.values.map(value => [value])], false)
      changes.addedColumns.push({ sheet: sheet.name, column, title: operation.title })
    } else {
      const range = parseRange(operation.range)
      requireMutation(rangeSize(range) <= MUTATION_LIMITS.changedCells, 'EDIT_LIMIT', 'Target range exceeds the per-edit cell limit.')
      for (let row = range.start.row; row <= range.end.row; row++) for (let col = range.start.col; col <= range.end.col; col++) {
        const address = cellAddress({ row, col })
        const formatOnly = operation.kind === 'set_number_format'
        track(sheet, address, formatOnly || operation.overwrite, formatOnly)
        if (operation.kind === 'set_formula') {
          const formula = fillFormula(operation.formula, sheet.name, sheets.map(sheet => sheet.name), row - range.start.row, col - range.start.col)
          const key = `${sheet.name}!${address}`
          if (formulas.get(key) === formula) continue
          const cell = sheet.getCell(address)
          clearCellValue(cell)
          cell.appendChild(element(sheet.doc, 'f', {}, formula))
          formulas.set(key, formula)
          changes.formulaCells++
          sheet.changed = true
          changed(sheet, address)
        } else {
          const cell = sheet.getCell(address)
          const style = styleWithFormat(pkg, cell.getAttribute('s'), operation.format)
          if (cell.getAttribute('s') === style) continue
          cell.setAttribute('s', style)
          sheet.changed = true
          changed(sheet, address)
        }
      }
    }
  }
  assertAcyclicFormulas(sheets.flatMap(sheet => [...sheet.cells.keys()].flatMap(address => {
    const formula = formulas.get(`${sheet.name}!${address}`)
    return formula !== undefined ? [{ sheet: sheet.name, address, formula }] : []
  })), sheets.map(sheet => sheet.name))
  changes.changedCells = [...touched.values()].reduce((sum, addresses) => sum + addresses.size, 0)
  changes.pendingCalculation = formulas.size > 0 && (changes.changedCells > 0 || changes.addedSheets.length > 0)
  if (changes.pendingCalculation) invalidateCalculation(pkg, sheets, formulas)
  for (const sheet of sheets) sheet.flush()
  for (const sheet of sheets.filter(sheet => touched.has(sheet.name) || changes.addedSheets.includes(sheet.name)).slice(0, 2)) {
    const points = [...(touched.get(sheet.name) || ['A1'])].map(parseAddress)
    const firstRow = Math.max(1, Math.min(...points.map(point => point.row)) - 1)
    const firstCol = Math.max(1, Math.min(...points.map(point => point.col)) - 2)
    const preview = { name: sheet.name, rows: [] as MutationChanges['sheets'][number]['rows'] }
    for (let row = firstRow; row <= Math.min(sheet.lastRow, firstRow + 19); row++) {
      const cells: MutationPreviewCell[] = []
      for (let col = firstCol; col <= Math.min(sheet.lastColumn, firstCol + 11); col++) {
        const address = cellAddress({ row, col })
        const formula = formulas.get(`${sheet.name}!${address}`)
        cells.push({ address, value: previewValue(valueOf(sheet.cells.get(address), shared, formula)), ...(formula !== undefined ? { formula: `=${formula}`, pendingCalculation: changes.pendingCalculation } : {}) })
      }
      preview.rows.push({ row, cells })
    }
    changes.sheets.push(preview)
  }
  const output = pkg.modified.size ? pkg.write() : bytes.slice()
  // Re-open using a separate parser in tests; runtime verifies the package and
  // the exact target values/formulas, without ever mutating the source bytes.
  const verified = new XlsxPackage(output)
  const checkedSheets = workbookSheets(verified)
  const checkedStrings = strings(verified)
  const checkedFormulas = formulaMap(checkedSheets)
  for (const [name, addresses] of touched) for (const address of addresses) {
    const sourceSheet = sheets.find(sheet => sheet.name === name)!
    const checked = checkedSheets.find(sheet => sheet.name === name)!
    const expectedFormula = formulas.get(`${name}!${address}`)
    requireMutation(checkedFormulas.get(`${name}!${address}`) === expectedFormula && valueOf(checked.cells.get(address), checkedStrings, expectedFormula) === valueOf(sourceSheet.cells.get(address), shared, expectedFormula), 'OUTPUT_VERIFICATION_FAILED', 'The generated workbook did not pass verification.')
  }
  return { bytes: output, changes }
}

function previewValue(value: MutationPreviewCell['value']) {
  return typeof value === 'string' && value.length > 240 ? `${value.slice(0, 240)}…` : value
}
