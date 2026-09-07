import { MUTATION_LIMITS } from '#shared/agent/workbookMutation'
import { cellAddress, parseAddress, type CellRange } from '#agent/workbook/mutation/coordinates'
import { requireMutation, WorkbookMutationError } from '#agent/workbook/mutation/errors'

// Deliberately bounded Excel grammar. Every character must be consumed; no
// evaluation, dynamic references, external names, UDFs or model-supplied caches.
const FUNCTIONS: Record<string, [number, number]> = {
  SUM: [1, 255], AVERAGE: [1, 255], MIN: [1, 255], MAX: [1, 255], COUNT: [1, 255], COUNTA: [1, 255],
  COUNTBLANK: [1, 1], SUMIF: [2, 3], SUMIFS: [3, 255], COUNTIF: [2, 2], COUNTIFS: [2, 254],
  AVERAGEIF: [2, 3], AVERAGEIFS: [3, 255], IF: [2, 3], IFERROR: [2, 2], IFNA: [2, 2],
  AND: [1, 255], OR: [1, 255], NOT: [1, 1], ABS: [1, 1], ROUND: [2, 2], ROUNDUP: [2, 2], ROUNDDOWN: [2, 2],
  INT: [1, 1], MOD: [2, 2], POWER: [2, 2], SQRT: [1, 1], PRODUCT: [1, 255], SUMPRODUCT: [1, 255],
  VLOOKUP: [3, 4], HLOOKUP: [3, 4], XLOOKUP: [3, 6], INDEX: [2, 4], MATCH: [2, 3],
  LEFT: [1, 2], RIGHT: [1, 2], MID: [3, 3], LEN: [1, 1], TRIM: [1, 1], UPPER: [1, 1], LOWER: [1, 1],
  CONCATENATE: [1, 255], CONCAT: [1, 255], TEXTJOIN: [3, 255], TEXT: [2, 2], VALUE: [1, 1],
  SUBSTITUTE: [3, 4], REPLACE: [4, 4], FIND: [2, 3], SEARCH: [2, 3], EXACT: [2, 2],
  DATE: [3, 3], YEAR: [1, 1], MONTH: [1, 1], DAY: [1, 1], DAYS: [2, 2], TODAY: [0, 0], NOW: [0, 0],
  ISBLANK: [1, 1], ISNUMBER: [1, 1], ISTEXT: [1, 1], ISERROR: [1, 1], ISNA: [1, 1],
  ROW: [0, 1], COLUMN: [0, 1], ROWS: [1, 1], COLUMNS: [1, 1]
}
const FUTURE_FUNCTIONS = new Set(['XLOOKUP', 'CONCAT', 'TEXTJOIN'])
type Token = { kind: 'string' | 'sheet' | 'word' | 'number' | 'symbol'; text: string; start: number; end: number }
type Reference = { sheet: string; range: CellRange }
type ReferenceToken = { start: number; end: number; text: string }
export type ParsedFormula = { formula: string; references: Reference[]; referenceTokens: ReferenceToken[] }

function tokenize(text: string): Token[] {
  const tokens: Token[] = []
  let position = 0
  while (position < text.length) {
    if (/\s/.test(text[position]!)) { position++; continue }
    const start = position
    const quote = text[position]
    if (quote === '"' || quote === "'") {
      position++
      let closed = false
      while (position < text.length) {
        if (text[position++] === quote) {
          if (text[position] === quote) position++
          else { closed = true; break }
        }
      }
      requireMutation(closed, 'INVALID_FORMULA', 'Formula contains an unterminated string or sheet name.')
      tokens.push({ kind: quote === '"' ? 'string' : 'sheet', text: text.slice(start, position), start, end: position })
      continue
    }
    const remaining = text.slice(position)
    const numeric = /^(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?/.exec(remaining)
    const word = /^\$?[\p{L}_][\p{L}\p{N}_.$]*/u.exec(remaining)
    const symbol = /^(?:<>|<=|>=|[+\-*/^&=<>%(),!:])/.exec(remaining)
    const match = numeric || word || symbol
    requireMutation(match, 'INVALID_FORMULA', 'Formula contains an unsupported character or external reference.')
    position += match[0].length
    tokens.push({ kind: numeric ? 'number' : word ? 'word' : 'symbol', text: match[0], start, end: position })
    requireMutation(tokens.length <= 768, 'FORMULA_LIMIT', 'Formula is too complex.')
  }
  return tokens
}

export function parseSafeFormula(raw: string, currentSheet: string, sheetNames: string[]): ParsedFormula {
  requireMutation(raw.startsWith('=') && raw.length <= MUTATION_LIMITS.formulaLength && !/[\x00-\x08\x0b\x0c\x0e-\x1f]/.test(raw), 'INVALID_FORMULA', 'Formula must start with = and fit the supported length.')
  const formula = raw.slice(1)
  const tokens = tokenize(formula)
  let index = 0
  let depth = 0
  const references: Reference[] = []
  const referenceTokens: ReferenceToken[] = []
  const replacements: Array<{ start: number; end: number; text: string }> = []
  const names = new Map(sheetNames.map(name => [name.toLocaleLowerCase('en-US'), name]))
  const peek = () => tokens[index]
  const take = (text: string) => {
    requireMutation(peek()?.text === text, 'INVALID_FORMULA', `Expected ${text} in formula.`)
    return tokens[index++]!
  }
  const ref = (sheet: string) => {
    const first = tokens[index++]
    requireMutation(first?.kind === 'word', 'INVALID_FORMULA', 'Expected a cell reference.')
    const start = parseAddress(first.text)
    referenceTokens.push(first)
    let end = start
    if (peek()?.text === ':') {
      index++
      const last = tokens[index++]
      requireMutation(last?.kind === 'word', 'INVALID_FORMULA', 'Expected a range end cell.')
      end = parseAddress(last.text)
      referenceTokens.push(last)
    }
    requireMutation(start.col <= end.col && start.row <= end.row, 'INVALID_FORMULA', 'Formula range is reversed.')
    references.push({ sheet, range: { start, end } })
  }
  const primary = () => {
    requireMutation(++depth <= 48, 'FORMULA_LIMIT', 'Formula nesting is too deep.')
    const token = peek()
    requireMutation(token, 'INVALID_FORMULA', 'Incomplete formula expression.')
    if (token.text === '+' || token.text === '-') { index++; primary() }
    else if (token.text === '(') { index++; expression(0); take(')') }
    else if (tokens[index + 1]?.text === '!') {
      requireMutation(token.kind === 'sheet' || token.kind === 'word', 'INVALID_FORMULA', 'Invalid sheet name.')
      const supplied = token.kind === 'sheet' ? token.text.slice(1, -1).replaceAll("''", "'") : token.text
      const name = names.get(supplied.toLocaleLowerCase('en-US'))
      requireMutation(name, 'INVALID_FORMULA', `Referenced worksheet was not found: ${supplied}`)
      index += 2
      ref(name)
    } else if (token.kind === 'number' || token.kind === 'string') index++
    else if (token.kind === 'word' && tokens[index + 1]?.text === '(') {
      const name = token.text.toUpperCase().replace(/^_XLFN\./, '')
      const limits = FUNCTIONS[name]
      requireMutation(limits, 'UNSUPPORTED_FORMULA', `Function ${name} is not supported for safe workbook editing yet.`)
      replacements.push({ start: token.start, end: token.end, text: `${FUTURE_FUNCTIONS.has(name) ? '_xlfn.' : ''}${name}` })
      index += 2
      let args = 0
      if (peek()?.text !== ')') {
        do {
          if (args > 0) take(',')
          expression(0)
          args++
        } while (peek()?.text === ',')
      }
      take(')')
      requireMutation(args >= limits[0] && args <= limits[1], 'INVALID_FORMULA', `${name} has an invalid number of arguments.`)
      if (['SUMIFS', 'AVERAGEIFS'].includes(name)) requireMutation(args % 2 === 1, 'INVALID_FORMULA', `${name} requires criteria pairs.`)
      if (name === 'COUNTIFS') requireMutation(args % 2 === 0, 'INVALID_FORMULA', 'COUNTIFS requires criteria pairs.')
    } else if (token.kind === 'word' && /^(TRUE|FALSE)$/i.test(token.text)) index++
    else ref(currentSheet)
    while (peek()?.text === '%') index++
    depth--
  }
  const precedence: Record<string, number> = { '=': 1, '<>': 1, '<': 1, '>': 1, '<=': 1, '>=': 1, '&': 2, '+': 3, '-': 3, '*': 4, '/': 4, '^': 5 }
  const expression = (minimum: number) => {
    primary()
    while (peek() && (precedence[peek()!.text] ?? -1) >= minimum) {
      const operator = tokens[index++]!.text
      expression(precedence[operator]! + 1)
    }
  }
  expression(0)
  requireMutation(index === tokens.length, 'INVALID_FORMULA', 'Formula contains trailing or unsupported syntax.')
  let normalized = formula
  for (const replacement of replacements.sort((a, b) => b.start - a.start)) {
    normalized = normalized.slice(0, replacement.start) + replacement.text + normalized.slice(replacement.end)
  }
  return { formula: normalized, references, referenceTokens }
}

export function fillFormula(raw: string, currentSheet: string, sheetNames: string[], rowOffset: number, colOffset: number): string {
  const parsed = parseSafeFormula(raw, currentSheet, sheetNames)
  let formula = raw.slice(1)
  // Token offsets refer to the original expression, not the normalized names.
  for (const token of [...parsed.referenceTokens].reverse()) {
    const match = /^(\$?)([A-Z]{1,3})(\$?)([1-9]\d*)$/i.exec(token.text)!
    const point = parseAddress(token.text)
    point.col += match[1] ? 0 : colOffset
    point.row += match[3] ? 0 : rowOffset
    const shifted = /^([A-Z]+)(\d+)$/.exec(cellAddress(point))!
    formula = formula.slice(0, token.start) + `${match[1]}${shifted[1]}${match[3]}${shifted[2]}` + formula.slice(token.end)
  }
  return parseSafeFormula(`=${formula}`, currentSheet, sheetNames).formula
}

export function assertAcyclicFormulas(formulas: Array<{ sheet: string; address: string; formula: string }>, sheetNames: string[]) {
  requireMutation(formulas.length <= MUTATION_LIMITS.formulas, 'FORMULA_LIMIT', 'Workbook contains too many formulas for safe editing.')
  const bySheet = new Map<string, Array<{ key: string; row: number; col: number }>>()
  for (const cell of formulas) {
    const group = bySheet.get(cell.sheet) || []
    group.push({ key: `${cell.sheet}!${cell.address}`, ...parseAddress(cell.address) })
    bySheet.set(cell.sheet, group)
  }
  for (const group of bySheet.values()) group.sort((a, b) => a.row - b.row || a.col - b.col)
  let checks = 0
  const edges = new Map<string, Set<string>>()
  for (const cell of formulas) {
    const dependencies = new Set<string>()
    for (const ref of parseSafeFormula(`=${cell.formula}`, cell.sheet, sheetNames).references) {
      const group = bySheet.get(ref.sheet) || []
      let low = 0
      let high = group.length
      while (low < high) {
        const middle = (low + high) >>> 1
        if (group[middle]!.row < ref.range.start.row) low = middle + 1
        else high = middle
      }
      for (let i = low; i < group.length && group[i]!.row <= ref.range.end.row; i++) {
        const candidate = group[i]!
        requireMutation(++checks <= MUTATION_LIMITS.dependencyChecks, 'FORMULA_LIMIT', 'Formula dependency graph is too complex for safe editing.')
        if (candidate.row >= ref.range.start.row && candidate.row <= ref.range.end.row && candidate.col >= ref.range.start.col && candidate.col <= ref.range.end.col) dependencies.add(candidate.key)
      }
    }
    edges.set(`${cell.sheet}!${cell.address}`, dependencies)
  }
  // Kahn's algorithm avoids JS call-stack limits on long formula chains.
  const incoming = new Map<string, number>()
  const dependants = new Map<string, string[]>()
  for (const [key, dependencies] of edges) {
    incoming.set(key, dependencies.size)
    for (const dependency of dependencies) {
      const group = dependants.get(dependency) || []
      group.push(key)
      dependants.set(dependency, group)
    }
  }
  const ready = [...incoming].filter(([, count]) => count === 0).map(([key]) => key)
  for (let cursor = 0; cursor < ready.length; cursor++) {
    for (const key of dependants.get(ready[cursor]!) || []) {
      const count = incoming.get(key)! - 1
      incoming.set(key, count)
      if (count === 0) ready.push(key)
    }
  }
  if (ready.length !== formulas.length) throw new WorkbookMutationError('CIRCULAR_FORMULA', 'Formula changes create or retain a circular reference. No modified file was published.')
}
