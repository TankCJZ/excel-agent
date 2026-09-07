import type { WorkbookAnalysis, WorkbookSummary } from '#agent/contracts'
import {
  summarizeWorkbookSheets,
  type CellValue,
  type WorkbookDataSheet
} from '#agent/workbook/analysis'
import type { ExcelWorkflowResult } from '#agent/types'
import { attachmentContentDisposition } from '#agent/http/fileDownload'
import { loadWorkbookContextCache } from '#agent/workbook/context'
import { hasUncalculatedXlsxFormulas } from '#agent/workbook/formulaCalculation'
import { assertReadableWorkbook } from '#agent/workbook/readSafety'

export type WorkbookAnalysisContext = {
  summary: WorkbookSummary
  sheets: WorkbookDataSheet[]
}

export async function createSampleWorkbook() {
  const XLSX = await importSheetJs()
  const rows: CellValue[][] = [
    ['Month', 'Region', 'Sales', 'Cost', 'Orders'],
    ['January', 'North', 12800, 7800, 84],
    ['January', 'South', 10400, 6500, 72],
    ['February', 'North', 14600, 8200, 91],
    ['February', 'South', 11900, 7100, 79],
    ['March', 'North', 16300, 8800, 104],
    ['March', 'South', 13200, 7600, 88]
  ]
  const workbook = XLSX.utils.book_new()
  const worksheet = XLSX.utils.aoa_to_sheet(rows)
  worksheet['!cols'] = [
    { wch: 12 },
    { wch: 14 },
    { wch: 14 },
    { wch: 14 },
    { wch: 12 }
  ]
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Sales Data')

  return {
    bytes: toUint8Array(XLSX.write(workbook, { type: 'array', bookType: 'xlsx' })),
    preview: rows
  }
}

export async function readWorkbookContext(bucket: R2Bucket, fileKey: string): Promise<WorkbookAnalysisContext> {
  const object = await bucket.get(fileKey)
  if (!object) {
    throw new Error('Workbook object was not found in R2')
  }

  return parseWorkbookBytes(await object.arrayBuffer())
}

export async function parseWorkbookBytes(bytes: ArrayBuffer | Uint8Array, options: { pendingCalculation?: boolean } = {}): Promise<WorkbookAnalysisContext> {
  const XLSX = await importSheetJs()
  const rawBytes = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes)
  let pendingCalculation = hasUncalculatedXlsxFormulas(rawBytes) || options.pendingCalculation === true
  const workbook = XLSX.read(bytes, {
    type: 'array',
    cellDates: true,
    cellFormula: true,
    sheetStubs: true,
    cellNF: false,
    cellStyles: false
  })
  assertReadableWorkbook(workbook)
  const formulaCells = workbook.SheetNames.flatMap(name => Object.entries(workbook.Sheets[name]!)
    .filter(([address, cell]) => !address.startsWith('!') && (typeof cell.f === 'string' || typeof cell.F === 'string'))
    .map(([, cell]) => cell as import('xlsx').CellObject))
  pendingCalculation ||= formulaCells.some(cell => cell.v === undefined || cell.t === 'z')
  if (pendingCalculation) {
    for (const cell of formulaCells) {
      cell.t = 's'
      cell.v = cell.f ? `=${cell.f}` : '[Formula pending calculation]'
      delete cell.w
    }
  }
  const visibleSheetNames = workbook.SheetNames.filter((name) => {
    if (name.startsWith('__')) return false
    const sheetIndex = workbook.SheetNames.indexOf(name)
    return !workbook.Workbook?.Sheets?.[sheetIndex]?.Hidden
  })
  const readableSheetNames = visibleSheetNames.length ? visibleSheetNames : workbook.SheetNames
  const sheets = readableSheetNames.map((name) => ({
    name,
    rows: normalizeRows(XLSX.utils.sheet_to_json<unknown[]>(workbook.Sheets[name]!, {
      header: 1,
      defval: null,
      raw: true
    }))
  }))

  if (!sheets.length) {
    throw new Error('Workbook does not contain a readable worksheet')
  }

  return {
    sheets,
    summary: {
      ...summarizeWorkbookSheets(sheets),
      ...(pendingCalculation ? { calculation: { status: 'pending' as const, formulaCells: formulaCells.length } } : {})
    }
  }
}

export async function loadWorkbookContext(
  bucket: R2Bucket,
  input: {
    workbookId: string
    sessionId: string
    fileKey: string
    contextR2Key: string | null
    contextVersion: number | null
    sourceEtag: string | null
  }
) {
  return loadWorkbookContextCache(bucket, input, parseWorkbookBytes)
}

export async function inspectWorkbook(bucket: R2Bucket, fileKey: string): Promise<WorkbookSummary> {
  return (await readWorkbookContext(bucket, fileKey)).summary
}

export async function processWorkbook(
  bucket: R2Bucket,
  input: {
    taskId: string
    inputR2Key: string
    inputName: string
    prompt: string
    analysis: WorkbookAnalysis
  }
): Promise<ExcelWorkflowResult> {
  const object = await bucket.get(input.inputR2Key)
  if (!object) {
    throw new Error('Input workbook was not found in R2')
  }

  const XLSX = await importSheetJs()
  const workbook = XLSX.read(await object.arrayBuffer(), {
    type: 'array',
    cellDates: true,
    cellFormula: true,
    cellStyles: true
  })
  if (!workbook.SheetNames.length) {
    throw new Error('Workbook does not contain a readable worksheet')
  }
  assertReadableWorkbook(workbook)

  const analysisRows = buildAnalysisSheet(input.inputName, input.analysis)
  const analysisSheet = XLSX.utils.aoa_to_sheet(analysisRows)
  analysisSheet['!cols'] = [{ wch: 24 }, { wch: 64 }, { wch: 20 }, { wch: 20 }]
  const analysisSheetName = uniqueSheetName(workbook, input.analysis.question)
  XLSX.utils.book_append_sheet(workbook, analysisSheet, analysisSheetName)

  if (input.analysis.chart) {
    const chartRows: CellValue[][] = [
      [input.analysis.chart.xAxisLabel, ...input.analysis.chart.series.map((series) => series.name)]
    ]
    input.analysis.chart.categories.forEach((category, index) => {
      chartRows.push([
        category,
        ...input.analysis.chart!.series.map((series) => series.data[index] ?? null)
      ])
    })
    const chartDataSheet = XLSX.utils.aoa_to_sheet(chartRows)
    chartDataSheet['!cols'] = [{ wch: 24 }, ...input.analysis.chart.series.map(() => ({ wch: 18 }))]
    const chartSheetName = uniqueSheetName(workbook, input.analysis.chart.title)
    XLSX.utils.book_append_sheet(workbook, chartDataSheet, chartSheetName)
  }

  const summary = summarizeWorkbookSheets(workbook.SheetNames.map((name) => ({
    name,
    rows: XLSX.utils.sheet_to_json<CellValue[]>(workbook.Sheets[name]!, {
      header: 1,
      defval: null,
      raw: true
    })
  })))

  const outputName = toOutputName(input.inputName)
  const outputR2Key = `agent-demo/results/${input.taskId}/${outputName}`
  const bytes = toUint8Array(XLSX.write(workbook, {
    type: 'array',
    bookType: 'xlsx',
    compression: true,
    cellStyles: true
  }))

  await bucket.put(outputR2Key, bytes, {
    httpMetadata: {
      contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      contentDisposition: attachmentContentDisposition(outputName)
    },
    customMetadata: {
      taskId: input.taskId,
      sourceKey: input.inputR2Key,
      generatedBy: 'excelgen-mastra-agent'
    }
  })

  return {
    kind: 'workbook_export',
    taskId: input.taskId,
    outputR2Key,
    outputName,
    summary,
    analysis: input.analysis,
    preview: analysisRows.slice(0, 18),
    previewSheetName: analysisSheetName,
    previewRowCount: Math.max(0, analysisRows.length - 1),
    sizeBytes: bytes.byteLength
  }
}

function buildAnalysisSheet(inputName: string, analysis: WorkbookAnalysis): CellValue[][] {
  const rows: CellValue[][] = [
    ['ExcelGen', inputName],
    [analysis.question, null]
  ]

  if (analysis.table) {
    rows.push([null, null], [...analysis.table.columns])
    rows.push(...analysis.table.rows)
  }
  if (analysis.metrics.length) {
    rows.push([null, null])
    for (const metric of analysis.metrics) rows.push([metric.label || metric.field || metric.operation, metric.value])
  }
  rows.push([null, null])
  for (const item of analysis.evidence) rows.push([`${item.sheet}!${item.range}`])
  return rows
}

function uniqueSheetName(workbook: import('xlsx').WorkBook, preferred: string) {
  const normalized = preferred
    .replace(/[\\/?*\[\]:]+/g, ' ')
    .replace(/[\u0000-\u001f]/g, '')
    .replace(/\s+/g, ' ')
    .replace(/^'+|'+$/g, '')
    .trim()
  const base = (normalized || 'ExcelGen').slice(0, 31)
  if (!workbook.SheetNames.includes(base)) return base
  for (let index = 2; index < 1000; index += 1) {
    const suffix = ` (${index})`
    const candidate = `${base.slice(0, 31 - suffix.length)}${suffix}`
    if (!workbook.SheetNames.includes(candidate)) return candidate
  }
  throw new Error('Workbook contains too many sheets with the same name')
}

function normalizeRows(rows: unknown[][]): CellValue[][] {
  return rows.map((row) => row.map(normalizeCell))
}

function normalizeCell(value: unknown): CellValue {
  if (value === null || typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean') {
    return typeof value === 'number' && !Number.isFinite(value) ? null : value
  }
  if (value === undefined) return null
  if (value instanceof Date) return value.toISOString()
  return String(value)
}

function toOutputName(inputName: string) {
  const base = inputName.replace(/\.(xlsx|xls|csv)$/i, '').replace(/[^\p{L}\p{N}_-]+/gu, '-').slice(0, 80)
  return `${base || 'workbook'}-excelgen-analysis.xlsx`
}

function toUint8Array(value: ArrayBuffer | Uint8Array) {
  return value instanceof Uint8Array ? value : new Uint8Array(value)
}

async function importSheetJs() {
  return import('xlsx')
}
