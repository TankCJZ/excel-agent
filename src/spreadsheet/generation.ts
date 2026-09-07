import { z } from 'zod'
import { spreadsheetRecordSchema } from '#agent/contracts'
import type {
  SpreadsheetColumn,
  SpreadsheetPlan,
  WebSearchResult,
  WebSearchSource,
  WorkbookAnalysis,
  SpreadsheetRecord
} from '#agent/contracts'
import { summarizeWorkbookSheets, type CellValue, type WorkbookDataSheet } from '#agent/workbook/analysis'
import type { AgentWorkerEnv, SpreadsheetGenerationResult, SpreadsheetWorkflowPayload } from '#agent/types'
import { attachmentContentDisposition } from '#agent/http/fileDownload'
import { runResponsesModel, extractResponseText } from '#agent/research/webSearch'
import { spreadsheetFormulas } from '#agent/spreadsheet/formulas'
import { finalizeGeneratedWorkbookFormulas } from '#agent/workbook/mutation/engine'
import { requireMutation } from '#agent/workbook/mutation/errors'

export type { SpreadsheetRecord } from '#agent/contracts'
type SheetJs = typeof import('xlsx')

export type PreparedSpreadsheetData = {
  records: SpreadsheetRecord[]
  sources: WebSearchSource[]
}

export async function prepareSpreadsheetData(
  env: AgentWorkerEnv,
  payload: SpreadsheetWorkflowPayload
): Promise<PreparedSpreadsheetData> {
  const research = await readResearchSnapshot(env.WORKBOOKS, payload.researchSnapshotR2Key)
  const maxRows = Math.min(payload.plan.maxRows, clampInteger(env.SPREADSHEET_MAX_ROWS, 1, 100, 50))
  const recordsKey = `agent-demo/tasks/${payload.taskId}/records.json`
  let records = await readRecordsSnapshot(env.WORKBOOKS, recordsKey)

  if (!records) {
    if (payload.plan.dataMode === 'conversation_data' || payload.plan.dataMode === 'model_knowledge') {
      throw new Error('Captured-record snapshot was not found in R2')
    }
    records = await resolveSpreadsheetRecords(env, payload, research, maxRows)
    await env.WORKBOOKS.put(recordsKey, JSON.stringify({ records }), {
      httpMetadata: { contentType: 'application/json; charset=utf-8' },
      customMetadata: { taskId: payload.taskId, kind: 'spreadsheet-records' }
    })
  }

  requireMutation(records.length <= maxRows || !records.some(record => Object.values(record).some(value => value && typeof value === 'object')), 'EDIT_LIMIT', 'Formula-enabled workbooks cannot be truncated. Reduce the number of rows before generating.')
  return {
    records: records.slice(0, maxRows),
    sources: research?.sources || []
  }
}

export async function writeSpreadsheetResult(
  env: AgentWorkerEnv,
  payload: SpreadsheetWorkflowPayload,
  prepared: PreparedSpreadsheetData
): Promise<SpreadsheetGenerationResult> {
  const generated = await buildSpreadsheetWorkbook(payload.plan, prepared.records, prepared.sources)
  const outputName = `${safeFileStem(payload.plan.title)}.xlsx`
  const outputR2Key = `agent-demo/results/${payload.taskId}/${outputName}`

  await env.WORKBOOKS.put(outputR2Key, generated.bytes, {
    httpMetadata: {
      contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      contentDisposition: attachmentContentDisposition(outputName)
    },
    customMetadata: {
      taskId: payload.taskId,
      generatedBy: 'excelgen-spreadsheet-workflow',
      dataMode: payload.plan.dataMode
    }
  })

  return {
    kind: 'spreadsheet_generation',
    taskId: payload.taskId,
    outputR2Key,
    outputName,
    summary: generated.summary,
    analysis: null,
    preview: generated.preview,
    previewSheetName: payload.plan.sheetName,
    previewRowCount: prepared.records.length,
    sizeBytes: generated.bytes.byteLength,
    columns: payload.plan.columns,
    sources: prepared.sources
  }
}

export async function generateSpreadsheet(
  env: AgentWorkerEnv,
  payload: SpreadsheetWorkflowPayload
): Promise<SpreadsheetGenerationResult> {
  return writeSpreadsheetResult(env, payload, await prepareSpreadsheetData(env, payload))
}

export async function buildSpreadsheetWorkbook(
  plan: SpreadsheetPlan,
  records: SpreadsheetRecord[],
  sources: WebSearchSource[] = []
) {
  const XLSX = await importSheetJs()
  const formulas = spreadsheetFormulas(plan, records)
  const mainRows: CellValue[][] = [
    plan.columns.map(column => column.label),
    ...records.map(record => plan.columns.map(column => {
      const value = record[column.key]
      return value && typeof value === 'object' ? value.formula : sanitizeSpreadsheetCell(value)
    }))
  ]
  const workbook = XLSX.utils.book_new()
  const mainSheet = XLSX.utils.aoa_to_sheet(mainRows)
  applyTablePresentation(XLSX, mainSheet, mainRows, plan.columns)
  for (const entry of formulas) {
    const cell = mainSheet[entry.address]!
    cell.t = 'n'
    cell.f = entry.formula
    delete cell.v
    delete cell.w
  }
  XLSX.utils.book_append_sheet(workbook, mainSheet, plan.sheetName)

  const sheets: WorkbookDataSheet[] = [{ name: plan.sheetName, rows: mainRows }]
  if (plan.includeSources && sources.length) {
    const sourceRows: CellValue[][] = [
      ['source_id', 'title', 'url', 'snippet'],
      ...sources.map(source => [source.id, source.title, source.url, source.snippet || ''])
    ]
    const sourceSheet = XLSX.utils.aoa_to_sheet(sourceRows)
    sourceSheet['!cols'] = [{ wch: 12 }, { wch: 36 }, { wch: 54 }, { wch: 70 }]
    applyTablePresentation(XLSX, sourceSheet, sourceRows)
    XLSX.utils.book_append_sheet(workbook, sourceSheet, '__sources')
    sheets.push({ name: '__sources', rows: sourceRows })
  }

  const methodologyRows: CellValue[][] = [
    ['key', 'value'],
    ['purpose', sanitizeSpreadsheetCell(plan.purpose)],
    ['data_mode', plan.dataMode],
    ['generated_at', new Date().toISOString()]
  ]
  const methodologySheet = XLSX.utils.aoa_to_sheet(methodologyRows)
  methodologySheet['!cols'] = [{ wch: 18 }, { wch: 72 }]
  applyTablePresentation(XLSX, methodologySheet, methodologyRows)
  XLSX.utils.book_append_sheet(workbook, methodologySheet, '__metadata')
  sheets.push({ name: '__metadata', rows: methodologyRows })
  workbook.Workbook = workbook.Workbook || {}
  workbook.Workbook.Sheets = workbook.SheetNames.map(name => ({
    name,
    Hidden: name.startsWith('__') ? 1 : 0
  }))

  let bytes = toUint8Array(XLSX.write(workbook, {
    type: 'array',
    bookType: 'xlsx',
    compression: true,
    cellStyles: true
  }))
  if (formulas.length) bytes = finalizeGeneratedWorkbookFormulas(bytes)
  const summary = summarizeWorkbookSheets(sheets.filter(sheet => !sheet.name.startsWith('__')))
  if (formulas.length) summary.calculation = { status: 'pending', formulaCells: formulas.length }
  return {
    bytes,
    preview: mainRows.slice(0, 11),
    summary
  }
}

function applyTablePresentation(
  XLSX: SheetJs,
  sheet: import('xlsx').WorkSheet,
  rows: CellValue[][],
  columns: SpreadsheetColumn[] = []
) {
  const columnCount = Math.max(1, rows[0]?.length || columns.length)
  const lastColumn = XLSX.utils.encode_col(columnCount - 1)
  const lastRow = Math.max(1, rows.length)
  sheet['!autofilter'] = { ref: `A1:${lastColumn}${lastRow}` }
  sheet['!rows'] = rows.map((_, index) => ({ hpt: index === 0 ? 25 : 21 }))
  if (columns.length) {
    sheet['!cols'] = columns.map(column => ({
      wch: column.type === 'url'
        ? 42
        : Math.min(32, Math.max(12, column.label.length + 5))
    }))
  } else if (!sheet['!cols']) {
    sheet['!cols'] = Array.from({ length: columnCount }, (_, index) => ({
      wch: Math.min(36, Math.max(12, String(rows[0]?.[index] || '').length + 4))
    }))
  }

  for (let rowIndex = 0; rowIndex < rows.length; rowIndex += 1) {
    for (let columnIndex = 0; columnIndex < columnCount; columnIndex += 1) {
      const address = XLSX.utils.encode_cell({ r: rowIndex, c: columnIndex })
      const cell = sheet[address]
      if (!cell) continue

      if (rowIndex === 0) {
        cell.s = {
          fill: { patternType: 'solid', fgColor: { rgb: '132A3A' } },
          font: { bold: true, color: { rgb: 'FFFFFF' } },
          alignment: { vertical: 'center', horizontal: 'left' },
          border: { bottom: { style: 'thin', color: { rgb: '169B68' } } }
        }
        continue
      }

      cell.s = {
        fill: rowIndex % 2 === 0
          ? { patternType: 'solid', fgColor: { rgb: 'F1F8F4' } }
          : { patternType: 'solid', fgColor: { rgb: 'FFFFFF' } },
        font: { color: { rgb: '30483E' } },
        alignment: { vertical: 'center', horizontal: isNumericColumn(columns[columnIndex]) ? 'right' : 'left' },
        border: { bottom: { style: 'hair', color: { rgb: 'DDE7E2' } } }
      }

      const column = columns[columnIndex]
      if (column?.type === 'percent') cell.z = '0.0%'
      if (column?.type === 'currency') cell.z = '#,##0.00'
      if (column?.type === 'number') cell.z = '#,##0.00'
      if (column?.type === 'date') cell.z = 'yyyy-mm-dd'
      if (column?.type === 'url' && typeof cell.v === 'string' && /^https?:\/\//i.test(cell.v)) {
        cell.l = { Target: cell.v }
        cell.s.font = { color: { rgb: '0F8458' }, underline: true }
      }
    }
  }
}

function isNumericColumn(column: SpreadsheetColumn | undefined) {
  return Boolean(column && ['number', 'currency', 'percent'].includes(column.type))
}

export function sanitizeSpreadsheetCell(value: unknown): CellValue {
  if (value === null || value === undefined) return null
  if (typeof value === 'number' || typeof value === 'boolean') return value
  const text = String(value).trim()
  return /^[=+\-@]/.test(text) ? `'${text}` : text
}

async function resolveSpreadsheetRecords(
  env: AgentWorkerEnv,
  payload: SpreadsheetWorkflowPayload,
  research: WebSearchResult | null,
  maxRows: number
): Promise<SpreadsheetRecord[]> {
  if (payload.plan.dataMode === 'blank_template') return []
  if (payload.plan.dataMode === 'workbook_analysis') {
    return recordsFromAnalysis(payload.plan.columns, payload.analysis).slice(0, maxRows)
  }
  if (payload.plan.dataMode === 'web_research' && !research) {
    throw new Error('A web-research spreadsheet requires a completed search result')
  }

  const schemaProperties = Object.fromEntries(payload.plan.columns.map(column => [column.key, jsonSchemaForColumn(column)]))
  const raw = await runResponsesModel(env, {
    instructions: [
      'Transform the supplied evidence into spreadsheet records.',
      'Use only facts in the prompt. Never guess or fabricate missing values.',
      'Return one record per distinct entity or observation, up to the requested limit.',
      'Use null for optional values that are not supported by evidence.',
      'Keep dates in ISO YYYY-MM-DD form when possible, percentages as decimal numbers, and currencies as numeric values without symbols.'
    ].join(' '),
    input: JSON.stringify({
      userRequest: payload.prompt,
      spreadsheetPurpose: payload.plan.purpose,
      columns: payload.plan.columns,
      maxRows,
      research: research ? { summary: research.summary, sources: research.sources } : null
    }),
    text: {
      format: {
        type: 'json_schema',
        name: 'spreadsheet_records',
        strict: true,
        schema: {
          type: 'object',
          additionalProperties: false,
          properties: {
            records: {
              type: 'array',
              maxItems: maxRows,
              items: {
                type: 'object',
                additionalProperties: false,
                properties: schemaProperties,
                required: payload.plan.columns.map(column => column.key)
              }
            }
          },
          required: ['records']
        }
      }
    },
    max_output_tokens: 8000,
    store: false
  })
  const parsed = parseJsonObject(extractResponseText(raw))
  const records = Array.isArray(parsed.records) ? parsed.records : []
  return records
    .filter((record): record is Record<string, unknown> => Boolean(record) && typeof record === 'object' && !Array.isArray(record))
    .slice(0, maxRows)
    .map(record => Object.fromEntries(payload.plan.columns.map(column => [column.key, coerceCell(record[column.key], column)])))
}

export function recordsFromAnalysis(columns: SpreadsheetColumn[], analysis: WorkbookAnalysis | null): SpreadsheetRecord[] {
  if (!analysis) throw new Error('A workbook-analysis spreadsheet requires a completed analysis')
  if (analysis.table) {
    const indexes = columns.map(column => {
      const index = analysis.table!.columns.indexOf(column.sourceField || '')
      if (index < 0) throw new Error(`Workbook analysis source field was not found: ${column.sourceField || column.key}`)
      return index
    })
    return analysis.table.rows.map(row => Object.fromEntries(columns.map((column, index) => [column.key, row[indexes[index]!] ?? null])))
  }
  return [Object.fromEntries(columns.map(column => {
    const metric = analysis.metrics.find(item => item.label === column.sourceField || item.field === column.sourceField)
    if (!metric) throw new Error(`Workbook analysis metric was not found: ${column.sourceField || column.key}`)
    return [column.key, metric.value]
  }))]
}

function jsonSchemaForColumn(column: SpreadsheetColumn) {
  const valueType = ['number', 'currency', 'percent'].includes(column.type)
    ? 'number'
    : column.type === 'boolean'
      ? 'boolean'
      : 'string'
  return column.required ? { type: valueType } : { type: [valueType, 'null'] }
}

function coerceCell(value: unknown, column: SpreadsheetColumn): CellValue {
  if (value === null || value === undefined) return null
  if (['number', 'currency', 'percent'].includes(column.type)) {
    const number = typeof value === 'number' ? value : Number(String(value).replace(/[,$%\s]/g, ''))
    return Number.isFinite(number) ? number : null
  }
  if (column.type === 'boolean') return typeof value === 'boolean' ? value : String(value).toLowerCase() === 'true'
  return sanitizeSpreadsheetCell(value)
}

async function readResearchSnapshot(bucket: R2Bucket, key: string | null) {
  if (!key) return null
  const object = await bucket.get(key)
  if (!object) throw new Error('Research snapshot was not found in R2')
  return JSON.parse(await object.text()) as WebSearchResult
}

async function readRecordsSnapshot(bucket: R2Bucket, key: string) {
  const object = await bucket.get(key)
  if (!object) return null
  const parsed = z.object({ schemaVersion: z.union([z.literal(1), z.literal(2)]).optional(), records: z.array(spreadsheetRecordSchema).max(100) }).parse(JSON.parse(await object.text()))
  return parsed.records
}

function parseJsonObject(text: string): Record<string, unknown> {
  const normalized = text.trim().replace(/^```(?:json)?\s*/i, '').replace(/\s*```$/, '')
  try {
    return JSON.parse(normalized) as Record<string, unknown>
  } catch {
    const start = normalized.indexOf('{')
    const end = normalized.lastIndexOf('}')
    if (start >= 0 && end > start) return JSON.parse(normalized.slice(start, end + 1)) as Record<string, unknown>
    throw new Error('The spreadsheet row generator returned invalid JSON')
  }
}

async function importSheetJs(): Promise<SheetJs> {
  const sheetJsWithStyles = await import('xlsx-js-style')
  return sheetJsWithStyles.default as unknown as SheetJs
}

function safeFileStem(value: string) {
  return value.replace(/[^\p{L}\p{N}._-]+/gu, '-').replace(/^-+|-+$/g, '').slice(0, 120) || 'excelgen-spreadsheet'
}

function toUint8Array(value: unknown) {
  if (value instanceof Uint8Array) return value
  if (value instanceof ArrayBuffer) return new Uint8Array(value)
  throw new Error('SheetJS returned an unsupported workbook payload')
}

function clampInteger(value: string | undefined, minimum: number, maximum: number, fallback: number) {
  const parsed = Number.parseInt(value || '', 10)
  return Number.isFinite(parsed) ? Math.min(maximum, Math.max(minimum, parsed)) : fallback
}
