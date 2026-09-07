import { z } from 'zod'
import { cellValueSchema, spreadsheetColumnSchema, spreadsheetExecutionPlanSchema, webSearchSourceSchema, workbookAnalysisSchema, workbookPlanSchema, workbookSummarySchema } from '#agent/contracts'
import { workbookMutationPlanSchema } from '#shared/agent/workbookMutation'
import type { AgentTaskPlan, AgentTaskResult, WorkbookMutationPayload, WorkbookMutationResult } from '#agent/types'

const id = z.string().min(1).max(200)
const key = z.string().min(1).max(512)
const count = z.number().int().nonnegative()
const preview = z.array(z.array(cellValueSchema))
const changesSchema = z.object({
  changedCells: count, formulaCells: count, addedSheets: z.array(z.string()),
  addedColumns: z.array(z.object({ sheet: z.string(), column: z.string(), title: z.string() })),
  pendingCalculation: z.boolean(), convertedFrom: z.enum(['xls', 'csv']).nullable(),
  sheets: z.array(z.object({ name: z.string(), rows: z.array(z.object({ row: count, cells: z.array(z.object({ address: z.string(), value: cellValueSchema, formula: z.string().optional(), pendingCalculation: z.boolean().optional() })) })) }))
})
const resultBase = z.object({
  taskId: id, outputR2Key: key, outputName: z.string(), summary: workbookSummarySchema,
  preview, previewSheetName: z.string().optional(), previewRowCount: count.optional(), sizeBytes: count
})
export const mutationResultSchema = resultBase.extend({
  kind: z.literal('workbook_mutation'), schemaVersion: z.literal(1), analysis: z.null(),
  previewSheetName: z.string(), previewRowCount: count, changes: changesSchema,
  sourceWorkbookId: id, workbookId: id, sourceEtag: id, contextR2Key: key,
  contextVersion: z.union([z.literal(1), z.literal(2)])
})
const mutationExecutionSchema = z.object({ action: z.literal('edit_workbook'), mutation: workbookMutationPlanSchema, steps: z.array(z.object({ id: z.string(), operation: z.string() })) })
const payloadSchema = z.object({
  kind: z.literal('workbook_mutation'), taskId: id, userId: id, sessionId: id, threadId: id, turnId: id,
  inputR2Key: key, inputName: z.string().min(1).max(200), plan: workbookMutationPlanSchema,
  resultWorkbookId: id, sourceRevision: z.number().int().min(-1), creditReservationId: id.nullable(), creditTurnId: id, creditToolNames: z.array(z.string().max(100)).max(40)
}).strict()
const legacyResults = z.discriminatedUnion('kind', [
  resultBase.extend({ kind: z.literal('workbook_export'), analysis: workbookAnalysisSchema }),
  resultBase.extend({ kind: z.literal('spreadsheet_generation'), workbookId: id.optional(), analysis: z.null(), previewSheetName: z.string(), previewRowCount: count, columns: z.array(spreadsheetColumnSchema), sources: z.array(webSearchSourceSchema) })
])
export function parseMutationPayload(value: unknown): WorkbookMutationPayload { return payloadSchema.parse(value) }
export function parseMutationResult(value: unknown): WorkbookMutationResult { return mutationResultSchema.parse(value) }

function object(value: unknown): Record<string, unknown> | null {
  return value && typeof value === 'object' && !Array.isArray(value) ? value as Record<string, unknown> : null
}
export function decodeStoredJson(value: string | null): unknown {
  if (!value) return null
  try { return JSON.parse(value) } catch { return null }
}

// Old records without a kind/version are interpreted only via their stored
// task discriminator. Unknown future versions are never executed as old ones.
export function readTaskResult(value: unknown, taskType: string): AgentTaskResult | null {
  const raw = object(value)
  if (!raw) return null
  if (raw.kind === 'workbook_mutation' || taskType === 'workbook_mutation') {
    const result = mutationResultSchema.safeParse(raw)
    return result.success ? result.data : null
  }
  if (raw.schemaVersion !== undefined && raw.schemaVersion !== 1) return null
  const kind = raw.kind || taskType
  const summary = object(raw.summary)
  const normalized = {
    ...raw, kind,
    ...(summary ? { summary: { ...summary, sheets: summary.sheets || [] } } : {}),
    ...(kind === 'spreadsheet_generation' ? { columns: raw.columns || [], sources: raw.sources || [], analysis: null, previewSheetName: raw.previewSheetName || summary?.primarySheet || '', previewRowCount: raw.previewRowCount ?? summary?.rowCount ?? 0 } : {})
  }
  const result = legacyResults.safeParse(normalized)
  return result.success ? result.data : null
}
export function readTaskPlan(value: unknown): AgentTaskPlan | null {
  const raw = object(value)
  if (!raw || (raw.schemaVersion !== undefined && raw.schemaVersion !== 1)) return null
  const schema = raw.action === 'edit_workbook' ? mutationExecutionSchema : raw.action === 'create_spreadsheet' ? spreadsheetExecutionPlanSchema : workbookPlanSchema
  const result = schema.safeParse(raw)
  return result.success ? result.data : null
}
