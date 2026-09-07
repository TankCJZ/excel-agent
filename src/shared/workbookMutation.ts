import { z } from 'zod'

export const MUTATION_LIMITS = {
  sourceBytes: 10 * 1024 * 1024,
  expandedBytes: 32 * 1024 * 1024,
  partBytes: 12 * 1024 * 1024,
  zipEntries: 1500,
  sheets: 32,
  physicalCells: 150_000,
  // XML nodes/attributes cost much more memory than their ZIP bytes. This
  // conservative budget is checked before materializing any worksheet DOM.
  xmlComplexity: 120_000,
  parsedXmlBytes: 8 * 1024 * 1024,
  changedCells: 10_000,
  operations: 40,
  planBytes: 256 * 1024,
  formulas: 10_000,
  dependencyChecks: 1_000_000,
  formulaLength: 2048
} as const

const sheetName = z.string().min(1).max(31).regex(/^[^\\/?*\[\]:\x00-\x1f]+$/)
const address = z.string().regex(/^[A-Z]{1,3}[1-9]\d{0,6}$/)
const range = z.string().regex(/^[A-Z]{1,3}[1-9]\d{0,6}(?::[A-Z]{1,3}[1-9]\d{0,6})?$/)
export const mutationValueSchema = z.union([
  z.string().max(32767), z.number().finite(), z.boolean(), z.null(),
  z.object({ type: z.literal('date'), value: z.string().regex(/^\d{4}-\d{2}-\d{2}$/) }).strict()
])
const matrix = z.array(z.array(mutationValueSchema).min(1).max(256)).max(MUTATION_LIMITS.changedCells)
export const mutationOperationSchema = z.discriminatedUnion('kind', [
  z.object({ kind: z.literal('set_values'), sheet: sheetName, start: address, values: matrix.min(1), overwrite: z.boolean() }).strict(),
  z.object({ kind: z.literal('append_column'), sheet: sheetName, title: z.string().trim().min(1).max(120), headerRow: z.number().int().min(1).max(1048576), values: z.array(mutationValueSchema).max(MUTATION_LIMITS.changedCells) }).strict(),
  z.object({ kind: z.literal('add_sheet'), sheet: sheetName, values: matrix }).strict(),
  z.object({ kind: z.literal('set_formula'), sheet: sheetName, range, formula: z.string().min(2).max(MUTATION_LIMITS.formulaLength), overwrite: z.boolean() }).strict(),
  z.object({ kind: z.literal('set_number_format'), sheet: sheetName, range, format: z.enum(['general', 'integer', 'decimal', 'percent', 'date', 'currency_usd']) }).strict()
])

function checkPlanBudget(input: { operations: z.infer<typeof mutationOperationSchema>[] }, context: z.RefinementCtx) {
  if (new TextEncoder().encode(JSON.stringify(input)).length > MUTATION_LIMITS.planBytes) {
    context.addIssue({ code: 'custom', message: 'The edit plan exceeds 256 KiB. Split it into smaller edits.' })
  }
  let cells = 0
  for (const operation of input.operations) {
    if (operation.kind === 'append_column') cells += operation.values.length + 1
    else if (operation.kind === 'set_values' || operation.kind === 'add_sheet') cells += operation.values.reduce((sum, row) => sum + row.length, 0)
    else {
      const [start, end = start] = operation.range.split(':')
      const coordinate = (address: string) => {
        const match = address?.match(/^([A-Z]+)(\d+)$/)
        if (!match) return null
        const [, letters, row] = match
        return { row: Number(row), col: [...letters!].reduce((n, letter) => n * 26 + letter.charCodeAt(0) - 64, 0) }
      }
      const a = coordinate(start!), b = coordinate(end!)
      if (!a || !b) continue
      // Engine validates bounds/direction too; never iterate an unbounded range.
      cells += (Math.abs(b.row - a.row) + 1) * (Math.abs(b.col - a.col) + 1)
    }
  }
  if (cells > MUTATION_LIMITS.changedCells) context.addIssue({ code: 'custom', message: 'An edit may target at most 10,000 cells across all operations.' })
}

const workbookMutationInputObject = z.object({
  title: z.string().trim().min(1).max(120),
  operations: z.array(mutationOperationSchema).min(1).max(MUTATION_LIMITS.operations)
}).strict()
export const workbookMutationInputSchema = workbookMutationInputObject.superRefine(checkPlanBudget)

// Only the server binds identity/version; the model cannot choose an R2 key.
export const workbookMutationPlanSchema = workbookMutationInputObject.extend({
  kind: z.literal('workbook_mutation'),
  schemaVersion: z.literal(1),
  sourceWorkbookId: z.string().min(1),
  sourceEtag: z.string().min(1)
}).strict().superRefine(checkPlanBudget)

export type MutationValue = z.infer<typeof mutationValueSchema>
export type MutationOperation = z.infer<typeof mutationOperationSchema>
export type WorkbookMutationInput = z.infer<typeof workbookMutationInputSchema>
export type WorkbookMutationPlan = z.infer<typeof workbookMutationPlanSchema>
export type MutationPreviewCell = {
  address: string
  value: string | number | boolean | null
  formula?: string
  pendingCalculation?: boolean
}
export type MutationPreviewSheet = { name: string; rows: Array<{ row: number; cells: MutationPreviewCell[] }> }
export type MutationChanges = {
  changedCells: number
  formulaCells: number
  addedSheets: string[]
  addedColumns: Array<{ sheet: string; column: string; title: string }>
  pendingCalculation: boolean
  convertedFrom: 'xls' | 'csv' | null
  sheets: MutationPreviewSheet[]
}
