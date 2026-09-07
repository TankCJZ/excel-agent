import { createTool } from '@mastra/core/tools'
import { z } from 'zod'
import { workbookMutationInputSchema, type WorkbookMutationPlan, type MutationChanges } from '#shared/agent/workbookMutation'
import type { AgentWorkerEnv } from '#agent/types'
import type { ExcelAgentRuntimeInput } from '#agent/mastra/runtimeTypes'
import { applyWorkbookMutation, inspectWorkbookForEditing } from '#agent/workbook/mutation/engine'
import { prepareEditableBytes } from '#agent/workbook/mutation/source'
import { requireMutation } from '#agent/workbook/mutation/errors'

export function createWorkbookEditingTools(env: AgentWorkerEnv, input: ExcelAgentRuntimeInput, assertOutputAvailable: () => void) {
  let source: { bytes: Uint8Array; etag: string; convertedFrom: 'xls' | 'csv' | null } | null = null
  let plan: WorkbookMutationPlan | null = null
  let changes: MutationChanges | null = null
  let inspected = false
  const load = async () => {
    if (source) return source
    requireMutation(input.fileKey && input.workbookId, 'NO_WORKBOOK', 'Attach a workbook before editing.')
    const object = await env.WORKBOOKS.get(input.fileKey)
    requireMutation(object, 'SOURCE_MISSING', 'The original workbook is no longer available. Upload it again to edit it.')
    requireMutation(!input.sourceEtag || object.etag === input.sourceEtag, 'SOURCE_CHANGED', 'The source workbook changed. Reload it before editing.')
    const prepared = prepareEditableBytes(new Uint8Array(await object.arrayBuffer()), input.fileName || 'workbook.xlsx')
    source = { ...prepared, etag: object.etag }
    return source
  }
  return {
    getPlan: () => plan,
    getChanges: () => changes,
    tools: {
      inspectWorkbookForEditing: createTool({
        id: 'inspect-workbook-for-editing',
        description: 'Inspect the currently attached workbook for editing. Returns exact physical A1 cell addresses, preview values (long text truncated with …), existing formulas, merged ranges and last used columns. Always call this before editWorkbook on each turn; normalized analysis row numbers are NOT physical Excel row numbers. Does not modify any file. CSV edits create a separate XLSX with original CSV fields preserved as text; explain this to the user. XLS must first be saved as XLSX in Excel. Never overwrite text by copying a truncated preview.',
        inputSchema: z.object({}),
        execute: async () => {
          const current = await load()
          const context = inspectWorkbookForEditing(current.bytes)
          inspected = true
          return { ...context, convertedFrom: current.convertedFrom, formulaResults: 'Formula values are not recalculated here. Do not infer results from missing values.' }
        }
      }),
      editWorkbook: createTool({
        id: 'edit-workbook',
        description: 'Validate and queue edits to a NEW COPY of the attached workbook. Requires inspectWorkbookForEditing first in this turn. Use exact physical addresses. Supports set_values (rectangular literal data), append_column (rightmost, explicit physical headerRow), add_sheet, set_formula (one formula at range top-left, relative/absolute references filled by server), set_number_format. Never use for mere advice, a plan, row deletion, mid-sheet column insertion or renaming sheets. Set overwrite=true only for replacement explicitly requested by the user; otherwise ask for confirmation. One output-producing tool per turn. Formulas are real Excel formulas, calculated on opening in Excel, NOT evaluated by the tool. An accepted result means queued, not completed.',
        inputSchema: workbookMutationInputSchema,
        execute: async (edit) => {
          requireMutation(inspected, 'INSPECTION_REQUIRED', 'Inspect exact physical cell addresses with inspectWorkbookForEditing before editing.')
          requireMutation(!plan, 'OUTPUT_ALREADY_PLANNED', 'An edit has already been accepted for this turn.')
          assertOutputAvailable()
          const current = await load()
          const validated = applyWorkbookMutation(current.bytes, edit)
          plan = { ...edit, kind: 'workbook_mutation', schemaVersion: 1, sourceWorkbookId: input.workbookId!, sourceEtag: current.etag }
          changes = { ...validated.changes, convertedFrom: current.convertedFrom }
          return { accepted: true, status: 'validated_not_yet_applied', title: edit.title, changedCells: changes.changedCells, formulaCells: changes.formulaCells, addedSheets: changes.addedSheets, addedColumns: changes.addedColumns, pendingCalculation: changes.pendingCalculation, convertedFrom: changes.convertedFrom }
        }
      })
    }
  }
}
