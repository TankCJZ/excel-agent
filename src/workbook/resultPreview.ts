import type { WorkbookSummary } from '#agent/contracts'
import { requireMutation } from '#agent/workbook/mutation/errors'

// Results cross Workflow step and D1 boundaries. Only the presentation copy is
// shortened; the workbook and R2 analysis context retain their exact values.
export function compactResultSummary(summary: WorkbookSummary): WorkbookSummary {
  const preview = (rows: WorkbookSummary['preview']) => rows.slice(0, 20).map(row => row.slice(0, 20).map(value =>
    typeof value === 'string' && value.length > 240 ? `${value.slice(0, 240)}…` : value))
  return { ...summary, preview: preview(summary.preview), sheets: summary.sheets.map(sheet => ({ ...sheet, preview: preview(sheet.preview) })) }
}

export function assertWorkflowResultBudget(result: unknown) {
  requireMutation(new TextEncoder().encode(JSON.stringify(result)).length <= 512 * 1024, 'FILE_LIMIT', 'The workbook result is too complex to preview safely. Use a smaller workbook. The original is unchanged.')
}
