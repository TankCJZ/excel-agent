import { readBoundedZip } from '#agent/workbook/mutation/boundedZip'
import { child, descendants, parseXml } from '#agent/workbook/mutation/package'
import { requireMutation } from '#agent/workbook/mutation/errors'
import type { WorkbookSummary } from '#agent/contracts'

// SheetJS may synthesize zero for a formula without <v>. Inspect the original
// XML before reading values, so that synthesized cache can never become data.
export function hasUncalculatedXlsxFormulas(bytes: Uint8Array): boolean {
  if (bytes[0] !== 0x50 || bytes[1] !== 0x4b) return false
  const sheets = readBoundedZip(bytes, name => /^xl\/worksheets\/[^/]+\.xml$/.test(name))
  for (const bytes of Object.values(sheets)) {
    for (const cell of descendants(parseXml(bytes), 'c')) {
      if (!child(cell, 'f')) continue
      const value = child(cell, 'v')
      if (!value || (!(value.textContent || '').trim() && cell.getAttribute('t') !== 'str')) return true
    }
  }
  return false
}

export function assertWorkbookCalculated(summary: WorkbookSummary) {
  requireMutation(summary.calculation?.status !== 'pending', 'WORKBOOK_CALCULATION_PENDING',
    'This workbook contains formulas that have not been recalculated. Download it, recalculate and save it in Excel, then upload the saved file before analyzing results or creating charts. You can still edit cells and formulas.')
}
