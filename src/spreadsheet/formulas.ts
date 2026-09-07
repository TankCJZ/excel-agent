import type { SpreadsheetPlan, SpreadsheetRecord } from '#agent/contracts'
import { assertAcyclicFormulas, parseSafeFormula } from '#agent/workbook/mutation/formulas'
import { cellAddress } from '#agent/workbook/mutation/coordinates'

export function spreadsheetFormulas(plan: SpreadsheetPlan, records: SpreadsheetRecord[]) {
  const formulas = records.flatMap((record, row) => plan.columns.flatMap((column, col) => {
    const value = record[column.key]
    if (!value || typeof value !== 'object') return []
    const formula = parseSafeFormula(value.formula, plan.sheetName, [plan.sheetName]).formula
    return [{ sheet: plan.sheetName, address: cellAddress({ row: row + 2, col: col + 1 }), formula }]
  }))
  assertAcyclicFormulas(formulas, [plan.sheetName])
  return formulas
}
