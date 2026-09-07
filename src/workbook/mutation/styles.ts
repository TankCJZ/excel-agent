import { child, children, element, type XlsxPackage } from '#agent/workbook/mutation/package'
import { requireMutation } from '#agent/workbook/mutation/errors'
import type { MutationValue } from '#shared/agent/workbookMutation'

const FORMATS = { general: 0, integer: 1, decimal: 2, percent: 10, date: 14, currency_usd: 164 } as const

export function styleWithFormat(pkg: XlsxPackage, baseStyle: string | null, format: keyof typeof FORMATS): string {
  const doc = pkg.xml('xl/styles.xml')
  const root = doc.documentElement!
  const cellXfs = child(root, 'cellXfs')
  requireMutation(cellXfs, 'UNSUPPORTED_WORKBOOK', 'Workbook has no editable cell styles.')
  const xfs = children(cellXfs, 'xf')
  const base = xfs[Number(baseStyle || 0)]
  requireMutation(base, 'INVALID_WORKBOOK', 'Cell references an invalid style.')
  let id: number = FORMATS[format]
  if (format === 'currency_usd') {
    let formats = child(root, 'numFmts')
    if (!formats) {
      formats = element(doc, 'numFmts', { count: '0' })
      root.insertBefore(formats, root.firstChild)
    }
    const list = children(formats, 'numFmt')
    const code = '"$"#,##0.00'
    const existing = list.find(item => item.getAttribute('formatCode') === code)
    id = existing ? Number(existing.getAttribute('numFmtId')) : Math.max(163, ...list.map(item => Number(item.getAttribute('numFmtId')))) + 1
    if (!existing) {
      formats.appendChild(element(doc, 'numFmt', { numFmtId: String(id), formatCode: code }))
      formats.setAttribute('count', String(list.length + 1))
    }
  }
  const candidate = base.cloneNode(true) as typeof base
  candidate.setAttribute('numFmtId', String(id))
  candidate.setAttribute('applyNumberFormat', '1')
  const existing = xfs.findIndex(xf => xf.toString() === candidate.toString())
  if (existing >= 0) return String(existing)
  requireMutation(xfs.length < 2000, 'FILE_LIMIT', 'Workbook has too many distinct cell styles.')
  cellXfs.appendChild(candidate)
  cellXfs.setAttribute('count', String(xfs.length + 1))
  pkg.update('xl/styles.xml', doc)
  return String(xfs.length)
}

export function dateSerial(value: Extract<MutationValue, { type: 'date' }>, date1904: boolean): number {
  const time = Date.parse(`${value.value}T00:00:00.000Z`)
  requireMutation(Number.isFinite(time) && new Date(time).toISOString().slice(0, 10) === value.value && value.value >= (date1904 ? '1904-01-01' : '1900-01-01'), 'INVALID_DATE', 'Use a valid calendar date in the workbook date system.')
  const epoch = Date.parse(date1904 ? '1904-01-01T00:00:00.000Z' : '1899-12-31T00:00:00.000Z')
  return (time - epoch) / 86400000 + (!date1904 && value.value >= '1900-03-01' ? 1 : 0)
}
