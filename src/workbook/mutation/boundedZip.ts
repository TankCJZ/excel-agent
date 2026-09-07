import { Unzip, UnzipInflate, unzipSync } from 'fflate'
import { MUTATION_LIMITS } from '#shared/agent/workbookMutation'
import { requireMutation } from '#agent/workbook/mutation/errors'

// Preflight the central directory WITHOUT inflating, then validate actual
// output incrementally. Declared sizes alone are not a decompression limit.
export function readBoundedZip(bytes: Uint8Array, keep: (name: string) => boolean = () => true) {
  const declared = new Map<string, number>()
  let declaredTotal = 0
  unzipSync(bytes, { filter(entry) {
    requireMutation(!declared.has(entry.name) && !entry.name.includes('\\') && !entry.name.split('/').includes('..') && !entry.name.startsWith('/'), 'INVALID_WORKBOOK', 'Workbook ZIP contains duplicate or unsafe paths.')
    declared.set(entry.name, entry.originalSize)
    declaredTotal += entry.originalSize
    requireMutation(declared.size <= MUTATION_LIMITS.zipEntries && declaredTotal <= MUTATION_LIMITS.expandedBytes && entry.originalSize <= MUTATION_LIMITS.partBytes, 'FILE_LIMIT', 'Expanded workbook exceeds safe processing limits.')
    return false
  } })
  const parts: Record<string, Uint8Array> = Object.create(null)
  const seen = new Set<string>()
  let expanded = 0, completed = 0
  const archive = new Unzip(file => {
    requireMutation(declared.has(file.name) && !seen.has(file.name), 'INVALID_WORKBOOK', 'Workbook ZIP directory does not match its entries.')
    seen.add(file.name)
    let size = 0
    const retain = keep(file.name)
    const chunks: Uint8Array[] = []
    file.ondata = (error, chunk, final) => {
      if (error) throw error
      size += chunk.length
      expanded += chunk.length
      requireMutation(size <= MUTATION_LIMITS.partBytes && expanded <= MUTATION_LIMITS.expandedBytes, 'FILE_LIMIT', 'Actual expanded workbook exceeds safe processing limits.')
      requireMutation(size <= declared.get(file.name)!, 'INVALID_WORKBOOK', 'Workbook ZIP entry is larger than its declared size.')
      if (retain) chunks.push(chunk)
      if (!final) return
      requireMutation(size === declared.get(file.name), 'INVALID_WORKBOOK', 'Workbook ZIP entry is truncated.')
      completed++
      if (retain) {
        const output = new Uint8Array(size)
        let offset = 0
        for (const chunk of chunks) { output.set(chunk, offset); offset += chunk.length }
        chunks.length = 0
        parts[file.name] = output
      }
    }
    file.start()
  })
  archive.register(UnzipInflate)
  // Bound transient inflate allocation as well as retained output. A small
  // compressed chunk prevents a forged tiny entry from expanding in one push.
  for (let offset = 0; offset < bytes.length; offset += 1024) archive.push(bytes.subarray(offset, offset + 1024), offset + 1024 >= bytes.length)
  requireMutation(completed === declared.size && seen.size === declared.size, 'INVALID_WORKBOOK', 'Workbook ZIP entries are incomplete.')
  return parts
}
