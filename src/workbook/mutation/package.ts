import { DOMParser, XMLSerializer, type Document, type Element, type Node } from '@xmldom/xmldom'
import { strFromU8, strToU8, zipSync } from 'fflate'
import { readBoundedZip } from '#agent/workbook/mutation/boundedZip'
import { MUTATION_LIMITS } from '#shared/agent/workbookMutation'
import { requireMutation } from '#agent/workbook/mutation/errors'

export const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
export const REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'
export const OFFICE_REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
export const CONTENT_NS = 'http://schemas.openxmlformats.org/package/2006/content-types'
export type XmlDocument = Document
export type XmlElement = Element

export function children(parent: Node, localName?: string): Element[] {
  return Array.from(parent.childNodes).filter((node): node is Element => node.nodeType === 1 && (!localName || (node as Element).localName === localName))
}

export function descendants(parent: Element | Document, name: string, namespace = NS): Element[] {
  return Array.from(parent.getElementsByTagNameNS(namespace, name))
}

export function child(parent: Node, name: string): Element | undefined { return children(parent, name)[0] }

export function element(doc: Document, name: string, attributes: Record<string, string> = {}, text?: string, namespace = NS): Element {
  const result = doc.createElementNS(namespace, name)
  for (const [key, value] of Object.entries(attributes)) result.setAttribute(key, value)
  if (text !== undefined) result.appendChild(doc.createTextNode(text))
  return result
}

export function parseXml(bytes: Uint8Array): Document {
  requireMutation(bytes.length <= MUTATION_LIMITS.parsedXmlBytes, 'FILE_LIMIT', 'Workbook XML exceeds the safe parsing memory budget.')
  let complexity = 0
  for (const byte of bytes) if (byte === 60 || byte === 61) complexity++
  requireMutation(complexity <= MUTATION_LIMITS.xmlComplexity, 'FILE_LIMIT', 'Workbook XML is too complex to parse safely. Use a smaller workbook.')
  const text = strFromU8(bytes)
  requireMutation(!/<!DOCTYPE|<!ENTITY/i.test(text), 'UNSUPPORTED_WORKBOOK', 'XML entities and document types are not supported.')
  requireMutation(!/[\x00-\x08\x0b\x0c\x0e-\x1f]/.test(text), 'INVALID_WORKBOOK', 'Workbook XML contains invalid control characters.')
  return new DOMParser({ onError: (_level, message) => { throw new Error(`Invalid workbook XML: ${message}`) } }).parseFromString(text, 'application/xml')
}

export function serializeXml(doc: Document): Uint8Array { return strToU8(new XMLSerializer().serializeToString(doc)) }

export class XlsxPackage {
  private readonly parts: Record<string, Uint8Array>
  private readonly documents = new Map<string, Document>()
  private parsedXmlBytes = 0
  private xmlComplexity = 0
  readonly modified = new Set<string>()

  constructor(bytes: Uint8Array) {
    requireMutation(bytes.length <= MUTATION_LIMITS.sourceBytes, 'FILE_LIMIT', 'Workbook exceeds the 10 MB editing limit.')
    this.parts = readBoundedZip(bytes)
    requireMutation(this.parts['xl/workbook.xml'] && this.parts['[Content_Types].xml'], 'UNSUPPORTED_WORKBOOK', 'A standard XLSX workbook is required for editing.')
    for (const [name, data] of Object.entries(this.parts)) {
      requireMutation(data.length <= MUTATION_LIMITS.partBytes, 'FILE_LIMIT', 'Workbook part exceeds safe editing limits.')
      requireMutation(!/(?:vbaProject|externalLinks\/|activeX\/|embeddings\/|_xmlsignatures\/|connections\.xml|pivotTables\/|pivotCache\/|slicers\/|queryTables\/)/i.test(name), 'UNSUPPORTED_WORKBOOK', 'Workbooks with macros, external data, embedded objects, pivot tables or signatures cannot be safely edited yet.')
      if (name.endsWith('.rels')) {
        for (const rel of descendants(this.xml(name), 'Relationship', REL_NS)) {
          requireMutation(rel.getAttribute('TargetMode') !== 'External' || rel.getAttribute('Type') === `${OFFICE_REL_NS}/hyperlink`, 'UNSUPPORTED_WORKBOOK', 'External workbook relationships are not supported for editing.')
        }
      }
    }
    const workbook = this.xml('xl/workbook.xml')
    requireMutation(workbook.documentElement?.namespaceURI === NS && !descendants(workbook, 'workbookProtection').length, 'UNSUPPORTED_WORKBOOK', 'Protected or non-standard workbooks cannot be edited.')
    requireMutation(!descendants(workbook, 'definedName').some(node => /(?:\[|\b(?:WEBSERVICE|RTD|DDE|HYPERLINK)\s*\()/i.test(node.textContent || '')), 'UNSUPPORTED_WORKBOOK', 'Workbook contains an unsafe defined name.')
  }

  names() { return Object.keys(this.parts) }
  has(path: string) { return Object.hasOwn(this.parts, path) }
  xml(path: string): Document {
    const cached = this.documents.get(path)
    if (cached) return cached
    const bytes = this.parts[path]
    requireMutation(bytes, 'INVALID_WORKBOOK', `Required workbook part is missing: ${path}`)
    this.parsedXmlBytes += bytes.length
    requireMutation(this.parsedXmlBytes <= MUTATION_LIMITS.parsedXmlBytes, 'FILE_LIMIT', 'Workbook XML exceeds the safe editing memory budget. Use a smaller workbook.')
    // Counting delimiters is deliberately conservative (including text) and
    // allocates no token array proportional to untrusted XML structure.
    for (const byte of bytes) if (byte === 60 || byte === 61) this.xmlComplexity++
    requireMutation(this.xmlComplexity <= MUTATION_LIMITS.xmlComplexity, 'FILE_LIMIT', 'Workbook XML is too complex to edit safely. Use a smaller workbook.')
    const doc = parseXml(bytes)
    this.documents.set(path, doc)
    return doc
  }

  update(path: string, doc: Document) { this.documents.set(path, doc); this.modified.add(path) }
  remove(path: string) { delete this.parts[path]; this.documents.delete(path); this.modified.delete(path) }

  write(): Uint8Array {
    for (const path of this.modified) this.parts[path] = serializeXml(this.documents.get(path)!)
    requireMutation(Object.values(this.parts).reduce((sum, part) => sum + part.length, 0) <= MUTATION_LIMITS.expandedBytes, 'FILE_LIMIT', 'Output workbook exceeds the expanded size limit.')
    const bytes = zipSync(this.parts, { level: 6 })
    requireMutation(bytes.length <= MUTATION_LIMITS.sourceBytes, 'FILE_LIMIT', 'Output workbook exceeds the 10 MB limit.')
    return bytes
  }
}

export function relationshipPath(target: string, base = 'xl'): string {
  requireMutation(!/[\\?#%]/.test(target), 'INVALID_WORKBOOK', 'Unsupported workbook relationship path.')
  const parts = target.startsWith('/') ? [] : base.split('/')
  for (const part of target.split('/')) {
    if (!part || part === '.') continue
    if (part === '..') {
      requireMutation(parts.length, 'INVALID_WORKBOOK', 'Invalid workbook relationship path.')
      parts.pop()
    } else parts.push(part)
  }
  return parts.join('/')
}
