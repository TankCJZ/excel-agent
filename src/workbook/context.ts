import { z } from 'zod'
import {
  cellValueSchema,
  workbookSummarySchema,
  type WorkbookSummary
} from '#agent/contracts'
import type { WorkbookDataSheet } from '#agent/workbook/analysis'

export const WORKBOOK_CONTEXT_VERSION = 2

export type WorkbookContextEnvelope = {
  version: typeof WORKBOOK_CONTEXT_VERSION
  sourceEtag: string
  summary: WorkbookSummary
  sheets: WorkbookDataSheet[]
}

export type StandardWorkbookContext = Pick<WorkbookContextEnvelope, 'summary' | 'sheets'>

export type WorkbookContextLoadResult = {
  context: StandardWorkbookContext
  cacheStatus: 'hit' | 'repaired' | 'fallback'
  contextR2Key: string
  contextVersion: number
  sourceEtag: string
  parsedAt: string
}

const workbookContextEnvelopeSchema = z.object({
  version: z.union([z.literal(1), z.literal(WORKBOOK_CONTEXT_VERSION)]),
  sourceEtag: z.string().min(1),
  summary: workbookSummarySchema,
  sheets: z.array(z.object({
    name: z.string().min(1),
    rows: z.array(z.array(cellValueSchema))
  })).min(1)
})

export function workbookContextKey(sessionId: string, workbookId: string) {
  return `agent-demo/contexts/${sessionId}/${workbookId}/context-v${WORKBOOK_CONTEXT_VERSION}.json.gz`
}

export async function encodeWorkbookContext(envelope: WorkbookContextEnvelope) {
  const json = JSON.stringify(workbookContextEnvelopeSchema.parse(envelope))
  const stream = new Blob([json]).stream().pipeThrough(new CompressionStream('gzip'))
  return new Uint8Array(await new Response(stream).arrayBuffer())
}

export async function decodeWorkbookContext(bytes: ArrayBuffer | Uint8Array) {
  const body = bytes instanceof Uint8Array ? Uint8Array.from(bytes).buffer : bytes
  const stream = new Blob([body]).stream().pipeThrough(new DecompressionStream('gzip'))
  const json = await new Response(stream).text()
  return workbookContextEnvelopeSchema.parse(JSON.parse(json))
}

export async function writeWorkbookContext(
  bucket: R2Bucket,
  key: string,
  context: { summary: WorkbookSummary, sheets: WorkbookDataSheet[] },
  sourceEtag: string
) {
  const bytes = await encodeWorkbookContext({
    version: WORKBOOK_CONTEXT_VERSION,
    sourceEtag,
    summary: context.summary,
    sheets: context.sheets
  })

  await bucket.put(key, bytes, {
    httpMetadata: {
      contentType: 'application/json',
      contentEncoding: 'gzip'
    },
    customMetadata: {
      contextVersion: String(WORKBOOK_CONTEXT_VERSION),
      sourceEtag
    }
  })

  return { key, sizeBytes: bytes.byteLength }
}

export async function readWorkbookContextObject(
  bucket: R2Bucket,
  key: string,
  expectedSourceEtag: string
) {
  const object = await bucket.get(key)
  if (!object) {
    throw new Error('Workbook context object was not found in R2')
  }

  const envelope = await decodeWorkbookContext(await object.arrayBuffer())
  if (envelope.version !== WORKBOOK_CONTEXT_VERSION) {
    throw new Error('Workbook context requires calculation-aware rebuilding')
  }
  if (envelope.sourceEtag !== expectedSourceEtag) {
    throw new Error('Workbook context does not match the source workbook')
  }

  return {
    summary: envelope.summary,
    sheets: envelope.sheets
  }
}

export async function loadWorkbookContextCache(
  bucket: R2Bucket,
  input: {
    workbookId: string
    sessionId: string
    fileKey: string
    contextR2Key: string | null
    contextVersion: number | null
    sourceEtag: string | null
  },
  parseSource: (bytes: ArrayBuffer) => Promise<StandardWorkbookContext>
): Promise<WorkbookContextLoadResult> {
  if (
    input.contextR2Key
    && input.contextVersion === WORKBOOK_CONTEXT_VERSION
    && input.sourceEtag
  ) {
    try {
      const context = await readWorkbookContextObject(
        bucket,
        input.contextR2Key,
        input.sourceEtag
      )
      return {
        context,
        cacheStatus: 'hit',
        contextR2Key: input.contextR2Key,
        contextVersion: WORKBOOK_CONTEXT_VERSION,
        sourceEtag: input.sourceEtag,
        parsedAt: new Date().toISOString()
      }
    } catch (error) {
      console.warn('Workbook context cache is unavailable; rebuilding from source', error)
    }
  }

  const source = await bucket.get(input.fileKey)
  if (!source) {
    throw new Error('Workbook object was not found in R2')
  }

  const context = await parseSource(await source.arrayBuffer())
  const contextR2Key = workbookContextKey(input.sessionId, input.workbookId)
  let cacheStatus: WorkbookContextLoadResult['cacheStatus'] = 'repaired'
  try {
    await writeWorkbookContext(bucket, contextR2Key, context, source.etag)
  } catch (error) {
    // A cache write must not turn a readable workbook into a failed user request.
    cacheStatus = 'fallback'
    console.error('Workbook context cache could not be repaired', error)
  }

  return {
    context,
    cacheStatus,
    contextR2Key,
    contextVersion: WORKBOOK_CONTEXT_VERSION,
    sourceEtag: source.etag,
    parsedAt: new Date().toISOString()
  }
}
