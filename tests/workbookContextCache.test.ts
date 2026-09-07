import assert from 'node:assert/strict'
import { readFile } from 'node:fs/promises'
import test from 'node:test'
import { summarizeWorkbookSheets, type WorkbookDataSheet } from '../src/workbook/analysis.ts'
import {
  WORKBOOK_CONTEXT_VERSION,
  decodeWorkbookContext,
  encodeWorkbookContext,
  loadWorkbookContextCache,
  workbookContextKey,
  writeWorkbookContext
} from '../src/workbook/context.ts'

const sheets: WorkbookDataSheet[] = [{
  name: '销售数据',
  rows: [
    ['月份', '销售额', '成交'],
    ['1月', 1200, true],
    ['2月', 1800, false]
  ]
}]
const context = { sheets, summary: summarizeWorkbookSheets(sheets) }

test('标准 Context 可以 gzip 往返且保留工作簿结构', async () => {
  const encoded = await encodeWorkbookContext({
    version: WORKBOOK_CONTEXT_VERSION,
    sourceEtag: 'source-etag-1',
    ...context
  })
  const decoded = await decodeWorkbookContext(encoded)

  assert.ok(encoded.byteLength > 0)
  assert.equal(decoded.version, WORKBOOK_CONTEXT_VERSION)
  assert.equal(decoded.sourceEtag, 'source-etag-1')
  assert.deepEqual(decoded.summary, context.summary)
  assert.deepEqual(decoded.sheets, sheets)
})

test('Context 命中时只读缓存，不读取或解析原始 Excel', async () => {
  const bucket = createMemoryBucket()
  const key = workbookContextKey('session-1', 'workbook-1')
  await writeWorkbookContext(bucket as never, key, context, 'source-etag-1')
  bucket.getCalls.length = 0
  let parseCount = 0

  const loaded = await loadWorkbookContextCache(bucket as never, {
    workbookId: 'workbook-1',
    sessionId: 'session-1',
    fileKey: 'agent-demo/uploads/session-1/source.xlsx',
    contextR2Key: key,
    contextVersion: WORKBOOK_CONTEXT_VERSION,
    sourceEtag: 'source-etag-1'
  }, async () => {
    parseCount += 1
    return context
  })

  assert.equal(loaded.cacheStatus, 'hit')
  assert.equal(parseCount, 0)
  assert.deepEqual(bucket.getCalls, [key])
  assert.deepEqual(loaded.context, context)
})

test('旧记录首次读取时解析一次并自修复，下一轮直接命中 Context', async () => {
  const fileKey = 'agent-demo/uploads/session-2/source.xlsx'
  const sourceBytes = new TextEncoder().encode('fake workbook bytes')
  const bucket = createMemoryBucket({
    [fileKey]: { bytes: sourceBytes, etag: 'source-etag-2' }
  })
  let parseCount = 0
  const parseSource = async (bytes: ArrayBuffer) => {
    parseCount += 1
    assert.deepEqual(new Uint8Array(bytes), sourceBytes)
    return context
  }

  const repaired = await loadWorkbookContextCache(bucket as never, {
    workbookId: 'workbook-2',
    sessionId: 'session-2',
    fileKey,
    contextR2Key: null,
    contextVersion: null,
    sourceEtag: null
  }, parseSource)
  assert.equal(repaired.cacheStatus, 'repaired')
  assert.equal(parseCount, 1)
  assert.ok(bucket.objects.has(repaired.contextR2Key))

  bucket.getCalls.length = 0
  const hit = await loadWorkbookContextCache(bucket as never, {
    workbookId: 'workbook-2',
    sessionId: 'session-2',
    fileKey,
    contextR2Key: repaired.contextR2Key,
    contextVersion: repaired.contextVersion,
    sourceEtag: repaired.sourceEtag
  }, parseSource)

  assert.equal(hit.cacheStatus, 'hit')
  assert.equal(parseCount, 1)
  assert.deepEqual(bucket.getCalls, [repaired.contextR2Key])
})

test('Context 回写失败时仍返回本轮解析结果，但不会伪装成已修复', async () => {
  const fileKey = 'agent-demo/uploads/session-3/source.xlsx'
  const bucket = createMemoryBucket({
    [fileKey]: { bytes: new Uint8Array([1, 2, 3]), etag: 'source-etag-3' }
  })
  bucket.put = async () => {
    throw new Error('simulated R2 write failure')
  }
  const originalError = console.error
  console.error = () => {}
  try {
    const loaded = await loadWorkbookContextCache(bucket as never, {
      workbookId: 'workbook-3',
      sessionId: 'session-3',
      fileKey,
      contextR2Key: null,
      contextVersion: null,
      sourceEtag: null
    }, async () => context)

    assert.equal(loaded.cacheStatus, 'fallback')
    assert.deepEqual(loaded.context, context)
  } finally {
    console.error = originalError
  }
})

test('v1 缓存保持可解码但必须从源文件重建计算状态，不能标记为 v2 后直接使用', async () => {
  const old = await encodeWorkbookContext({ version: 1 as never, sourceEtag: 'etag-old', ...context })
  assert.equal((await decodeWorkbookContext(old)).version, 1)
  const bucket = createMemoryBucket({ old: { bytes: old, etag: 'cache' }, source: { bytes: new Uint8Array([1]), etag: 'etag-old' } })
  let parsed = 0
  const pending = { ...context, summary: { ...context.summary, calculation: { status: 'pending' as const, formulaCells: 2 } } }
  const originalWarn = console.warn
  console.warn = () => {}
  try {
    const loaded = await loadWorkbookContextCache(bucket as never, {
      workbookId: 'old', sessionId: 'session', fileKey: 'source', contextR2Key: 'old', contextVersion: WORKBOOK_CONTEXT_VERSION, sourceEtag: 'etag-old'
    }, async () => { parsed++; return pending })
    assert.equal(parsed, 1)
    assert.equal(loaded.cacheStatus, 'repaired')
    assert.deepEqual(loaded.context.summary.calculation, pending.summary.calculation)
  } finally { console.warn = originalWarn }
})

test('D1 和 Workflow 已接入 Context 元数据且导出只解析原文件一次', async () => {
  const [migration, toolSource, workflowSource] = await Promise.all([
    readFile(new URL('../migrations/0001_initial_agent_schema.sql', import.meta.url), 'utf8'),
    readFile(new URL('../src/mastra/tools/excelTools.ts', import.meta.url), 'utf8'),
    readFile(new URL('../src/workflows/workbookExport.ts', import.meta.url), 'utf8')
  ])

  assert.match(migration, /context_r2_key/)
  assert.match(migration, /source_etag/)
  assert.match(migration, /context_version/)
  assert.match(toolSource, /return \(await getContext\(\)\)\.summary/)
  assert.doesNotMatch(workflowSource, /inspectWorkbook/)
  assert.equal((workflowSource.match(/processWorkbook\(/g) || []).length, 1)
})

type StoredObject = { bytes: Uint8Array, etag: string }

function createMemoryBucket(initial: Record<string, StoredObject> = {}) {
  const objects = new Map<string, StoredObject>(
    Object.entries(initial).map(([key, value]) => [key, {
      bytes: Uint8Array.from(value.bytes),
      etag: value.etag
    }])
  )
  const getCalls: string[] = []

  return {
    objects,
    getCalls,
    async put(key: string, value: ArrayBuffer | Uint8Array) {
      const bytes = value instanceof Uint8Array
        ? Uint8Array.from(value)
        : new Uint8Array(value.slice(0))
      const stored = { bytes, etag: `etag-${objects.size + 1}` }
      objects.set(key, stored)
      return { key, etag: stored.etag }
    },
    async get(key: string) {
      getCalls.push(key)
      const stored = objects.get(key)
      if (!stored) return null
      return {
        key,
        etag: stored.etag,
        async arrayBuffer() {
          return Uint8Array.from(stored.bytes).buffer
        }
      }
    }
  }
}
