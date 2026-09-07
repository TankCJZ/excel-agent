import type { AgentIdentity } from '#agent/platform/auth'
import { sampleFileRequestSchema, taskIdSchema, taskQuerySchema } from '#agent/contracts'
import { createExcelAgentWorkbook, getTaskForSession } from '#agent/persistence/agentRepository'
import { attachmentContentDisposition } from '#agent/http/fileDownload'
import { corsHeaders, HttpError, internalErrorMessage } from '#agent/http/responses'
import type { AgentWorkerEnv } from '#agent/types'
import {
  createSampleWorkbook,
  parseWorkbookBytes,
  type WorkbookAnalysisContext
} from '#agent/workbook/service'
import {
  WORKBOOK_CONTEXT_VERSION,
  workbookContextKey,
  writeWorkbookContext
} from '#agent/workbook/context'
import { authorizeWorkbookDownload, getBillingAccess } from '#agent/billing/billingRepository'
import { getPlanEntitlements } from '#shared/billing/planCatalog'

const workbookMime = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
const allowedExtensions = new Set(['xlsx', 'xls', 'csv'])

export async function createSampleWorkbookForUser(request: Request, env: AgentWorkerEnv, identity: AgentIdentity) {
  const body = sampleFileRequestSchema.parse(await request.json())
  const sample = await createSampleWorkbook()
  const fileName = 'excelgen-sales-demo.xlsx'
  const workbookId = crypto.randomUUID()
  const fileKey = `agent-demo/uploads/${body.sessionId}/${workbookId}-${fileName}`
  const context = await parseWorkbookBytes(sample.bytes)
  return persistWorkbook(env, {
    workbookId,
    userId: identity.userId,
    sessionId: body.sessionId,
    fileKey,
    fileName,
    bytes: sample.bytes,
    contentType: workbookMime,
    source: 'sample',
    context
  })
}

export async function uploadWorkbookForUser(request: Request, env: AgentWorkerEnv, identity: AgentIdentity) {
  const form = await request.formData()
  const sessionId = sampleFileRequestSchema.shape.sessionId.parse(form.get('sessionId'))
  const rawFile = form.get('file')
  if (!rawFile || typeof rawFile === 'string') throw new HttpError(400, 'VALIDATION_ERROR', 'A workbook file is required')
  const file = rawFile as unknown as File

  const access = identity.localDemo
    ? { planKey: 'max' as const, entitlements: getPlanEntitlements('max') }
    : await getBillingAccess(env, identity.userId)
  const hardLimit = Number(env.MAX_UPLOAD_BYTES || getPlanEntitlements('max').maxUploadBytes)
  const maxUploadBytes = Math.min(hardLimit, access.entitlements.maxUploadBytes)
  if (file.size <= 0 || file.size > maxUploadBytes) {
    throw new HttpError(413, 'FILE_TOO_LARGE', `The ${access.planKey} plan supports workbooks up to ${formatMegabytes(maxUploadBytes)}`, {
      plan: access.planKey,
      maxUploadBytes,
      actualBytes: file.size
    })
  }
  const extension = file.name.split('.').pop()?.toLowerCase() || ''
  if (!allowedExtensions.has(extension)) {
    throw new HttpError(400, 'VALIDATION_ERROR', 'Only XLSX, XLS, and CSV files are supported')
  }

  const fileName = sanitizeFileName(file.name)
  const workbookId = crypto.randomUUID()
  const fileKey = `agent-demo/uploads/${sessionId}/${workbookId}-${fileName}`
  const bytes = await file.arrayBuffer()
  let context: WorkbookAnalysisContext
  try {
    context = await parseWorkbookBytes(bytes)
    assertWorkbookWithinPlan(context, access.entitlements)
  } catch (error) {
    if (error instanceof HttpError) throw error
    console.warn('Workbook parsing failed', { error: internalErrorMessage(error) })
    throw new HttpError(400, 'INVALID_WORKBOOK', 'The uploaded workbook could not be parsed')
  }

  return persistWorkbook(env, {
    workbookId,
    userId: identity.userId,
    sessionId,
    fileKey,
    fileName,
    bytes,
    contentType: file.type || workbookMime,
    source: 'upload',
    context
  })
}

function assertWorkbookWithinPlan(
  context: WorkbookAnalysisContext,
  entitlements: ReturnType<typeof getPlanEntitlements>
) {
  if (context.sheets.length > entitlements.maxWorkbookSheets) {
    throw new HttpError(422, 'WORKBOOK_LIMIT_EXCEEDED', 'The workbook contains too many worksheets for this plan', {
      limit: entitlements.maxWorkbookSheets,
      actual: context.sheets.length,
      kind: 'worksheets'
    })
  }
  const cells = context.sheets.reduce((total, sheet) => total + sheet.rows.reduce(
    (sheetTotal, row) => sheetTotal + row.filter(cell => cell !== null && cell !== '').length,
    0
  ), 0)
  if (cells > entitlements.maxWorkbookCells) {
    throw new HttpError(422, 'WORKBOOK_LIMIT_EXCEEDED', 'The workbook contains too many populated cells for this plan', {
      limit: entitlements.maxWorkbookCells,
      actual: cells,
      kind: 'cells'
    })
  }
}

function formatMegabytes(bytes: number) {
  return `${Math.round(bytes / (1024 * 1024))}MB`
}

export async function downloadTaskWorkbook(
  request: Request,
  env: AgentWorkerEnv,
  identity: AgentIdentity,
  rawTaskId: string,
  rawSessionId: string | null
) {
  const taskId = taskIdSchema.parse(rawTaskId)
  const query = taskQuerySchema.parse({ sessionId: rawSessionId })
  const record = await getTaskForSession(env, taskId, identity.userId, query.sessionId)
  if (!record?.task.outputR2Key || record.task.status !== 'completed') {
    throw new HttpError(404, 'NOT_FOUND', 'Completed workbook result was not found')
  }
  const object = await env.WORKBOOKS.get(record.task.outputR2Key)
  if (!object) throw new HttpError(404, 'NOT_FOUND', 'R2 result object was not found')
  if (!identity.localDemo) {
    await authorizeWorkbookDownload(env, {
      userId: identity.userId,
      taskId,
      completedAt: record.task.completedAt
    })
  }
  const headers = corsHeaders(request, env)
  object.writeHttpMetadata(headers)
  headers.set('etag', object.httpEtag)
  headers.set('content-type', workbookMime)
  headers.set('content-disposition', attachmentContentDisposition(record.task.outputName || 'excelgen-result.xlsx'))
  return new Response(object.body, { headers })
}

async function persistWorkbook(
  env: AgentWorkerEnv,
  input: {
    workbookId: string
    userId: string
    sessionId: string
    fileKey: string
    fileName: string
    bytes: ArrayBuffer | Uint8Array
    contentType: string
    source: 'sample' | 'upload'
    context: WorkbookAnalysisContext
  }
) {
  const contextR2Key = workbookContextKey(input.sessionId, input.workbookId)
  let sourceStored = false
  let contextStored = false
  try {
    const source = await env.WORKBOOKS.put(input.fileKey, input.bytes, {
      httpMetadata: { contentType: input.contentType },
      customMetadata: {
        sessionId: input.sessionId,
        originalName: input.fileName,
        source: input.source
      }
    })
    sourceStored = true
    await writeWorkbookContext(env.WORKBOOKS, contextR2Key, input.context, source.etag)
    contextStored = true
    return await createExcelAgentWorkbook(env, {
      id: input.workbookId,
      userId: input.userId,
      sessionId: input.sessionId,
      r2Key: input.fileKey,
      fileName: input.fileName,
      sizeBytes: input.bytes.byteLength,
      summary: input.context.summary,
      contextR2Key,
      sourceEtag: source.etag,
      contextVersion: WORKBOOK_CONTEXT_VERSION,
      parsedAt: new Date().toISOString()
    })
  } catch (error) {
    const cleanup: Promise<void>[] = []
    if (sourceStored) cleanup.push(env.WORKBOOKS.delete(input.fileKey))
    if (contextStored) cleanup.push(env.WORKBOOKS.delete(contextR2Key))
    await Promise.allSettled(cleanup)
    throw error
  }
}

function sanitizeFileName(value: string) {
  return value.replace(/[^\p{L}\p{N}._-]+/gu, '-').replace(/^-+|-+$/g, '').slice(0, 140) || 'workbook.xlsx'
}
