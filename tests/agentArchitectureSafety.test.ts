import assert from 'node:assert/strict'
import test from 'node:test'
import { SignJWT } from 'jose'
import { z } from 'zod'
import {
  analysisPlanSchema,
  spreadsheetCreationInputSchema,
  spreadsheetPlanSchema,
  type WorkbookAnalysis
} from '#shared/agent/contracts'
import { authenticateAgentRequest } from '#agent/platform/auth'
import { workflowEventId } from '#agent/persistence/agentRepository'
import { HttpError, normalizeHttpError } from '#agent/http/responses'
import { recordsFromAnalysis } from '#agent/spreadsheet/generation'
import type { AgentWorkerEnv } from '#agent/types'

const baseColumns = [
  { key: 'name', label: 'Name', description: 'Entity name', type: 'text' as const, required: true },
  { key: 'value', label: 'Value', description: 'Entity value', type: 'number' as const, required: false }
]

test('uses deterministic workflow event ids across retries', () => {
  assert.equal(
    workflowEventId('task-1', 'spreadsheet.generating'),
    workflowEventId('task-1', 'spreadsheet.generating')
  )
  assert.notEqual(
    workflowEventId('task-1', 'spreadsheet.generating'),
    workflowEventId('task-1', 'workflow.processing')
  )
})

test('rejects SheetJS-invalid and internal spreadsheet sheet names', () => {
  const basePlan = {
    title: 'Safe workbook',
    sheetName: 'Data',
    purpose: 'Store structured records',
    dataMode: 'blank_template' as const,
    columns: baseColumns,
    maxRows: 20,
    includeSources: true
  }
  assert.equal(spreadsheetPlanSchema.safeParse(basePlan).success, true)
  for (const sheetName of ['bad/name', 'bad:name', "'quoted", '__sources', '__METADATA']) {
    assert.equal(spreadsheetPlanSchema.safeParse({ ...basePlan, sheetName }).success, false, sheetName)
  }
})

test('rejects duplicate metric aliases and positional workbook-analysis mappings', () => {
  const plan = {
    goal: 'Compare totals',
    sheet: 'Data',
    groupBy: ['Region'],
    metrics: [
      { operation: 'sum' as const, field: 'Sales', alias: 'Total' },
      { operation: 'average' as const, field: 'Sales', alias: 'total' }
    ]
  }
  assert.equal(analysisPlanSchema.safeParse(plan).success, false)

  const workbookPlan = {
    title: 'Analysis export',
    sheetName: 'Analysis',
    purpose: 'Export deterministic results',
    dataMode: 'workbook_analysis' as const,
    columns: baseColumns,
    maxRows: 20,
    includeSources: false
  }
  assert.equal(spreadsheetPlanSchema.safeParse(workbookPlan).success, false)
})

test('captures conversation records in the validated creation protocol', () => {
  const plan = {
    title: 'Conversation export',
    sheetName: 'Data',
    purpose: 'Export facts supplied by the user',
    dataMode: 'conversation_data' as const,
    columns: baseColumns,
    maxRows: 20,
    includeSources: false
  }
  assert.equal(spreadsheetCreationInputSchema.safeParse({ plan, records: [] }).success, false)
  assert.equal(spreadsheetCreationInputSchema.safeParse({
    plan,
    records: [{ cells: [
      { key: 'name', value: 'Alpha' },
      { key: 'value', value: 10 }
    ] }]
  }).success, true)
  assert.equal(spreadsheetCreationInputSchema.safeParse({
    plan,
    records: [{ cells: [
      { key: 'value', value: 10 },
      { key: 'undeclared', value: 'no' }
    ] }]
  }).success, false)
})

test('exposes explicit record cells to strict Responses tool schemas', () => {
  const schema = z.toJSONSchema(spreadsheetCreationInputSchema) as {
    properties?: {
      records?: {
        items?: {
          properties?: {
            cells?: {
              items?: { properties?: Record<string, unknown>, required?: string[] }
            }
          }
        }
      }
    }
  }
  const cell = schema.properties?.records?.items?.properties?.cells?.items

  assert.deepEqual(Object.keys(cell?.properties || {}).sort(), ['key', 'value'])
  assert.deepEqual(cell?.required?.sort(), ['key', 'value'])
})

test('maps workbook analysis records by explicit source field instead of position', () => {
  const analysis: WorkbookAnalysis = {
    kind: 'analysis',
    question: 'Compare values',
    metrics: [],
    table: { columns: ['Value', 'Name'], rows: [[10, 'Alpha']] },
    chart: null,
    evidence: [{ sheet: 'Data', range: 'A1:B2', basis: 'validated_plan' }],
    stats: { sourceRowCount: 1, matchingRowCount: 1, outputRowCount: 1, groupCount: 0 }
  }
  const records = recordsFromAnalysis([
    { ...baseColumns[0]!, sourceField: 'Name' },
    { ...baseColumns[1]!, sourceField: 'Value' }
  ], analysis)
  assert.deepEqual(records, [{ name: 'Alpha', value: 10 }])
})

test('requires signed Agent tokens outside loopback and never exposes internal errors', async () => {
  const secret = 'test-agent-secret-with-enough-entropy'
  const token = await new SignJWT({ sid: 'session-1', scope: ['agent:use'] })
    .setProtectedHeader({ alg: 'HS256' })
    .setIssuer('excelgen-app')
    .setAudience('excelgen-agent')
    .setSubject('user-1')
    .setExpirationTime('5m')
    .sign(new TextEncoder().encode(secret))
  const env = { AGENT_TOKEN_SECRET: secret } as AgentWorkerEnv
  const identity = await authenticateAgentRequest(new Request('https://agent.example.com/api/chat', {
    headers: { authorization: `Bearer ${token}` }
  }), env)
  assert.equal(identity.userId, 'user-1')
  assert.equal(identity.billingEnabled, true)

  const localDemo = await authenticateAgentRequest(
    new Request('http://127.0.0.1:8788/api/chat'),
    { AGENT_ENV: 'development' } as AgentWorkerEnv
  )
  assert.equal(localDemo.localDemo, true)
  assert.equal(localDemo.billingEnabled, false)

  await assert.rejects(
    authenticateAgentRequest(
      new Request('http://127.0.0.1:8788/api/chat'),
      { AGENT_ENV: 'development', LOCAL_BILLING_ENABLED: 'true' } as AgentWorkerEnv
    ),
    (error: unknown) => error instanceof HttpError && error.code === 'AUTH_NOT_CONFIGURED'
  )

  const localUnbilled = await authenticateAgentRequest(new Request('http://127.0.0.1:8788/api/chat', {
    headers: { authorization: `Bearer ${token}` }
  }), { AGENT_ENV: 'development', AGENT_TOKEN_SECRET: secret, LOCAL_BILLING_ENABLED: 'false' } as AgentWorkerEnv)
  assert.equal(localUnbilled.localDemo, false)
  assert.equal(localUnbilled.billingEnabled, false)

  const localBilled = await authenticateAgentRequest(new Request('http://127.0.0.1:8788/api/chat', {
    headers: { authorization: `Bearer ${token}` }
  }), { AGENT_ENV: 'development', AGENT_TOKEN_SECRET: secret, LOCAL_BILLING_ENABLED: 'true' } as AgentWorkerEnv)
  assert.equal(localBilled.localDemo, false)
  assert.equal(localBilled.billingEnabled, true)

  await assert.rejects(
    authenticateAgentRequest(new Request('https://agent.example.com/api/chat'), {} as AgentWorkerEnv),
    (error: unknown) => error instanceof HttpError && error.code === 'AUTH_NOT_CONFIGURED'
  )
  assert.equal(normalizeHttpError(new Error('database password leaked')).message, 'Unexpected server error')
})
