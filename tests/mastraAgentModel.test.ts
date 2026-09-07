import assert from 'node:assert/strict'
import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import test from 'node:test'
import {
  AgentConfigurationError,
  DEFAULT_MODEL_ID,
  resolveCloudflareResponsesModel,
} from '../src/mastra/models/cloudflareResponses.ts'
import { resolveExcelAgentRuntimeModel } from '../src/mastra/runtime.ts'
import {
  MOCK_MODEL_ID,
  createMockExcelModel,
} from './fixtures/mockExcelModel.ts'

const rootDir = resolve(import.meta.dirname, '..')
const config = JSON.parse(
  readFileSync(resolve(rootDir, 'wrangler.jsonc'), 'utf8'),
) as {
  ai?: {
    binding?: string
    remote?: boolean
  }
  vars?: Record<string, unknown>
}

test('Mastra agent defaults to the Cloudflare Responses GPT-5.6 Luna model', () => {
  assert.equal(config.vars?.MASTRA_MODEL, 'openai/gpt-5.6-luna')
  assert.equal(config.vars?.CLOUDFLARE_ACCOUNT_ID, 'ec3c6c647b0cb93c9a237028f499f427')
})

test('Mastra agent exposes the remote Workers AI binding', () => {
  assert.deepEqual(config.ai, {
    binding: 'AI',
    remote: true,
  })
})

test('Mastra agent no longer defaults to the mock planner model', () => {
  assert.notEqual(config.vars?.MASTRA_MODEL, 'mock/excel-planner')
})

test('model ids remain stable', () => {
  assert.equal(DEFAULT_MODEL_ID, 'openai/gpt-5.6-luna')
  assert.equal(MOCK_MODEL_ID, 'mock/excel-planner')
})

test('test model is injected into the runtime without reading the production AI binding', () => {
  const env = { MASTRA_MODEL: DEFAULT_MODEL_ID }
  Object.defineProperty(env, 'AI', {
    get() {
      throw new Error('AI binding must not be read for the mock model')
    },
  })

  const model = resolveExcelAgentRuntimeModel(env as never, {
    modelFactory: () => createMockExcelModel(),
  })

  assert.equal(model.provider, 'excelgen-mock')
  assert.equal(model.modelId, MOCK_MODEL_ID)
})

test('real model uses Cloudflare Responses REST without reading the Workers AI binding', () => {
  const expectedModel = { provider: 'test-cloudflare-responses', modelId: DEFAULT_MODEL_ID }
  let receivedOptions: Record<string, unknown> | undefined
  let receivedModelId: string | undefined

  const model = resolveCloudflareResponsesModel(
    {
      CLOUDFLARE_ACCOUNT_ID: 'account-1',
      CLOUDFLARE_API_TOKEN: 'secret-token',
      MASTRA_MODEL: DEFAULT_MODEL_ID,
    },
    {
      providerFactory: (options) => {
        receivedOptions = options as Record<string, unknown>
        return {
          responses: (modelId: string) => {
            receivedModelId = modelId
            return expectedModel as never
          },
        }
      },
    },
  )

  assert.equal(receivedOptions?.apiKey, 'secret-token')
  assert.equal(receivedOptions?.baseURL, 'https://api.cloudflare.com/client/v4/accounts/account-1/ai/v1')
  assert.equal(receivedOptions?.name, 'cloudflare-responses')
  assert.equal(receivedOptions?.headers, undefined)
  assert.equal(receivedModelId, DEFAULT_MODEL_ID)
  assert.equal(model, expectedModel)
})

test('empty model configuration uses the default model id and optional AI Gateway header', () => {
  let receivedModelId: string | undefined
  let receivedHeaders: Record<string, string> | undefined

  resolveCloudflareResponsesModel(
    {
      CLOUDFLARE_ACCOUNT_ID: 'account-1',
      CLOUDFLARE_API_TOKEN: 'secret-token',
      AI_GATEWAY_ID: 'production',
      MASTRA_MODEL: '',
    },
    {
      providerFactory: (options) => {
        receivedHeaders = options.headers as Record<string, string>
        return {
          responses: (modelId: string) => {
            receivedModelId = modelId
            return { provider: 'test-cloudflare-responses', modelId } as never
          },
        }
      },
    },
  )

  assert.equal(receivedModelId, DEFAULT_MODEL_ID)
  assert.deepEqual(receivedHeaders, { 'cf-aig-gateway-id': 'production' })
})

test('real model requires the Cloudflare account id before creating a provider', () => {
  assertConfigurationError(
    { CLOUDFLARE_API_TOKEN: 'secret-token', MASTRA_MODEL: DEFAULT_MODEL_ID },
    'missing_cloudflare_account_id',
  )
})

test('real model requires a Cloudflare API token before creating a provider', () => {
  assertConfigurationError(
    { CLOUDFLARE_ACCOUNT_ID: 'account-1', MASTRA_MODEL: DEFAULT_MODEL_ID },
    'missing_cloudflare_api_token',
  )
})

test('unsupported model is rejected before reading credentials or creating a provider', () => {
  const env = { MASTRA_MODEL: '@cf/stabilityai/stable-diffusion-xl-base-1.0' }
  Object.defineProperty(env, 'CLOUDFLARE_API_TOKEN', {
    get() {
      throw new Error('Credentials must not be read for an unsupported model')
    },
  })

  assertConfigurationError(env, 'unsupported_model')
})

function assertConfigurationError(env: Record<string, unknown>, code: string) {
  let providerFactoryCalled = false
  assert.throws(
    () => resolveCloudflareResponsesModel(env, {
      providerFactory: () => {
        providerFactoryCalled = true
        throw new Error('provider factory must not be called')
      },
    }),
    (error: unknown) => {
      assert.ok(error instanceof AgentConfigurationError)
      assert.equal(error.code, code)
      return true
    },
  )
  assert.equal(providerFactoryCalled, false)
}

test('mock agent inspects the workbook before planning a schema-dependent action', async () => {
  const [greeting, description, analysis, chart, exportRequest, unsupportedEdit] = await Promise.all([
    generateMockTurn('你好，你能做什么？'),
    generateMockTurn('这是一份什么数据表？'),
    generateMockTurn('分析这份表的数据质量。'),
    generateMockTurn('生成一张销售额趋势图。'),
    generateMockTurn('分析后导出一个新的 Excel 文件。'),
    generateMockTurn('修改 B 列并写回原始 Excel 文件。'),
  ])

  assert.equal(toolName(greeting), null)
  assert.equal(toolName(description), 'inspectWorkbook')
  assert.equal(toolName(analysis), 'inspectWorkbook')
  assert.equal(toolName(chart), 'inspectWorkbook')
  assert.equal(toolName(exportRequest), 'inspectWorkbook')
  assert.equal(toolName(unsupportedEdit), null)
})

test('mock agent treats 工作簿 as a spreadsheet creation request', async () => {
  const result = await generateMockTurn('创建一份季度销售分析工作簿，包含三个地区的示例数据。')

  assert.equal(toolName(result), 'createSpreadsheet')
})

test('mock agent creates a structured plan from inspected sheet and field names', async () => {
  const [analysis, chart, exportRequest] = await Promise.all([
    generateMockTurnAfterInspection('哪个区域的销售额最高？'),
    generateMockTurnAfterInspection('生成一张按月份比较销售额的折线图。'),
    generateMockTurnAfterInspection('分析数据质量后导出一个新的 Excel 文件。'),
  ])

  assert.equal(toolName(analysis), 'analyzeWorkbook')
  assert.equal(toolName(chart), 'createChart')
  assert.equal(toolName(exportRequest), 'exportAnalysisWorkbook')

  const analysisPlan = toolInput(analysis).plan as Record<string, unknown>
  assert.equal(analysisPlan.sheet, '销售数据')
  assert.deepEqual(analysisPlan.groupBy, ['区域'])
  assert.deepEqual(analysisPlan.metrics, [{ operation: 'sum', field: '销售额', alias: '销售额_sum', format: 'number' }])

  const chartPlan = toolInput(chart).plan as Record<string, unknown>
  assert.deepEqual(chartPlan.groupBy, ['月份'])
  assert.equal((chartPlan.chart as Record<string, unknown>).type, 'line')
})

test('mock planner covers names, people count, and total amount with one grouped plan', async () => {
  const result = await generateMockTurnAfterAttendanceInspection('这个表中有哪些人收到了处罚？处罚总金额是多少？')

  assert.equal(toolName(result), 'analyzeWorkbook')
  const plan = toolInput(result).plan as Record<string, unknown>
  assert.deepEqual(plan.groupBy, ['姓名'])
  assert.deepEqual(plan.metrics, [{ operation: 'sum', field: '处罚', alias: '处罚_sum', format: 'number' }])
  assert.deepEqual(plan.where, {
    logic: 'and',
    conditions: [{ field: '处罚', operator: 'gt', value: 0 }]
  })
})

async function generateMockTurn(text: string) {
  return createMockExcelModel().doGenerate({
    prompt: [{ role: 'user', content: [{ type: 'text', text }] }],
  } as never)
}

async function generateMockTurnAfterInspection(text: string) {
  return createMockExcelModel().doGenerate({
    prompt: [
      { role: 'user', content: [{ type: 'text', text }] },
      {
        role: 'assistant',
        content: [{ type: 'tool-call', toolCallId: 'inspect-1', toolName: 'inspectWorkbook', input: '{}' }],
      },
      {
        role: 'tool',
        content: [{
          type: 'tool-result',
          toolCallId: 'inspect-1',
          toolName: 'inspectWorkbook',
          output: {
            type: 'json',
            value: {
              primarySheet: '销售数据',
              headers: ['月份', '区域', '销售额', '成本', '订单数'],
              numericColumns: [
                { name: '销售额' },
                { name: '成本' },
                { name: '订单数' },
              ],
            },
          },
        }],
      },
    ],
  } as never)
}

async function generateMockTurnAfterAttendanceInspection(text: string) {
  return createMockExcelModel().doGenerate({
    prompt: [
      { role: 'user', content: [{ type: 'text', text }] },
      {
        role: 'assistant',
        content: [{ type: 'tool-call', toolCallId: 'inspect-attendance', toolName: 'inspectWorkbook', input: '{}' }],
      },
      {
        role: 'tool',
        content: [{
          type: 'tool-result',
          toolCallId: 'inspect-attendance',
          toolName: 'inspectWorkbook',
          output: {
            type: 'json',
            value: {
              primarySheet: '12月考勤汇总表',
              headers: ['序号', '姓名', '实薪（天）', '处罚'],
              numericColumns: [{ name: '序号' }, { name: '实薪（天）' }, { name: '处罚' }],
            },
          },
        }],
      },
    ],
  } as never)
}

function toolName(result: Awaited<ReturnType<typeof generateMockTurn>>) {
  const part = result.content.find(item => item.type === 'tool-call')
  return part?.type === 'tool-call' ? part.toolName : null
}

function toolInput(result: Awaited<ReturnType<typeof generateMockTurn>>) {
  const part = result.content.find(item => item.type === 'tool-call')
  assert.ok(part?.type === 'tool-call')
  return JSON.parse(part.input) as Record<string, unknown>
}
