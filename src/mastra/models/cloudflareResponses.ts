import { createOpenAI } from '@ai-sdk/openai'

export const DEFAULT_MODEL_ID = 'openai/gpt-5.6-luna'

const SUPPORTED_MODEL_IDS = new Set<string>([DEFAULT_MODEL_ID])

export type AgentConfigurationErrorCode =
  | 'missing_cloudflare_account_id'
  | 'missing_cloudflare_api_token'
  | 'unsupported_model'

export class AgentConfigurationError extends Error {
  readonly code: AgentConfigurationErrorCode

  constructor(code: AgentConfigurationErrorCode, message: string) {
    super(message)
    this.name = 'AgentConfigurationError'
    this.code = code
  }
}

type CloudflareResponsesModel = ReturnType<ReturnType<typeof createOpenAI>['responses']>
type CloudflareResponsesProvider = {
  responses: (modelId: string) => CloudflareResponsesModel
}
type CloudflareResponsesProviderFactory = (options: Parameters<typeof createOpenAI>[0]) => CloudflareResponsesProvider

export type CloudflareResponsesModelDependencies = {
  providerFactory?: CloudflareResponsesProviderFactory
}

export type CloudflareResponsesModelEnv = {
  CLOUDFLARE_ACCOUNT_ID?: string
  CLOUDFLARE_API_TOKEN?: string
  AI_GATEWAY_ID?: string
  MASTRA_MODEL?: string
}

export function resolveCloudflareResponsesModel(
  env: CloudflareResponsesModelEnv,
  dependencies: CloudflareResponsesModelDependencies = {}
) {
  const modelId = env.MASTRA_MODEL?.trim() || DEFAULT_MODEL_ID
  if (!SUPPORTED_MODEL_IDS.has(modelId)) {
    throw new AgentConfigurationError('unsupported_model', `Unsupported Mastra model: ${modelId}`)
  }

  const accountId = env.CLOUDFLARE_ACCOUNT_ID?.trim()
  if (!accountId) {
    throw new AgentConfigurationError(
      'missing_cloudflare_account_id',
      'CLOUDFLARE_ACCOUNT_ID is required for the Cloudflare Responses API'
    )
  }

  const apiToken = env.CLOUDFLARE_API_TOKEN?.trim()
  if (!apiToken) {
    throw new AgentConfigurationError(
      'missing_cloudflare_api_token',
      'CLOUDFLARE_API_TOKEN is required for the Cloudflare Responses API'
    )
  }

  const gatewayId = env.AI_GATEWAY_ID?.trim()
  const providerFactory = dependencies.providerFactory || createOpenAI
  const provider = providerFactory({
    name: 'cloudflare-responses',
    baseURL: `https://api.cloudflare.com/client/v4/accounts/${encodeURIComponent(accountId)}/ai/v1`,
    apiKey: apiToken,
    headers: gatewayId && gatewayId !== 'default'
      ? { 'cf-aig-gateway-id': gatewayId }
      : undefined
  })

  return provider.responses(modelId)
}
