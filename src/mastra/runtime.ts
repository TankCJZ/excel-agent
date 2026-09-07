import { D1Store } from '@mastra/cloudflare-d1'
import { Agent } from '@mastra/core/agent'
import type { MastraModelConfig } from '@mastra/core/llm'
import { Mastra } from '@mastra/core/mastra'
import { Memory } from '@mastra/memory'

import { resolveCloudflareResponsesModel } from '#agent/mastra/models/cloudflareResponses'
import { buildExcelAgentInstructions } from '#agent/mastra/instructions/excelAgent'
import { buildSpreadsheetExecutionPlan, buildWorkbookPlan } from '#agent/mastra/plans'
import type { ExcelAgentRuntimeInput } from '#agent/mastra/runtimeTypes'
import { appropriateToolUseScorer } from '#agent/mastra/scorers/toolUse'
import { createExcelAgentTools } from '#agent/mastra/tools/excelTools'
import type { AgentWorkerEnv } from '#agent/types'
import type { AgentTaskPlan } from '#agent/types'
import { workbookHistoryProcessors } from '#agent/mastra/historyCompatibility'
import type { FollowupCopyGenerator } from '#agent/mastra/localizeFollowups'

export type { ExcelAgentRuntimeInput } from '#agent/mastra/runtimeTypes'
export { buildSpreadsheetExecutionPlan, buildWorkbookPlan } from '#agent/mastra/plans'

export type ExcelAgentRuntimeDependencies = {
  modelFactory?: (env: AgentWorkerEnv) => MastraModelConfig
  followupCopyGenerator?: FollowupCopyGenerator
}

export function resolveExcelAgentRuntimeModel(
  env: AgentWorkerEnv,
  dependencies: ExcelAgentRuntimeDependencies = {}
) {
  return dependencies.modelFactory
    ? dependencies.modelFactory(env)
    : resolveCloudflareResponsesModel(env)
}

export function createExcelAgentMemory(env: AgentWorkerEnv) {
  const storage = new D1Store({
    id: 'excelgen-agent-d1',
    binding: env.DB,
    tablePrefix: 'mastra_excel_'
  })
  const memory = new Memory({
    storage,
    options: { lastMessages: 16 }
  })
  return { storage, memory }
}

export function createExcelAgentRuntime(
  env: AgentWorkerEnv,
  input: ExcelAgentRuntimeInput,
  dependencies: ExcelAgentRuntimeDependencies = {}
) {
  const { storage, memory } = createExcelAgentMemory(env)
  const { tools, state, getToolDetail, getMutationPlan } = createExcelAgentTools(env, input)
  const agent = new Agent({
    id: 'excelgen-agent',
    name: 'ExcelGen Agent',
    instructions: buildExcelAgentInstructions(Boolean(input.summary), input.requestedOutcome, Boolean(input.supportsWorkbookMutation && env.WORKBOOK_EDITING_ENABLED === 'true')),
    model: () => resolveExcelAgentRuntimeModel(env, dependencies),
    tools,
    inputProcessors: workbookHistoryProcessors(),
    memory
  })
  const mastra = new Mastra({
    storage,
    agents: { excelWorkbookAgent: agent },
    scorers: { appropriateToolUseScorer }
  })

  return {
    mastra,
    agent,
    storage,
    memory,
    getBillingError: () => state.billingError,
    getSummary: () => state.context?.summary || input.summary || null,
    getAnalysis: () => state.analysis,
    isWorkflowRequested: () => state.workflowRequested,
    getPlan: (): AgentTaskPlan | null => {
      const mutation = getMutationPlan()
      if (mutation) return { action: 'edit_workbook', mutation, steps: mutation.operations.map((operation, index) => ({ id: `edit-${index + 1}`, operation: operation.kind })) }
      if (state.spreadsheetPlan) return buildSpreadsheetExecutionPlan(state.spreadsheetPlan)
      return state.analysis && input.summary ? buildWorkbookPlan(state.analysis, state.context?.summary || input.summary) : null
    },
    getResearch: () => state.research,
    didUseExternalWebSearch: () => state.researchSource === 'external',
    getWebSearchFallback: () => state.webSearchFallback,
    getResearchSnapshotR2Key: () => state.researchSnapshotR2Key,
    getSpreadsheetPlan: () => state.spreadsheetPlan,
    getSpreadsheetRecords: () => state.spreadsheetRecords,
    isSpreadsheetWorkflowRequested: () => Boolean(state.spreadsheetPlan),
    getToolDetail
  }
}
