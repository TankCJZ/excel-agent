import { runEvals } from '@mastra/core/evals'
import { createExcelAgentRuntime } from '#agent/mastra/runtime'
import { DEFAULT_MODEL_ID } from '#agent/mastra/models/cloudflareResponses'
import { appropriateToolUseScorer } from '#agent/mastra/scorers/toolUse'
import type { AgentWorkerEnv } from '#agent/types'
import { WORKBOOK_CONTEXT_VERSION, workbookContextKey } from '#agent/workbook/context'
import { createSampleWorkbook, loadWorkbookContext } from '#agent/workbook/service'

export async function runArchitectureEvals(env: AgentWorkerEnv, sessionId: string) {
  const workbookId = 'sales-eval-fixture'
  const fileKey = `agent-demo/evals/${sessionId}/sales-eval-fixture.xlsx`
  let source = await env.WORKBOOKS.head(fileKey)
  if (!source) {
    const sample = await createSampleWorkbook()
    source = await env.WORKBOOKS.put(fileKey, sample.bytes, {
      httpMetadata: {
        contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      }
    })
  }

  const contextR2Key = workbookContextKey(sessionId, workbookId)
  const loaded = await loadWorkbookContext(env.WORKBOOKS, {
    workbookId,
    sessionId,
    fileKey,
    contextR2Key,
    contextVersion: WORKBOOK_CONTEXT_VERSION,
    sourceEtag: source.etag
  })
  const runtime = createExcelAgentRuntime(env, {
    workbookId,
    fileKey,
    contextR2Key: loaded.contextR2Key,
    contextVersion: loaded.contextVersion,
    sourceEtag: loaded.sourceEtag,
    summary: loaded.context.summary,
    sessionId,
    initialContext: loaded.context,
    persistContextMetadata: false
  })
  const result = await runEvals({
    target: runtime.agent,
    data: [
      { input: 'Which region has the highest sales total?' },
      { input: 'Create a chart comparing sales by month.' },
      { input: 'Check this workbook for blanks and duplicate rows.' }
    ],
    scorers: {
      trajectory: [appropriateToolUseScorer]
    },
    targetOptions: {
      maxSteps: 7
    },
    concurrency: 1
  })

  return {
    model: env.MASTRA_MODEL || DEFAULT_MODEL_ID,
    totalItems: result.summary.totalItems,
    verdict: result.verdict || 'scored',
    scores: result.scores
  }
}
