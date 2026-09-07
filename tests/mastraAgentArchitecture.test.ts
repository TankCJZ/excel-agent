import assert from 'node:assert/strict'
import { readdir, readFile } from 'node:fs/promises'
import { resolve } from 'node:path'
import test from 'node:test'

const rootDir = resolve(import.meta.dirname, '..')
const agentRoot = resolve(rootDir, 'src')

test('keeps the Worker entrypoint composition-only and domain modules grouped', async () => {
  const entries = await readdir(agentRoot, { withFileTypes: true })
  const topLevelFiles = entries
    .filter(entry => entry.isFile())
    .map(entry => entry.name)
    .sort()
  assert.deepEqual(topLevelFiles, ['contracts.ts', 'index.ts', 'types.ts'])

  const indexSource = await readFile(resolve(agentRoot, 'index.ts'), 'utf8')
  assert.ok(indexSource.split(/\r?\n/).length <= 15)
  assert.match(indexSource, /handleAgentRequest/)
  assert.match(indexSource, /ExcelAgentWorkflow/)
  assert.match(indexSource, /SpreadsheetGenerationWorkflow/)
})

test('keeps Worker code independent from Nuxt server internals', async () => {
  const sources = await readTypeScriptSources(agentRoot)
  for (const [file, source] of sources) {
    assert.doesNotMatch(source, /#database\/|server\/db\//, file)
  }

  const sharedSchema = await readFile(resolve(rootDir, 'src/shared/agentSchema.ts'), 'utf8')
  assert.match(sharedSchema, /excelAgentWorkbooks/)
  assert.match(sharedSchema, /excelAgentTaskEvents/)
})

test('keeps test models completely outside the production Worker dependency graph', async () => {
  const sources = await readTypeScriptSources(agentRoot)
  for (const [file, source] of sources) {
    assert.doesNotMatch(source, /(?:from\s+|import\s*\()['"]ai\/test['"]/, file)
    assert.doesNotMatch(source, /mockExcelModel|mock\/excel-planner|tests\/fixtures/, file)
  }

  const runtimeSource = await readFile(resolve(agentRoot, 'mastra/runtime.ts'), 'utf8')
  assert.match(runtimeSource, /modelFactory/)
  assert.match(runtimeSource, /resolveCloudflareResponsesModel/)
})

test('tracks all configured Worker bindings in generated Env types', async () => {
  const generatedTypes = await readFile(resolve(rootDir, 'worker-configuration.d.ts'), 'utf8')
  for (const binding of [
    'AGENT_REQUEST_TIMEOUT_MS',
    'AGENT_ANSWER_TIMEOUT_MS',
    'EXCEL_AGENT_WORKFLOW',
    'SPREADSHEET_GENERATION_WORKFLOW',
    'WORKBOOKS',
    'DB',
    'AI'
  ]) {
    assert.match(generatedTypes, new RegExp(`\\b${binding}\\b`), binding)
  }
})

async function readTypeScriptSources(directory: string): Promise<Array<[string, string]>> {
  const output: Array<[string, string]> = []
  for (const entry of await readdir(directory, { withFileTypes: true })) {
    const path = resolve(directory, entry.name)
    if (entry.isDirectory()) output.push(...await readTypeScriptSources(path))
    else if (entry.name.endsWith('.ts')) output.push([path, await readFile(path, 'utf8')])
  }
  return output
}
