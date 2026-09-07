import assert from 'node:assert/strict'
import test from 'node:test'
import { decodeStoredJson, readTaskResult, parseMutationPayload, readTaskPlan } from '../src/persistence/taskProtocol.ts'
import { createSampleWorkbook, parseWorkbookBytes } from '../src/workbook/service.ts'

test('legacy generation results without discriminator, new metadata or sources remain readable', async () => {
  const sample = await createSampleWorkbook()
  const summary = (await parseWorkbookBytes(sample.bytes)).summary
  const { sheets: _sheets, ...oldSummary } = summary
  const old = { taskId: 'task', outputR2Key: 'old/output.xlsx', outputName: 'old.xlsx', summary: oldSummary, preview: sample.preview, sizeBytes: sample.bytes.length }
  const result = readTaskResult(old, 'spreadsheet_generation')
  assert.equal(result?.kind, 'spreadsheet_generation')
  assert.equal(result?.previewSheetName, 'Sales Data')
  assert.equal(result?.summary.calculation, undefined)
  assert.deepEqual(result?.summary.sheets, [])
  assert.equal(readTaskResult({ ...old, schemaVersion: 999 }, 'spreadsheet_generation'), null)
  assert.equal(readTaskResult({ ...old, kind: 'future-tool' }, 'spreadsheet_generation'), null)
})

test('corrupted and unknown mutation plans cannot be cast into executable payloads', () => {
  assert.equal(decodeStoredJson('{broken'), null)
  assert.equal(readTaskResult([], 'workbook_export'), null)
  assert.equal(readTaskPlan({ action: 'future-action', schemaVersion: 999 }), null)
  assert.equal(readTaskResult({ kind: 'workbook_mutation', schemaVersion: 2 }, 'workbook_mutation'), null)
  assert.throws(() => parseMutationPayload({ kind: 'workbook_mutation', taskId: 'task', plan: { kind: 'workbook_mutation', schemaVersion: 999 } }))
})
