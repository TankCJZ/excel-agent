import assert from 'node:assert/strict'
import test from 'node:test'
import { createBillingDatabase } from './fixtures/billingDatabase.ts'
import { createSampleWorkbook } from '../src/workbook/service.ts'
import { adoptTaskWorkbook } from '../src/persistence/adoptTaskWorkbook.ts'

async function fixture() {
  const db = createBillingDatabase()
  const input = { userId: 'billing-test', sessionId: crypto.randomUUID(), threadId: 'conversation', taskId: crypto.randomUUID(), workbookRevision: 0 }
  db.sqlite.prepare(`INSERT INTO excel_agent_conversations(id,user_id,session_id,title,created_at,updated_at) VALUES (?,'billing-test',?,'Test','now','now')`).run(input.threadId, input.sessionId)
  db.sqlite.prepare(`INSERT INTO excel_agent_tasks(id,user_id,session_id,thread_id,task_type,input_r2_key,input_name,output_r2_key,output_name,result_json,prompt,status,created_at,updated_at)
    VALUES (?,'billing-test',?,?,'spreadsheet_generation','','','agent-demo/results/test/output.xlsx','output.xlsx','invalid old JSON','test','completed','now','now')`).run(input.taskId, input.sessionId, input.threadId)
  const bytes = (await createSampleWorkbook()).bytes
  const putKeys: string[] = []
  db.env.WORKBOOKS = {
    async get() { return { size: bytes.length, etag: 'etag', async arrayBuffer() { return bytes.slice().buffer } } },
    async put(key: string) { putKeys.push(key); return { etag: 'context' } }
  } as never
  return { ...db, input, putKeys }
}

test('legacy result with invalid JSON can be adopted from the owned XLSX, idempotently without credits or downloads', async t => {
  const f = await fixture(); t.after(() => f.sqlite.close())
  const result = await adoptTaskWorkbook(f.env, f.input)
  assert.equal(result.workbook.summary.primarySheet, 'Sales Data')
  assert.equal(result.conversation.workbookRevision, 1)
  const repeated = await adoptTaskWorkbook(f.env, f.input)
  assert.equal(repeated.workbook.workbookId, result.workbook.workbookId)
  assert.equal(repeated.conversation.workbookRevision, 1)
  assert.equal(f.putKeys.length, 1)
  assert.equal(f.balance(), 30)
  assert.equal(f.sqlite.prepare("SELECT free_xlsx_download_task_id FROM users WHERE id='billing-test'").get()!.free_xlsx_download_task_id, null)
})

test('adoption rejects other owners/conversations, stale selections, expired files and deletion', async t => {
  const f = await fixture(); t.after(() => f.sqlite.close())
  await assert.rejects(adoptTaskWorkbook(f.env, { ...f.input, userId: 'someone-else' }), /not found/)
  await assert.rejects(adoptTaskWorkbook(f.env, { ...f.input, threadId: 'other-thread' }), /not found/)
  await assert.rejects(adoptTaskWorkbook(f.env, { ...f.input, workbookRevision: 1 }), /changed/)
  const bucket = f.env.WORKBOOKS
  f.env.WORKBOOKS = { async get() { return null } } as never
  await assert.rejects(adoptTaskWorkbook(f.env, f.input), /no longer available/)
  f.env.WORKBOOKS = bucket
  f.sqlite.prepare("INSERT INTO excel_agent_conversation_tombstones VALUES ('conversation','billing-test',?,'now')").run(f.input.sessionId)
  await assert.rejects(adoptTaskWorkbook(f.env, f.input), /deleted/i)
  assert.equal(f.putKeys.length, 0)
})

test('concurrent result adoption creates a single version and advances revision once', async t => {
  const f = await fixture(); t.after(() => f.sqlite.close())
  const results = await Promise.all([adoptTaskWorkbook(f.env, f.input), adoptTaskWorkbook(f.env, f.input)])
  assert.equal(results[0]!.workbook.workbookId, results[1]!.workbook.workbookId)
  assert.equal(results[0]!.conversation.workbookRevision, 1)
  assert.equal(f.sqlite.prepare('SELECT COUNT(*) AS count FROM excel_agent_workbooks').get()!.count, 1)
})
