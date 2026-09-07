import assert from 'node:assert/strict'
import test from 'node:test'
import { SignJWT } from 'jose'
import { createBillingDatabase } from './fixtures/billingDatabase.ts'
import { handleAgentRequest } from '../src/http/router.ts'

test('asynchronous download entitlement errors return structured 403, not an unhandled 500', async t => {
  const { env, sqlite } = createBillingDatabase()
  t.after(() => sqlite.close())
  const taskId = crypto.randomUUID(), sessionId = crypto.randomUUID()
  sqlite.prepare("UPDATE users SET free_xlsx_download_task_id='previous-download' WHERE id='billing-test'").run()
  sqlite.prepare(`INSERT INTO excel_agent_tasks(id,user_id,session_id,thread_id,input_r2_key,input_name,output_r2_key,output_name,prompt,status,created_at,updated_at)
    VALUES (?,'billing-test',?,'thread','input.xlsx','input.xlsx','output.xlsx','output.xlsx','test','completed','now','now')`).run(taskId, sessionId)
  env.AGENT_TOKEN_SECRET = 'test-only-download-secret'
  env.LOCAL_BILLING_ENABLED = 'true'
  env.WORKBOOKS = { async get() { return { body: new ReadableStream(), writeHttpMetadata() {}, httpEtag: 'etag' } } } as never
  const token = await new SignJWT({ sid: 'test', scope: ['agent:use'] }).setProtectedHeader({ alg: 'HS256' }).setSubject('billing-test').setIssuer('excelgen-app').setAudience('excelgen-agent').setExpirationTime('1m').sign(new TextEncoder().encode(env.AGENT_TOKEN_SECRET))
  const response = await handleAgentRequest(new Request(`http://127.0.0.1:8791/api/tasks/${taskId}/download?sessionId=${sessionId}`, { headers: { authorization: `Bearer ${token}` } }), env)
  assert.equal(response.status, 403)
  assert.equal((await response.json() as any).error.code, 'SUBSCRIPTION_REQUIRED')
})
