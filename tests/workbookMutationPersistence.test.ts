import assert from 'node:assert/strict'
import test from 'node:test'
import { createBillingDatabase } from './fixtures/billingDatabase.ts'
import { createMutationTask, publishMutationResult, stageMutationResult, stableMutationId } from '../src/persistence/workbookMutations.ts'
import { reserveTurnCredits } from '../src/billing/billingRepository.ts'
import type { WorkbookMutationPayload, WorkbookMutationResult } from '../src/types.ts'
import { reconcileWorkbookMutations } from '../src/workflows/mutationRecovery.ts'
import { discardMutationArtifacts } from '../src/workflows/mutationCleanup.ts'

async function fixture() {
  const db = createBillingDatabase()
  db.env.WORKBOOKS = { delete: async () => {} } as never
  db.env.EXCEL_AGENT_WORKFLOW = { get: async () => ({ status: async () => ({ status: 'errored' }) }) } as never
  const now = '2026-09-05T00:00:00.000Z'
  db.sqlite.prepare(`INSERT INTO excel_agent_conversations (id,user_id,session_id,title,workbook_id,created_at,updated_at) VALUES ('conversation','billing-test','session','Test','source',?,?)`).run(now, now)
  db.sqlite.prepare(`INSERT INTO excel_agent_workbooks (id,user_id,session_id,r2_key,file_name,size_bytes,summary_json,created_at,updated_at) VALUES ('source','billing-test','session','uploads/source.xlsx','source.xlsx',100,'{}',?,?)`).run(now, now)
  const reservation = await reserveTurnCredits(db.env, { userId: 'billing-test', conversationId: 'conversation', turnId: 'turn', estimatedCredits: 14 })
  const payload: WorkbookMutationPayload = {
    kind: 'workbook_mutation', taskId: 'task', userId: 'billing-test', sessionId: 'session', threadId: 'conversation', turnId: 'turn',
    inputR2Key: 'uploads/source.xlsx', inputName: 'source.xlsx',
    plan: { kind: 'workbook_mutation', schemaVersion: 1, sourceWorkbookId: 'source', sourceEtag: 'source-etag', title: 'Edit', operations: [{ kind: 'set_values', sheet: 'Sheet1', start: 'B2', values: [[2]], overwrite: false }] },
    resultWorkbookId: 'version2', sourceRevision: 0, creditReservationId: reservation.id, creditTurnId: 'turn', creditToolNames: ['inspectWorkbookForEditing']
  }
  await createMutationTask(db.env, payload, { action: 'edit_workbook', mutation: payload.plan, steps: [{ id: '1', operation: 'set_values' }] }, 'edit')
  db.sqlite.prepare("UPDATE excel_agent_tasks SET status = 'processing' WHERE id = 'task'").run()
  const result: WorkbookMutationResult = {
    kind: 'workbook_mutation', schemaVersion: 1, taskId: 'task', outputR2Key: 'results/task.xlsx', outputName: 'edited.xlsx',
    summary: { primarySheet: 'Sheet1', sheetNames: ['Sheet1'], rowCount: 1, columnCount: 1, headers: ['Title'], sheets: [], preview: [], numericColumns: [], blankCellCount: 0, duplicateRowCount: 0 },
    analysis: null, preview: [], previewSheetName: 'Sheet1', previewRowCount: 1, sizeBytes: 100,
    changes: { changedCells: 1, formulaCells: 0, addedSheets: [], addedColumns: [], pendingCalculation: false, convertedFrom: null, sheets: [] },
    sourceWorkbookId: 'source', workbookId: 'version2', sourceEtag: 'output-etag', contextR2Key: 'results/context.json.gz', contextVersion: 1
  }
  await stageMutationResult(db.env, 'task', result)
  return { db, payload, result }
}

test('publishes workbook lineage, task completion and credit settlement in one atomic transaction', async t => {
  const { db, payload, result } = await fixture()
  t.after(() => db.sqlite.close())
  await publishMutationResult(db.env, payload, result)
  assert.equal(db.balance(), 16)
  assert.equal(db.sqlite.prepare("SELECT status FROM excel_agent_tasks WHERE id='task'").get()!.status, 'completed')
  assert.equal(db.sqlite.prepare("SELECT workbook_id FROM excel_agent_conversations WHERE id='conversation'").get()!.workbook_id, 'version2')
  const workbook = db.sqlite.prepare("SELECT parent_workbook_id,root_workbook_id,source_task_id FROM excel_agent_workbooks WHERE id='version2'").get()!
  assert.deepEqual({ ...workbook }, { parent_workbook_id: 'source', root_workbook_id: 'source', source_task_id: 'task' })
  await publishMutationResult(db.env, payload, result)
  assert.equal(db.balance(), 16)
  assert.equal(db.count('credit_usage_events'), 4)
})

test('database failure rolls back publication AND settlement; retry finishes exactly once', async t => {
  const { db, payload, result } = await fixture()
  t.after(() => db.sqlite.close())
  db.failNext('INSERT INTO excel_agent_workbooks')
  await assert.rejects(publishMutationResult(db.env, payload, result), /Injected/)
  assert.equal(db.sqlite.prepare("SELECT status FROM credit_reservations WHERE turn_id='turn'").get()!.status, 'reserved')
  assert.equal(db.sqlite.prepare("SELECT status FROM excel_agent_tasks WHERE id='task'").get()!.status, 'processing')
  assert.equal(db.sqlite.prepare("SELECT id FROM excel_agent_workbooks WHERE id='version2'").get(), undefined)
  await publishMutationResult(db.env, payload, result)
  assert.equal(db.sqlite.prepare("SELECT state FROM excel_agent_mutations WHERE task_id='task'").get()!.state, 'completed')
})

test('publication does not override user file selection or a newer revision', async t => {
  const { db, payload, result } = await fixture()
  t.after(() => db.sqlite.close())
  db.sqlite.prepare("UPDATE excel_agent_conversations SET workbook_revision=1 WHERE id='conversation'").run()
  await publishMutationResult(db.env, payload, result)
  assert.equal(db.sqlite.prepare("SELECT workbook_id FROM excel_agent_conversations WHERE id='conversation'").get()!.workbook_id, 'source')
  assert.equal(db.sqlite.prepare("SELECT status FROM excel_agent_tasks WHERE id='task'").get()!.status, 'completed')
})

test('cancelled tasks cannot publish or settle; no-op has no edit/artifact charge', async t => {
  const { db, payload, result } = await fixture()
  t.after(() => db.sqlite.close())
  db.sqlite.prepare("UPDATE excel_agent_mutations SET state='cancelled' WHERE task_id='task'").run()
  await assert.rejects(publishMutationResult(db.env, payload, result), /publication/)
  assert.equal(db.sqlite.prepare("SELECT status FROM credit_reservations WHERE turn_id='turn'").get()!.status, 'reserved')
  db.sqlite.prepare("UPDATE excel_agent_mutations SET state='staged' WHERE task_id='task'").run()
  const noop = { ...result, workbookId: 'source', changes: { ...result.changes, changedCells: 0 } }
  await stageMutationResult(db.env, 'task', noop)
  await publishMutationResult(db.env, payload, noop)
  assert.equal(db.balance(), 26)
  assert.equal(db.sqlite.prepare("SELECT id FROM excel_agent_workbooks WHERE id='version2'").get(), undefined)
  assert.equal(db.sqlite.prepare("SELECT workbook_revision FROM excel_agent_conversations WHERE id='conversation'").get()!.workbook_revision, 0)
})

test('an output ID collision cannot publish another workbook or settle credits', async t => {
  const { db, payload, result } = await fixture()
  t.after(() => db.sqlite.close())
  db.sqlite.prepare(`INSERT INTO excel_agent_workbooks (id,user_id,session_id,r2_key,file_name,size_bytes,summary_json,created_at,updated_at)
    VALUES ('version2','billing-test','different-session','unrelated.xlsx','unrelated.xlsx',1,'{}','now','now')`).run()
  await assert.rejects(publishMutationResult(db.env, payload, result), /conflicts/)
  assert.equal(db.sqlite.prepare("SELECT status FROM credit_reservations WHERE turn_id='turn'").get()!.status, 'reserved')
  assert.equal(db.sqlite.prepare("SELECT status FROM excel_agent_tasks WHERE id='task'").get()!.status, 'processing')
  assert.equal(db.sqlite.prepare("SELECT r2_key FROM excel_agent_workbooks WHERE id='version2'").get()!.r2_key, 'unrelated.xlsx')
})

test('stable mutation IDs isolate accounts/turns and satisfy UUID validators', async () => {
  const id = await stableMutationId('u', 'c', 't')
  assert.equal(id, await stableMutationId('u', 'c', 't'))
  assert.notEqual(id, await stableMutationId('u2', 'c', 't'))
  assert.match(id, /^[a-f0-9]{8}-[a-f0-9]{4}-5[a-f0-9]{3}-[89ab][a-f0-9]{3}-[a-f0-9]{12}$/)
})

test('only one active workbook edit per conversation may be queued', async t => {
  const { db, payload } = await fixture()
  t.after(() => db.sqlite.close())
  const competing = { ...payload, taskId: 'task2', turnId: 'turn2' }
  await assert.rejects(createMutationTask(db.env, competing, { action: 'edit_workbook', mutation: payload.plan, steps: [] }, 'another edit'), /already has a workbook edit running/)
  assert.equal(db.sqlite.prepare("SELECT id FROM excel_agent_tasks WHERE id='task2'").get(), undefined)
})

test('deletion fence prevents both publication and settlement and recovery refunds the hold', async t => {
  const { db, payload, result } = await fixture()
  t.after(() => db.sqlite.close())
  db.sqlite.prepare("INSERT INTO excel_agent_conversation_tombstones VALUES ('conversation','billing-test','session','now')").run()
  await assert.rejects(publishMutationResult(db.env, payload, result), /publication/)
  assert.equal(db.count('credit_usage_events'), 0)
  db.sqlite.prepare("UPDATE excel_agent_mutations SET updated_at='2020-01-01' WHERE task_id='task'").run()
  await reconcileWorkbookMutations(db.env, 'task')
  assert.equal(db.balance(), 30)
  assert.equal(db.sqlite.prepare("SELECT status FROM excel_agent_tasks WHERE id='task'").get()!.status, 'failed')
})

for (const field of ['payload_json', 'result_json'] as const) {
  test(`recovery quarantines unknown ${field} and releases only this task's hold`, async t => {
    const { db } = await fixture()
    t.after(() => db.sqlite.close())
    db.sqlite.prepare(`UPDATE excel_agent_mutations SET ${field}='{"kind":"workbook_mutation","schemaVersion":999}',updated_at='2020-01-01' WHERE task_id='task'`).run()
    await reconcileWorkbookMutations(db.env, 'task')
    assert.equal(db.balance(), 30)
    assert.equal(db.sqlite.prepare("SELECT status FROM excel_agent_tasks WHERE id='task'").get()!.status, 'failed')
    assert.equal(db.sqlite.prepare("SELECT state FROM excel_agent_mutations WHERE task_id='task'").get()!.state, 'failed')
    assert.equal(db.count('credit_usage_events'), 0)
  })
}

for (const missing of ['source', 'conversation', 'source-owner', 'source-key'] as const) {
  test(`publication cannot settle or complete after ${missing} changes`, async t => {
    const { db, payload, result } = await fixture()
    t.after(() => db.sqlite.close())
    if (missing === 'source') db.sqlite.prepare("DELETE FROM excel_agent_workbooks WHERE id='source'").run()
    if (missing === 'conversation') db.sqlite.prepare("DELETE FROM excel_agent_conversations WHERE id='conversation'").run()
    if (missing === 'source-owner') db.sqlite.prepare("UPDATE excel_agent_workbooks SET session_id='other' WHERE id='source'").run()
    if (missing === 'source-key') db.sqlite.prepare("UPDATE excel_agent_workbooks SET r2_key='other.xlsx' WHERE id='source'").run()
    await assert.rejects(publishMutationResult(db.env, payload, result), /publication/)
    assert.equal(db.sqlite.prepare("SELECT status FROM credit_reservations WHERE turn_id='turn'").get()!.status, 'reserved')
    assert.equal(db.sqlite.prepare("SELECT status FROM excel_agent_tasks WHERE id='task'").get()!.status, 'processing')
    assert.equal(db.sqlite.prepare("SELECT id FROM excel_agent_workbooks WHERE id='version2'").get(), undefined)
    assert.equal(db.count('credit_usage_events'), 0)
  })
}

test('recovery finalizes an unpublishable staged task without charging edit/artifact costs', async t => {
  const { db } = await fixture()
  t.after(() => db.sqlite.close())
  db.sqlite.prepare("DELETE FROM excel_agent_workbooks WHERE id='source'").run()
  db.sqlite.prepare("UPDATE excel_agent_mutations SET updated_at='2020-01-01' WHERE task_id='task'").run()
  await reconcileWorkbookMutations(db.env, 'task')
  assert.equal(db.sqlite.prepare("SELECT state FROM excel_agent_mutations WHERE task_id='task'").get()!.state, 'failed')
  assert.equal(db.sqlite.prepare("SELECT status FROM excel_agent_tasks WHERE id='task'").get()!.status, 'failed')
  assert.equal(db.balance(), 26)
  assert.equal(db.sqlite.prepare("SELECT COUNT(*) AS n FROM credit_usage_events WHERE action='editWorkbook'").get()!.n, 0)
})

test('recovery retries partial cancellation finalization and releases a hold exactly once', async t => {
  const { db } = await fixture()
  t.after(() => db.sqlite.close())
  db.sqlite.prepare("UPDATE excel_agent_mutations SET state='cancelled',updated_at='2020-01-01' WHERE task_id='task'").run()
  db.failNext('UPDATE credit_reservations')
  const originalError = console.error
  console.error = () => {}
  try { await reconcileWorkbookMutations(db.env, 'task') } finally { console.error = originalError }
  assert.equal(db.sqlite.prepare("SELECT state FROM excel_agent_mutations WHERE task_id='task'").get()!.state, 'cancelled')
  assert.equal(db.sqlite.prepare("SELECT status FROM credit_reservations WHERE turn_id='turn'").get()!.status, 'reserved')
  db.sqlite.prepare("UPDATE excel_agent_mutations SET updated_at='2020-01-01' WHERE task_id='task'").run()
  await reconcileWorkbookMutations(db.env, 'task')
  await reconcileWorkbookMutations(db.env, 'task')
  assert.equal(db.balance(), 30)
  assert.equal(db.sqlite.prepare("SELECT state FROM excel_agent_mutations WHERE task_id='task'").get()!.state, 'failed')
})

test('missing Workflow get is recoverable, and overlapping polls dispatch the same instance only once', async t => {
  const { db } = await fixture()
  t.after(() => db.sqlite.close())
  db.sqlite.prepare("UPDATE excel_agent_mutations SET state='queued',result_json=NULL,updated_at='2020-01-01' WHERE task_id='task'").run()
  const created: string[] = []
  db.env.EXCEL_AGENT_WORKFLOW = {
    get: async () => { throw new Error('Instance not found') },
    create: async ({ id }: { id: string }) => { created.push(id); return { id } }
  } as never
  await Promise.all([reconcileWorkbookMutations(db.env, 'task'), reconcileWorkbookMutations(db.env, 'task')])
  assert.deepEqual(created, ['edit-task'])
  assert.equal(db.sqlite.prepare("SELECT state FROM excel_agent_mutations WHERE task_id='task'").get()!.state, 'queued')
})

test('artifact cleanup fences off staged/completed results and deletes only fixed task output keys', async t => {
  const { db, payload, result } = await fixture()
  t.after(() => db.sqlite.close())
  const deleted: string[][] = []
  db.env.WORKBOOKS = { delete: async (keys: string[]) => { deleted.push(keys) } } as never
  await discardMutationArtifacts(db.env, 'task')
  assert.deepEqual(deleted, [])
  await publishMutationResult(db.env, payload, result)
  await discardMutationArtifacts(db.env, 'task')
  assert.deepEqual(deleted, [])
  db.sqlite.prepare("UPDATE excel_agent_mutations SET state='failing' WHERE task_id='task'").run()
  db.sqlite.prepare("UPDATE excel_agent_tasks SET status='processing' WHERE id='task'").run()
  await discardMutationArtifacts(db.env, 'task')
  assert.deepEqual(deleted, [['agent-demo/results/task/edited.xlsx', 'agent-demo/results/task/context-v2.json.gz']])
  await assert.rejects(discardMutationArtifacts(db.env, '../source'), /identity/)
})

test('recovery stops an active writer before removing outputs; status outages retain artifacts and held credits', async t => {
  const { db } = await fixture()
  t.after(() => db.sqlite.close())
  db.sqlite.prepare("UPDATE excel_agent_tasks SET workflow_instance_id='edit-task' WHERE id='task'").run()
  db.sqlite.prepare("UPDATE excel_agent_mutations SET state='failing' WHERE task_id='task'").run()
  const actions: string[] = []
  db.env.WORKBOOKS = { delete: async () => { actions.push('delete') } } as never
  db.env.EXCEL_AGENT_WORKFLOW = { get: async () => ({ status: async () => ({ status: 'running' }), terminate: async () => { actions.push('terminate') } }) } as never
  await discardMutationArtifacts(db.env, 'task', true)
  assert.deepEqual(actions, ['terminate', 'delete'])
  db.env.EXCEL_AGENT_WORKFLOW = { get: async () => { throw new Error('status outage') } } as never
  await assert.rejects(discardMutationArtifacts(db.env, 'task', true), /status outage/)
  assert.equal(actions.length, 2)
  assert.equal(db.sqlite.prepare("SELECT status FROM credit_reservations WHERE turn_id='turn'").get()!.status, 'reserved')
})

test('recovery completes a partial failure after the task row is already terminal', async t => {
  const { db } = await fixture()
  t.after(() => db.sqlite.close())
  db.sqlite.prepare("UPDATE excel_agent_tasks SET status='failed' WHERE id='task'").run()
  db.sqlite.prepare("UPDATE excel_agent_mutations SET state='cancelled',updated_at='2020-01-01' WHERE task_id='task'").run()
  await reconcileWorkbookMutations(db.env, 'task')
  assert.equal(db.balance(), 30)
  assert.equal(db.sqlite.prepare("SELECT state FROM excel_agent_mutations WHERE task_id='task'").get()!.state, 'failed')
})
