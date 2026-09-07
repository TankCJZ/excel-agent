import assert from 'node:assert/strict'
import test from 'node:test'
import { createBillingDatabase } from './fixtures/billingDatabase.ts'
import { reserveTurnCredits, ensureTurnCreditReservation, releaseTurnCredits } from '../src/billing/billingRepository.ts'

test('validated edits top up underestimated holds atomically and concurrent calls do not double debit', async t => {
  const db = createBillingDatabase(30, 10)
  t.after(() => db.sqlite.close())
  const hold = await reserveTurnCredits(db.env, { userId: 'billing-test', conversationId: 'c', turnId: 't', estimatedCredits: 4 })
  await Promise.all([ensureTurnCreditReservation(db.env, hold.id, 14), ensureTurnCreditReservation(db.env, hold.id, 14)])
  assert.equal(db.balance(), 16)
  const reservation = db.sqlite.prepare('SELECT estimated_credits,held_subscription_credits,held_free_credits FROM credit_reservations WHERE id=?').get(hold.id)!
  assert.deepEqual({ ...reservation }, { estimated_credits: 14, held_subscription_credits: 10, held_free_credits: 4 })
  assert.equal(db.count('credit_ledger'), 2)
  await releaseTurnCredits(db.env, hold.id, 'test cancellation')
  assert.equal(db.balance(), 30)
})

test('insufficient balance or transaction failure does not partially change a hold', async t => {
  const db = createBillingDatabase(8)
  t.after(() => db.sqlite.close())
  const hold = await reserveTurnCredits(db.env, { userId: 'billing-test', conversationId: 'c', turnId: 't', estimatedCredits: 4 })
  await assert.rejects(ensureTurnCreditReservation(db.env, hold.id, 14), /Not enough credits/)
  assert.equal(db.balance(), 4)
  db.failNext('UPDATE credit_reservations')
  await assert.rejects(ensureTurnCreditReservation(db.env, hold.id, 6), /Injected/)
  assert.equal(db.balance(), 4)
  assert.equal(db.count('credit_ledger'), 1)
  assert.equal(db.sqlite.prepare('SELECT estimated_credits FROM credit_reservations WHERE id=?').get(hold.id)!.estimated_credits, 4)
  await ensureTurnCreditReservation(db.env, hold.id, 6)
  assert.equal(db.balance(), 2)
})
