import {
  CREDIT_POLICY_VERSION,
  allocateCreditUsage,
  getPlanEntitlements,
  normalizePlanKey,
  type PlanKey
} from '#shared/billing/planCatalog'
import type { CreditTurnSummary } from '#shared/billing/creditSummary'
import { HttpError } from '#agent/http/responses'
import type { AgentWorkerEnv } from '#agent/types'

const ACTIVE_SUBSCRIPTION_STATUSES = ['active', 'trialing', 'past_due', 'incomplete']
const RESERVATION_TTL_MS = 60 * 60 * 1000

type BillingAccountRow = {
  id: string
  credits_balance: number
  free_credits_balance: number
  subscription_credits_balance: number
  free_xlsx_download_task_id: string | null
  plan_key: string | null
}

type CreditReservationRow = {
  id: string
  user_id: string
  conversation_id: string
  turn_id: string
  task_id: string | null
  estimated_credits: number
  held_free_credits: number
  held_subscription_credits: number
  settled_credits: number
  status: 'reserved' | 'settled' | 'released'
  expires_at: string
}

export type BillingAccess = {
  planKey: PlanKey
  credits: {
    total: number
    free: number
    subscription: number
  }
  freeXlsxDownloadTaskId: string | null
  entitlements: ReturnType<typeof getPlanEntitlements>
}

export type CreditReservation = {
  id: string
  turnId: string
  estimatedCredits: number
  status: CreditReservationRow['status']
  balanceAfterHold: number
}

export async function getBillingAccess(env: AgentWorkerEnv, userId: string): Promise<BillingAccess> {
  const row = await findBillingAccount(env, userId)
  if (!row) throw new HttpError(401, 'UNAUTHORIZED', 'The billing account was not found')
  const planKey = normalizePlanKey(row.plan_key)
  return {
    planKey,
    credits: {
      total: row.credits_balance,
      free: row.free_credits_balance,
      subscription: row.subscription_credits_balance
    },
    freeXlsxDownloadTaskId: row.free_xlsx_download_task_id,
    entitlements: getPlanEntitlements(planKey)
  }
}

export async function reserveTurnCredits(
  env: AgentWorkerEnv,
  input: {
    userId: string
    conversationId: string
    turnId: string
    estimatedCredits: number
    rejectExisting?: boolean
  }
): Promise<CreditReservation> {
  await releaseExpiredReservationsForUser(env, input.userId)
  const existing = await findReservationByTurn(env, input.turnId)
  if (existing) {
    if (existing.user_id !== input.userId || existing.conversation_id !== input.conversationId) throw new HttpError(409, 'CONFLICT', 'Turn ID is already in use')
    if (existing.status !== 'reserved') {
      throw new HttpError(409, 'TURN_ALREADY_FINALIZED', 'This turn has already been finalized')
    }
    if (input.rejectExisting) throw new HttpError(409, 'TURN_IN_PROGRESS', 'This turn is already running')
    const account = await getBillingAccess(env, input.userId)
    return mapReservation(existing, account.credits.total)
  }

  const estimatedCredits = Math.max(1, Math.floor(input.estimatedCredits))
  const account = await getBillingAccess(env, input.userId)
  if (account.credits.total < estimatedCredits) {
    throw new HttpError(402, 'INSUFFICIENT_CREDITS', 'Not enough credits to start this request', {
      required: estimatedCredits,
      available: account.credits.total,
      shortfall: estimatedCredits - account.credits.total
    })
  }

  const heldSubscriptionCredits = Math.min(account.credits.subscription, estimatedCredits)
  const heldFreeCredits = estimatedCredits - heldSubscriptionCredits
  if (heldFreeCredits > account.credits.free) {
    throw new HttpError(402, 'INSUFFICIENT_CREDITS', 'Not enough credits to start this request')
  }

  const id = crypto.randomUUID()
  const now = new Date().toISOString()
  const expiresAt = new Date(Date.now() + RESERVATION_TTL_MS).toISOString()
  const results = await env.DB.batch([
    env.DB.prepare(`
      INSERT INTO credit_reservations (
        id, user_id, conversation_id, turn_id, policy_version, estimated_credits,
        held_free_credits, held_subscription_credits, settled_credits, status,
        expires_at, created_at, updated_at
      ) SELECT ?, ?, ?, ?, ?, ?, ?, ?, 0, 'reserved', ?, ?, ?
      FROM users WHERE id = ? AND credits_balance >= ?
        AND free_credits_balance >= ? AND subscription_credits_balance >= ?
      ON CONFLICT(turn_id) DO NOTHING
    `).bind(
      id,
      input.userId,
      input.conversationId,
      input.turnId,
      CREDIT_POLICY_VERSION,
      estimatedCredits,
      heldFreeCredits,
      heldSubscriptionCredits,
      expiresAt,
      now,
      now,
      input.userId,
      estimatedCredits,
      heldFreeCredits,
      heldSubscriptionCredits
    ),
    env.DB.prepare(`
      UPDATE users
      SET credits_balance = credits_balance - ?,
          free_credits_balance = free_credits_balance - ?,
          subscription_credits_balance = subscription_credits_balance - ?,
          updated_at = ?
      WHERE id = ?
        AND credits_balance >= ?
        AND free_credits_balance >= ?
        AND subscription_credits_balance >= ?
        AND EXISTS (SELECT 1 FROM credit_reservations WHERE id = ? AND status = 'reserved')
    `).bind(
      estimatedCredits,
      heldFreeCredits,
      heldSubscriptionCredits,
      now,
      input.userId,
      estimatedCredits,
      heldFreeCredits,
      heldSubscriptionCredits,
      id
    ),
    ledgerStatement(env, {
      reservationId: id,
      userId: input.userId,
      type: 'agent_credit_reserved',
      amount: -estimatedCredits,
      idempotencyKey: `agent-credit-reserved:${input.turnId}`,
      note: `Reserved up to ${estimatedCredits} credits for Agent turn ${input.turnId}`,
      now
    }),
    env.DB.prepare('SELECT credits_balance FROM users WHERE id = ?').bind(input.userId)
  ])

  if (!changed(results[0]) || !changed(results[1])) {
    const concurrent = await findReservationByTurn(env, input.turnId)
    if (concurrent) throw new HttpError(409, 'TURN_IN_PROGRESS', 'This turn is already running')
    const current = await getBillingAccess(env, input.userId)
    throw new HttpError(402, 'INSUFFICIENT_CREDITS', 'Credits changed before this request could be reserved', {
      required: estimatedCredits,
      available: current.credits.total
    })
  }

  const balanceAfter = (results[3].results[0] as { credits_balance: number }).credits_balance

  return {
    id,
    turnId: input.turnId,
    estimatedCredits,
    status: 'reserved',
    balanceAfterHold: balanceAfter
  }
}

// Upgrade a heuristic reservation to the validated edit cost BEFORE dispatch.
// The balance debit, reservation change and ledger entry are one transaction.
export async function ensureTurnCreditReservation(env: AgentWorkerEnv, reservationId: string, requiredCredits: number): Promise<CreditReservation> {
  if (!Number.isSafeInteger(requiredCredits) || requiredCredits < 1) throw new Error('Invalid credit requirement')
  for (let attempt = 0; attempt < 3; attempt++) {
    const reservation = await findReservationById(env, reservationId)
    if (!reservation || reservation.status !== 'reserved') throw new HttpError(409, 'TURN_ALREADY_FINALIZED', 'This turn is no longer active')
    const account = await getBillingAccess(env, reservation.user_id)
    const held = reservation.held_free_credits + reservation.held_subscription_credits
    if (held >= requiredCredits) return mapReservation(reservation, account.credits.total)
    const delta = requiredCredits - held
    if (account.credits.total < delta) throw new HttpError(402, 'INSUFFICIENT_CREDITS', 'Not enough credits to execute the validated workbook edit', { required: delta, available: account.credits.total })
    const subscription = Math.min(delta, account.credits.subscription)
    const free = delta - subscription
    const now = new Date().toISOString()
    const results = await env.DB.batch([
      env.DB.prepare(`UPDATE users SET credits_balance=credits_balance-?, subscription_credits_balance=subscription_credits_balance-?, free_credits_balance=free_credits_balance-?, updated_at=?
        WHERE id=? AND credits_balance>=? AND subscription_credits_balance>=? AND free_credits_balance>=?
        AND EXISTS (SELECT 1 FROM credit_reservations WHERE id=? AND status='reserved' AND estimated_credits=? AND held_free_credits=? AND held_subscription_credits=?)`)
        .bind(delta, subscription, free, now, reservation.user_id, delta, subscription, free, reservationId, reservation.estimated_credits, reservation.held_free_credits, reservation.held_subscription_credits),
      env.DB.prepare(`UPDATE credit_reservations SET estimated_credits=?, held_free_credits=held_free_credits+?, held_subscription_credits=held_subscription_credits+?, updated_at=?
        WHERE id=? AND changes()=1`)
        .bind(requiredCredits, free, subscription, now, reservationId),
      env.DB.prepare(`INSERT INTO credit_ledger (id,user_id,task_id,type,amount,balance_after,idempotency_key,note,created_at)
        SELECT ?,id,NULL,'agent_credit_reserved',-?,credits_balance,?,?,? FROM users WHERE id=? AND changes()=1`)
        .bind(crypto.randomUUID(), delta, `agent-credit-topup:${reservationId}:${requiredCredits}`, `Additional hold for validated workbook edit (${requiredCredits} credits total)`, now, reservation.user_id)
    ])
    if (changed(results[0]!) && changed(results[1]!)) {
      const current = await findReservationById(env, reservationId)
      return mapReservation(current!, (await getBillingAccess(env, reservation.user_id)).credits.total)
    }
  }
  throw new HttpError(409, 'CONFLICT', 'Credits changed while reserving the edit. Try again')
}

export async function attachReservationTask(env: AgentWorkerEnv, reservationId: string, taskId: string) {
  const result = await env.DB.prepare(`
    UPDATE credit_reservations SET task_id = ?, updated_at = ?
    WHERE id = ? AND status = 'reserved' AND (task_id IS NULL OR task_id = ?)
  `).bind(taskId, new Date().toISOString(), reservationId, taskId).run()
  if (changed(result)) return
  const reservation = await findReservationById(env, reservationId)
  if (!reservation || reservation.status !== 'reserved' || reservation.task_id !== taskId) {
    throw new HttpError(409, 'CREDIT_RESERVATION_UNAVAILABLE', 'The credit reservation is no longer available')
  }
}

export async function settleTurnCredits(
  env: AgentWorkerEnv,
  input: {
    reservationId: string
    actualCredits: number
    taskId?: string | null
    actions?: Array<{ action: string, credits: number, idempotencyKey: string }>
    atomicCommit?: {
      taskId: string
      statements: (finalizationKey: string) => D1PreparedStatement[]
    }
  }
) {
  const reservation = await findReservationById(env, input.reservationId)
  if (!reservation) throw new HttpError(404, 'NOT_FOUND', 'Credit reservation was not found')
  if (reservation.status !== 'reserved') return summaryWithBalance(env, reservation)

  const heldCredits = reservation.held_free_credits + reservation.held_subscription_credits
  if (input.atomicCommit && input.actualCredits > heldCredits) throw new HttpError(409, 'INSUFFICIENT_CREDIT_RESERVATION', 'Workbook edit cost was not fully reserved before execution')
  const actualCredits = Math.max(0, Math.min(heldCredits, Math.floor(input.actualCredits)))
  const consumedSubscription = Math.min(reservation.held_subscription_credits, actualCredits)
  const consumedFree = actualCredits - consumedSubscription
  const refundSubscription = reservation.held_subscription_credits - consumedSubscription
  const refundFree = reservation.held_free_credits - consumedFree
  const refundCredits = refundSubscription + refundFree
  const now = new Date().toISOString()
  const finalizationKey = crypto.randomUUID()

  const statements: D1PreparedStatement[] = [env.DB.prepare(`
    UPDATE credit_reservations
    SET task_id = COALESCE(?, task_id), settled_credits = ?, status = 'settled',
        updated_at = ?, settled_at = ?, finalization_key = ?
    WHERE id = ? AND status = 'reserved'
      ${input.atomicCommit ? `AND EXISTS (
        SELECT 1 FROM excel_agent_mutations m JOIN excel_agent_tasks t ON t.id = m.task_id
        JOIN excel_agent_conversations c ON c.id = t.thread_id AND c.user_id = t.user_id AND c.session_id = t.session_id
        JOIN excel_agent_workbooks w ON w.id = json_extract(m.payload_json, '$.plan.sourceWorkbookId')
          AND w.user_id = t.user_id AND w.session_id = t.session_id AND w.r2_key = t.input_r2_key
          WHERE m.task_id = ? AND m.state = 'staged' AND t.status = 'processing'
            AND NOT EXISTS (SELECT 1 FROM excel_agent_conversation_tombstones d WHERE d.conversation_id=c.id)
          AND t.user_id = credit_reservations.user_id AND t.id = credit_reservations.task_id
      )` : ''}
  `).bind(input.taskId || null, actualCredits, now, now, finalizationKey, reservation.id, ...(input.atomicCommit ? [input.atomicCommit.taskId] : []))]
  if (refundCredits > 0) {
    statements.push(env.DB.prepare(`
      UPDATE users
      SET credits_balance = credits_balance + ?,
          free_credits_balance = free_credits_balance + ?,
          subscription_credits_balance = subscription_credits_balance + ?,
          updated_at = ?
      WHERE id = ?
        AND EXISTS (
          SELECT 1 FROM credit_reservations
          WHERE id = ? AND status = 'settled' AND finalization_key = ?
        )
    `).bind(refundCredits, refundFree, refundSubscription, now, reservation.user_id, reservation.id, finalizationKey))
  }
  const actions = allocateCreditUsage(input.actions || [{
    action: 'response', credits: actualCredits, idempotencyKey: `credit-usage:${reservation.turn_id}:response`
  }], actualCredits)
  if (actions.reduce((sum, action) => sum + action.credits, 0) !== actualCredits) {
    throw new HttpError(500, 'CREDIT_USAGE_MISMATCH', 'Credit usage does not match the settlement')
  }
  for (const action of actions) {
    statements.push(env.DB.prepare(`
      INSERT INTO credit_usage_events (
        id, reservation_id, user_id, turn_id, task_id, action, credits,
        idempotency_key, created_at, metadata_json
      ) SELECT ?, ?, ?, ?, ?, ?, ?, ?, ?, ?
      WHERE EXISTS (
        SELECT 1 FROM credit_reservations
        WHERE id = ? AND status = 'settled' AND finalization_key = ?
      )
      ON CONFLICT(idempotency_key) DO NOTHING
    `).bind(
      crypto.randomUUID(),
      reservation.id,
      reservation.user_id,
      reservation.turn_id,
      input.taskId || reservation.task_id,
      action.action,
      action.credits,
      action.idempotencyKey,
      now,
      JSON.stringify({ nominalCredits: action.nominalCredits, waivedCredits: action.nominalCredits - action.credits }),
      reservation.id,
      finalizationKey
    ))
  }
  if (refundCredits > 0) {
    statements.push(ledgerStatement(env, {
      reservationId: reservation.id, finalizationKey, userId: reservation.user_id,
      type: 'agent_credit_refund', amount: refundCredits,
      idempotencyKey: `agent-credit-refund:${reservation.turn_id}`,
      note: `Refunded ${refundCredits} unused credits for Agent turn ${reservation.turn_id}`, now
    }))
  }
  if (input.atomicCommit) statements.push(...input.atomicCommit.statements(finalizationKey))
  const results = await env.DB.batch(statements)
  if (!changed(results[0])) {
    const current = await findReservationById(env, reservation.id)
    if (!current) throw new HttpError(404, 'NOT_FOUND', 'Credit reservation was not found')
    return summaryWithBalance(env, current)
  }

  const account = await getBillingAccess(env, reservation.user_id)
  return {
    reservationId: reservation.id,
    turnId: reservation.turn_id,
    estimatedCredits: reservation.estimated_credits,
    actualCredits,
    refundedCredits: refundCredits,
    balance: account.credits.total,
    status: 'settled' as const
  }
}

export async function releaseTurnCredits(env: AgentWorkerEnv, reservationId: string, reason: string) {
  const reservation = await findReservationById(env, reservationId)
  if (!reservation) return null
  if (reservation.status !== 'reserved') return summaryWithBalance(env, reservation)

  const refundCredits = reservation.held_free_credits + reservation.held_subscription_credits
  const now = new Date().toISOString()
  const finalizationKey = crypto.randomUUID()
  const results = await env.DB.batch([
    env.DB.prepare(`
      UPDATE credit_reservations
      SET status = 'released', updated_at = ?, released_at = ?, finalization_key = ?
      WHERE id = ? AND status = 'reserved'
    `).bind(now, now, finalizationKey, reservation.id),
    env.DB.prepare(`
      UPDATE users
      SET credits_balance = credits_balance + ?,
          free_credits_balance = free_credits_balance + ?,
          subscription_credits_balance = subscription_credits_balance + ?,
          updated_at = ?
      WHERE id = ?
        AND EXISTS (
          SELECT 1 FROM credit_reservations
          WHERE id = ? AND status = 'released' AND finalization_key = ?
        )
    `).bind(
      refundCredits,
      reservation.held_free_credits,
      reservation.held_subscription_credits,
      now,
      reservation.user_id,
      reservation.id,
      finalizationKey
    ),
    ledgerStatement(env, {
      reservationId: reservation.id, finalizationKey, userId: reservation.user_id,
      type: 'agent_credit_released', amount: refundCredits,
      idempotencyKey: `agent-credit-released:${reservation.turn_id}`,
      note: `${reason.slice(0, 240)} (${reservation.turn_id})`, now
    })
  ])
  if (!changed(results[0])) {
    const current = await findReservationById(env, reservation.id)
    return current ? summaryWithBalance(env, current) : null
  }
  const account = await getBillingAccess(env, reservation.user_id)
  return {
    reservationId: reservation.id,
    turnId: reservation.turn_id,
    estimatedCredits: reservation.estimated_credits,
    actualCredits: 0,
    refundedCredits: refundCredits,
    balance: account.credits.total,
    status: 'released' as const
  }
}

export async function assertConcurrentTaskCapacity(env: AgentWorkerEnv, userId: string, access: BillingAccess) {
  const row = await env.DB.prepare(`
    SELECT COUNT(*) AS count FROM excel_agent_tasks
    WHERE user_id = ? AND status IN ('planning', 'queued', 'processing')
  `).bind(userId).first<{ count: number }>()
  if (Number(row?.count || 0) >= access.entitlements.maxConcurrentTasks) {
    throw new HttpError(429, 'CONCURRENT_TASK_LIMIT', 'The current plan task limit has been reached', {
      limit: access.entitlements.maxConcurrentTasks,
      plan: access.planKey
    })
  }
}

export async function assertFreeActionAvailable(
  env: AgentWorkerEnv,
  userId: string,
  action: 'searchWeb' | 'createChart'
) {
  const access = await getBillingAccess(env, userId)
  const limit = action === 'searchWeb'
    ? access.entitlements.freeResearchRuns
    : access.entitlements.freeChartRuns
  if (limit === null) return
  const row = await env.DB.prepare(`
    SELECT COUNT(*) AS count FROM credit_usage_events WHERE user_id = ? AND action = ?
  `).bind(userId, action).first<{ count: number }>()
  if (Number(row?.count || 0) >= limit) {
    throw new HttpError(403, 'SUBSCRIPTION_REQUIRED', 'This free plan trial has already been used', {
      feature: action,
      plan: access.planKey
    })
  }
}

export async function authorizeWorkbookDownload(
  env: AgentWorkerEnv,
  input: { userId: string, taskId: string, completedAt: string | null }
) {
  const access = await getBillingAccess(env, input.userId)
  if (input.completedAt) {
    const expiresAt = Date.parse(input.completedAt) + access.entitlements.artifactRetentionDays * 86_400_000
    if (Date.now() > expiresAt) {
      throw new HttpError(410, 'ARTIFACT_EXPIRED', 'This workbook result has expired', {
        retentionDays: access.entitlements.artifactRetentionDays,
        plan: access.planKey
      })
    }
  }
  if (access.planKey !== 'free') return access
  if (access.freeXlsxDownloadTaskId === input.taskId) return access
  if (access.freeXlsxDownloadTaskId) {
    throw new HttpError(403, 'SUBSCRIPTION_REQUIRED', 'Upgrade to download another XLSX workbook', {
      feature: 'xlsx_download'
    })
  }

  const result = await env.DB.prepare(`
    UPDATE users SET free_xlsx_download_task_id = ?, updated_at = ?
    WHERE id = ? AND free_xlsx_download_task_id IS NULL
  `).bind(input.taskId, new Date().toISOString(), input.userId).run()
  if (!changed(result)) {
    const refreshed = await getBillingAccess(env, input.userId)
    if (refreshed.freeXlsxDownloadTaskId !== input.taskId) {
      throw new HttpError(403, 'SUBSCRIPTION_REQUIRED', 'Upgrade to download another XLSX workbook')
    }
  }
  return getBillingAccess(env, input.userId)
}

async function findBillingAccount(env: AgentWorkerEnv, userId: string) {
  const statusPlaceholders = ACTIVE_SUBSCRIPTION_STATUSES.map(() => '?').join(', ')
  return env.DB.prepare(`
    SELECT u.id, u.credits_balance, u.free_credits_balance, u.subscription_credits_balance,
           u.free_xlsx_download_task_id,
           (
             SELECT s.plan_key FROM subscriptions s
             WHERE s.user_id = u.id AND s.status IN (${statusPlaceholders})
             ORDER BY s.updated_at DESC LIMIT 1
           ) AS plan_key
    FROM users u WHERE u.id = ? LIMIT 1
  `).bind(...ACTIVE_SUBSCRIPTION_STATUSES, userId).first<BillingAccountRow>()
}

async function findReservationByTurn(env: AgentWorkerEnv, turnId: string) {
  return env.DB.prepare('SELECT * FROM credit_reservations WHERE turn_id = ? LIMIT 1')
    .bind(turnId)
    .first<CreditReservationRow>()
}

async function findReservationById(env: AgentWorkerEnv, id: string) {
  return env.DB.prepare('SELECT * FROM credit_reservations WHERE id = ? LIMIT 1')
    .bind(id)
    .first<CreditReservationRow>()
}

export async function releaseExpiredReservationsForUser(env: AgentWorkerEnv, userId?: string) {
  const expired = await env.DB.prepare(`
    SELECT r.id FROM credit_reservations r
    WHERE (? IS NULL OR r.user_id = ?) AND r.status = 'reserved' AND r.expires_at <= ?
      AND NOT EXISTS (
        SELECT 1 FROM excel_agent_tasks t WHERE t.id = r.task_id
          AND t.status IN ('planning', 'queued', 'processing')
      )
    ORDER BY r.expires_at ASC LIMIT 100
  `).bind(userId || null, userId || null, new Date().toISOString()).all<{ id: string }>()
  for (const row of expired.results || []) {
    await releaseTurnCredits(env, row.id, 'Expired Agent credit reservation released')
  }
}

function ledgerStatement(
  env: AgentWorkerEnv,
  input: {
    userId: string
    type: string
    amount: number
    reservationId: string
    finalizationKey?: string
    idempotencyKey: string
    note: string
    now: string
  }
) {
  return env.DB.prepare(`
    INSERT INTO credit_ledger (
      id, user_id, task_id, type, amount, balance_after, idempotency_key, note, created_at
    ) SELECT ?, u.id, NULL, ?, ?, u.credits_balance, ?, ?, ?
    FROM users u JOIN credit_reservations r ON r.user_id = u.id
    WHERE u.id = ? AND r.id = ?
      AND ${input.finalizationKey ? 'r.finalization_key = ?' : "r.status = 'reserved'"}
    ON CONFLICT(idempotency_key) DO NOTHING
  `).bind(
    crypto.randomUUID(),
    input.type,
    input.amount,
    input.idempotencyKey,
    input.note,
    input.now,
    input.userId,
    input.reservationId,
    ...(input.finalizationKey ? [input.finalizationKey] : [])
  )
}

function changed(result: D1Result<unknown>) {
  return Number(result.meta?.changes || 0) > 0
}

function mapReservation(row: CreditReservationRow, balance: number): CreditReservation {
  return {
    id: row.id,
    turnId: row.turn_id,
    estimatedCredits: row.estimated_credits,
    status: row.status,
    balanceAfterHold: balance
  }
}

function reservationSummary(row: CreditReservationRow): CreditTurnSummary {
  return {
    reservationId: row.id,
    turnId: row.turn_id,
    estimatedCredits: row.estimated_credits,
    actualCredits: row.status === 'reserved' ? undefined : row.settled_credits,
    refundedCredits: row.status === 'reserved' ? 0 : Math.max(0, row.estimated_credits - row.settled_credits),
    status: row.status
  }
}

async function summaryWithBalance(env: AgentWorkerEnv, row: CreditReservationRow) {
  const account = await getBillingAccess(env, row.user_id)
  return { ...reservationSummary(row), balance: account.credits.total }
}

export async function getTurnCredits(env: AgentWorkerEnv, userId: string, turnId: string) {
  const row = await findReservationByTurn(env, turnId)
  return row?.user_id === userId ? summaryWithBalance(env, row) : null
}

export async function getTaskCredits(env: AgentWorkerEnv, userId: string, taskId: string) {
  const row = await env.DB.prepare('SELECT * FROM credit_reservations WHERE user_id = ? AND task_id = ? LIMIT 1')
    .bind(userId, taskId).first<CreditReservationRow>()
  return row ? summaryWithBalance(env, row) : null
}

export async function listConversationCredits(env: AgentWorkerEnv, userId: string, conversationId: string) {
  const rows = await env.DB.prepare('SELECT * FROM credit_reservations WHERE user_id = ? AND conversation_id = ?')
    .bind(userId, conversationId).all<CreditReservationRow>()
  return (rows.results || []).map(reservationSummary)
}
