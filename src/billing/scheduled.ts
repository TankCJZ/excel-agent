import { releaseExpiredReservationsForUser } from '#agent/billing/billingRepository'
import type { AgentWorkerEnv } from '#agent/types'
import { reconcileWorkbookMutations } from '#agent/workflows/mutationRecovery'

export async function releaseExpiredCredits(_controller: ScheduledController, env: AgentWorkerEnv) {
  await reconcileWorkbookMutations(env)
  await releaseExpiredReservationsForUser(env)
}
