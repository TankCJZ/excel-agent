export type CreditTurnSummary = {
  reservationId: string
  turnId: string
  estimatedCredits: number
  actualCredits?: number
  refundedCredits: number
  status: 'reserved' | 'settled' | 'released'
}

export function creditDisplayFields(summary: CreditTurnSummary) {
  return {
    estimatedCredits: summary.estimatedCredits,
    actualCredits: summary.status === 'reserved' ? undefined : summary.actualCredits ?? 0,
    refundedCredits: summary.status === 'reserved' ? 0 : summary.refundedCredits
  }
}
