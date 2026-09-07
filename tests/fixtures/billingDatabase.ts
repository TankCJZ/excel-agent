import { DatabaseSync } from 'node:sqlite'
import { readFileSync, readdirSync } from 'node:fs'
import { resolve } from 'node:path'
import type { AgentWorkerEnv } from '../../src/types.ts'

export function createBillingDatabase(balance = 30, subscription = 0) {
  const sqlite = new DatabaseSync(':memory:')
  const migrations = resolve(import.meta.dirname, '../../migrations')
  for (const file of readdirSync(migrations).filter(file => file.endsWith('.sql')).sort()) {
    sqlite.exec(readFileSync(resolve(migrations, file), 'utf8'))
  }
  sqlite.exec('PRAGMA foreign_keys = ON')
  sqlite.prepare(`INSERT INTO users
    (id, email, password_hash, credits_balance, free_credits_balance, subscription_credits_balance, created_at, updated_at)
    VALUES ('billing-test', 'billing@example.invalid', 'not-a-password', ?, ?, ?, '2026-09-02', '2026-09-02')
  `).run(balance, balance - subscription, subscription)

  let failPattern: string | null = null
  const DB = {
    prepare(sql: string) {
      let values: Array<string | number | null> = []
      return {
        bind(...input: typeof values) { values = input; return this },
        async first() { return sqlite.prepare(sql).get(...values) || null },
        async all() { return { success: true, results: sqlite.prepare(sql).all(...values) } },
        async raw() { return sqlite.prepare(sql).all(...values).map(row => Object.values(row)) },
        execute() {
          if (failPattern && sql.includes(failPattern)) {
            failPattern = null
            throw new Error('Injected database failure')
          }
          if (/^\s*SELECT\b/i.test(sql)) return { success: true, meta: { changes: 0 }, results: sqlite.prepare(sql).all(...values) }
          return { success: true, meta: { changes: Number(sqlite.prepare(sql).run(...values).changes) }, results: [] }
        },
        async run() { return this.execute() }
      }
    },
    async batch(statements: Array<{ execute: () => unknown }>) {
      sqlite.exec('BEGIN')
      try {
        const result = statements.map(statement => statement.execute())
        sqlite.exec('COMMIT')
        return result
      } catch (error) {
        sqlite.exec('ROLLBACK')
        throw error
      }
    }
  }
  return {
    sqlite,
    env: { DB } as unknown as AgentWorkerEnv,
    failNext: (pattern: string) => { failPattern = pattern },
    balance: () => sqlite.prepare("SELECT credits_balance FROM users WHERE id = 'billing-test'").get()!.credits_balance as number,
    count: (table: 'credit_reservations' | 'credit_ledger' | 'credit_usage_events') => Number(sqlite.prepare(`SELECT COUNT(*) AS count FROM ${table}`).get()!.count)
  }
}
