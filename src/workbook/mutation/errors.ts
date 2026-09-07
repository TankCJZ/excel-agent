export class WorkbookMutationError extends Error {
  readonly code: string
  constructor(code: string, message: string) {
    super(message)
    this.name = 'WorkbookMutationError'
    this.code = code
  }
}

export function requireMutation(condition: unknown, code: string, message: string): asserts condition {
  if (!condition) throw new WorkbookMutationError(code, message)
}
