import path from 'node:path'
import { fileURLToPath } from 'node:url'

export async function delay(options: { seconds?: number, milliseconds?: number } = {}) {
  const duration = typeof options.seconds === 'number'
    ? options.seconds * 1000
    : options.milliseconds

  if (typeof duration !== 'number') {
    throw new TypeError('Expected seconds or milliseconds')
  }

  await new Promise((resolve) => setTimeout(resolve, duration))
}

export function toPath(value: string | URL) {
  return value instanceof URL ? fileURLToPath(value.toString()) : value
}

export function* traversePathUp(startPath: string | URL) {
  let directory = path.resolve(toPath(startPath))

  while (true) {
    yield directory
    const parent = path.dirname(directory)
    if (parent === directory) {
      return
    }
    directory = parent
  }
}
