// Serializes continuous brush clicks, including their history groups. Save/Undo wait
// for the last click so they cannot land halfway through a fill + text update.
let pending: Promise<unknown> = Promise.resolve()
export function queueFormatEdit<T>(edit: () => Promise<T>): Promise<T> {
  const result = pending.then(edit)
  pending = result.catch(() => undefined)
  return result
}
export async function flushFormatEdits(): Promise<void> {
  await pending
}
