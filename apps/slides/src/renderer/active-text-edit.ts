let active: (() => void | Promise<void>) | null = null
let pending: Promise<void> | null = null
export function registerTextCommit(commit: () => void | Promise<void>): () => void {
  active = commit
  return () => {
    if (active === commit) active = null
  }
}
export async function flushTextCommit(): Promise<void> {
  if (!pending) {
    const commit = active
    const work = Promise.resolve().then(() => commit?.())
    pending = work
    void work
      .finally(() => {
        if (pending === work) pending = null
      })
      .catch(() => undefined)
  }
  await pending
}
