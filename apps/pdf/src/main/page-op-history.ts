/**
 * Undo/redo for in-place page rewrites (Crop, Page Size). Those operations
 * write the file and reload it, so they cannot live in the renderer's pending
 * EditSnapshot (nothing could display a pending CropBox or media-box resize).
 * Instead the main process keeps the bytes on either side of each rewrite and
 * hands the renderer a token; undo/redo swap the file back, refusing when the
 * file on disk is not the version the step expects (changed elsewhere).
 */

export type PageOpDirection = 'undo' | 'redo'

interface Entry {
  readonly before: Uint8Array
  readonly after: Uint8Array
}

/** keep at most this many steps per file and this many bytes overall */
const MAX_STEPS_PER_PATH = 5
const MAX_TOTAL_BYTES = 256 * 1024 * 1024

export class PageOpHistory {
  private readonly byPath = new Map<string, Map<string, Entry>>()
  private next = 1

  record(path: string, before: Uint8Array, after: Uint8Array): string | null {
    if (before.byteLength + after.byteLength > MAX_TOTAL_BYTES) return null
    const token = `pageop-${this.next++}`
    const steps = this.byPath.get(path) ?? new Map<string, Entry>()
    steps.set(token, { before, after })
    while (steps.size > MAX_STEPS_PER_PATH) steps.delete(steps.keys().next().value!)
    this.byPath.set(path, steps)
    this.evict()
    return token
  }

  /**
   * The bytes to write for `direction`, or an error when the step is unknown or
   * the file no longer holds the version this step starts from.
   */
  resolve(
    path: string,
    token: string,
    direction: PageOpDirection,
    current: Uint8Array,
  ): { ok: true; bytes: Uint8Array } | { ok: false; error: string } {
    const entry = this.byPath.get(path)?.get(token)
    if (!entry) return { ok: false, error: 'pdf: this page operation can no longer be undone' }
    const expected = direction === 'undo' ? entry.after : entry.before
    if (!sameBytes(current, expected)) {
      return { ok: false, error: 'pdf: the file changed on disk since this page operation' }
    }
    return { ok: true, bytes: direction === 'undo' ? entry.before : entry.after }
  }

  forget(path: string): void {
    this.byPath.delete(path)
  }

  private evict(): void {
    let total = 0
    for (const steps of this.byPath.values())
      for (const e of steps.values()) total += e.before.byteLength + e.after.byteLength
    for (const steps of this.byPath.values()) {
      for (const key of [...steps.keys()]) {
        if (total <= MAX_TOTAL_BYTES) return
        const e = steps.get(key)!
        total -= e.before.byteLength + e.after.byteLength
        steps.delete(key)
      }
    }
  }
}

export function sameBytes(a: Uint8Array, b: Uint8Array): boolean {
  if (a.byteLength !== b.byteLength) return false
  for (let i = 0; i < a.byteLength; i++) if (a[i] !== b[i]) return false
  return true
}
