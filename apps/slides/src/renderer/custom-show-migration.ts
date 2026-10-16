import type { CustomShow } from './slideshow-utils'

export interface LegacyShowStorage {
  getItem(key: string): string | null
  removeItem(key: string): void
}

export interface PendingLegacyCustomShows {
  key: string
  raw: string
  /** Set only by an explicit user edit deleting every migrated show. */
  allowEmptyAfterExplicitDelete?: boolean
  /** Keeps Redo aware of an intentional clear after Undo resets the current intent. */
  hadExplicitDelete?: boolean
}

export function legacyCustomShowsKey(path: string): string {
  return `ai-slides-custom-shows:${path}`
}

/** Reject malformed/stale data instead of dropping an old show's slides during migration. */
export function parseLegacyCustomShows(raw: string, slideCount: number): CustomShow[] | null {
  let parsed: unknown
  try {
    parsed = JSON.parse(raw)
  } catch {
    return null
  }
  if (!Array.isArray(parsed)) return null
  const seenLegacy = new Set<string>()
  const usedNative = new Set<string>()
  for (const candidate of parsed) {
    if (
      !candidate ||
      typeof candidate !== 'object' ||
      typeof candidate.id !== 'string' ||
      seenLegacy.has(candidate.id) ||
      typeof candidate.name !== 'string' ||
      !candidate.name.trim() ||
      !Array.isArray(candidate.slideIndices) ||
      candidate.slideIndices.some(
        (index: unknown) =>
          !Number.isInteger(index) || (index as number) < 0 || (index as number) >= slideCount,
      )
    )
      return null
    seenLegacy.add(candidate.id)
    if (/^(?:0|[1-9]\d*)$/.test(candidate.id) && Number(candidate.id) <= 0xffffffff)
      usedNative.add(candidate.id)
  }
  let nextId = 0
  return parsed.map((candidate: { id: string; name: string; slideIndices: number[] }) => {
    let id = candidate.id
    if (!usedNative.has(id)) {
      while (usedNative.has(String(nextId))) nextId++
      id = String(nextId++)
      usedNative.add(id)
    }
    return { id, name: candidate.name, slideIndices: [...candidate.slideIndices] }
  })
}

/** Only an empty native list may be upgraded. The old key remains until a successful file save. */
export async function importLegacyCustomShows(options: {
  path: string
  slideCount: number
  storage: LegacyShowStorage
  readNative: () => Promise<CustomShow[]>
  writeNative: (shows: CustomShow[]) => Promise<CustomShow[] | null>
}): Promise<{
  kind: 'none' | 'invalid' | 'imported'
  shows?: CustomShow[]
  pending?: PendingLegacyCustomShows
}> {
  const native = await options.readNative()
  if (native.length) return { kind: 'none', shows: native }
  const key = legacyCustomShowsKey(options.path)
  const raw = options.storage.getItem(key)
  if (raw == null) return { kind: 'none', shows: native }
  const legacy = parseLegacyCustomShows(raw, options.slideCount)
  if (!legacy) return { kind: 'invalid', shows: native }
  if (!legacy.length) return { kind: 'none', shows: native }
  // Recheck the native list immediately before writing: another editor may have added a show.
  const current = await options.readNative()
  if (current.length) return { kind: 'none', shows: current }
  const written = await options.writeNative(legacy)
  if (!written || JSON.stringify(written) !== JSON.stringify(legacy))
    return { kind: 'invalid', shows: current }
  return { kind: 'imported', shows: written, pending: { key, raw } }
}

/** Do not erase a legacy record changed by another window or an unsaved/undone migration. */
export function retireLegacyCustomShows(
  storage: LegacyShowStorage,
  pending: PendingLegacyCustomShows | null,
  hasNativeShows: boolean,
): boolean {
  if (
    !pending ||
    (!hasNativeShows && !pending.allowEmptyAfterExplicitDelete) ||
    storage.getItem(pending.key) !== pending.raw
  )
    return false
  storage.removeItem(pending.key)
  return true
}
