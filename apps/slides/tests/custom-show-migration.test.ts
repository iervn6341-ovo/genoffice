import { describe, expect, it } from 'vitest'
import {
  importLegacyCustomShows,
  legacyCustomShowsKey,
  parseLegacyCustomShows,
  retireLegacyCustomShows,
} from '../src/renderer/custom-show-migration'
import type { CustomShow } from '../src/renderer/slideshow-utils'

function storage() {
  const entries = new Map<string, string>()
  return {
    getItem: (key: string) => entries.get(key) ?? null,
    setItem: (key: string, value: string) => entries.set(key, value),
    removeItem: (key: string) => entries.delete(key),
  }
}

describe('legacy custom-show migration', () => {
  it('remaps legacy ids, preserves slide order/duplicates, and retires the old key only after native save', async () => {
    const store = storage()
    const path = '/tmp/legacy.pptx'
    const key = legacyCustomShowsKey(path)
    store.setItem(
      key,
      JSON.stringify([
        { id: 'cs-a', name: 'Board', slideIndices: [3, 0, 3] },
        { id: '4', name: 'Detail', slideIndices: [2] },
      ]),
    )
    let native: CustomShow[] = []
    const result = await importLegacyCustomShows({
      path,
      slideCount: 5,
      storage: store,
      readNative: async () => native,
      writeNative: async (shows) => (native = shows),
    })
    expect(result.kind).toBe('imported')
    expect(native).toEqual([
      { id: '0', name: 'Board', slideIndices: [3, 0, 3] },
      { id: '4', name: 'Detail', slideIndices: [2] },
    ])
    expect(store.getItem(key)).not.toBeNull()
    expect(retireLegacyCustomShows(store, result.pending!, false)).toBe(false)
    expect(store.getItem(key)).not.toBeNull()
    expect(retireLegacyCustomShows(store, result.pending!, true)).toBe(true)
    expect(store.getItem(key)).toBeNull()
  })

  it('retires an intentionally cleared migrated list, but keeps the backup after an undo', () => {
    const store = storage()
    const key = legacyCustomShowsKey('/tmp/clear.pptx')
    store.setItem(key, 'legacy-data')
    const pending = { key, raw: 'legacy-data', allowEmptyAfterExplicitDelete: true }
    expect(retireLegacyCustomShows(store, pending, false)).toBe(true)
    store.setItem(key, 'legacy-data')
    pending.allowEmptyAfterExplicitDelete = false
    expect(retireLegacyCustomShows(store, pending, false)).toBe(false)
    expect(store.getItem(key)).toBe('legacy-data')
  })

  it('never overwrites an existing native list or removes a key changed by another window', async () => {
    const store = storage()
    const path = '/tmp/native.pptx'
    const key = legacyCustomShowsKey(path)
    store.setItem(key, JSON.stringify([{ id: 'cs-a', name: 'Old', slideIndices: [0] }]))
    const native = [{ id: '12', name: 'Native', slideIndices: [1] }]
    let writes = 0
    const result = await importLegacyCustomShows({
      path,
      slideCount: 2,
      storage: store,
      readNative: async () => native,
      writeNative: async () => {
        writes++
        return null
      },
    })
    expect(result).toMatchObject({ kind: 'none', shows: native })
    expect(writes).toBe(0)
    expect(store.getItem(key)).not.toBeNull()
    const pending = { key, raw: store.getItem(key)! }
    store.setItem(key, 'newer value')
    expect(retireLegacyCustomShows(store, pending, true)).toBe(false)
    expect(store.getItem(key)).toBe('newer value')
  })

  it('rejects malformed or stale legacy records rather than losing slides', () => {
    expect(parseLegacyCustomShows('not-json', 3)).toBeNull()
    expect(
      parseLegacyCustomShows(JSON.stringify([{ id: 'x', name: 'Stale', slideIndices: [9] }]), 3),
    ).toBeNull()
    expect(
      parseLegacyCustomShows(JSON.stringify([{ id: 'x', name: '', slideIndices: [0] }]), 3),
    ).toBeNull()
    expect(
      parseLegacyCustomShows(
        JSON.stringify([
          { id: 'x', name: 'One', slideIndices: [0] },
          { id: 'x', name: 'Two', slideIndices: [1] },
        ]),
        3,
      ),
    ).toBeNull()
  })
})
