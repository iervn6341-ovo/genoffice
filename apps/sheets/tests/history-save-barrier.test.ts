import { describe, expect, it, vi } from 'vitest'

import { HistorySaveBarrier } from '../src/renderer/history-save-barrier'

function deferred<T>() {
  let resolve!: (value: T) => void
  const promise = new Promise<T>((settle) => {
    resolve = settle
  })
  return { promise, resolve }
}

describe('HistorySaveBarrier', () => {
  it('waits for Save and carried history before Undo, then waits for Undo before another Save', async () => {
    const barrier = new HistorySaveBarrier()
    const firstSave = deferred<void>()
    const carry = deferred<boolean>()
    const undo = deferred<void>()
    const events: string[] = []
    const saving = barrier.runSave(async () => {
      events.push('first-save-start')
      await firstSave.promise
      events.push('first-save-end')
    })
    const undoing = barrier.runHistory(
      async () => {
        events.push('undo-start')
        await undo.promise
        events.push('undo-end')
      },
      () => carry.promise,
    )
    const savingUndo = barrier.runSave(async () => {
      events.push('second-save')
    })
    await vi.waitFor(() => expect(events).toContain('first-save-start'))
    expect(events).toEqual(['first-save-start'])
    firstSave.resolve()
    await saving
    expect(events).toEqual(['first-save-start', 'first-save-end'])
    carry.resolve(true)
    await vi.waitFor(() => expect(events).toContain('undo-start'))
    expect(events).toContain('undo-start')
    expect(events).not.toContain('second-save')
    undo.resolve()
    await Promise.all([undoing, savingUndo])
    expect(events).toEqual([
      'first-save-start',
      'first-save-end',
      'undo-start',
      'undo-end',
      'second-save',
    ])
  })

  it('does not save an unexecuted history change when carried history times out', async () => {
    const barrier = new HistorySaveBarrier()
    let applied = false
    const undoing = barrier.runHistory(
      async () => {
        applied = true
      },
      async () => false,
    )
    await expect(undoing).rejects.toThrow('history is still loading')
    expect(applied).toBe(false)
    const saved = await barrier.runSave(async () => 'later save remains available')
    expect(saved).toBe('later save remains available')
  })

  it('holds a second Save until the previous Save has installed its carried history', async () => {
    const barrier = new HistorySaveBarrier()
    const carry = deferred<boolean>()
    let writes = 0
    const first = await barrier.runSave(async () => {
      writes += 1
    })
    expect(first).toBeUndefined()
    const second = barrier.runSave(
      async () => {
        writes += 1
      },
      () => carry.promise,
    )
    await Promise.resolve()
    expect(writes).toBe(1)
    carry.resolve(true)
    await second
    expect(writes).toBe(2)

    const rejected = barrier.runSave(
      async () => {
        writes += 1
      },
      async () => false,
    )
    await expect(rejected).rejects.toThrow('history is still loading')
    expect(writes).toBe(2)
  })
})
