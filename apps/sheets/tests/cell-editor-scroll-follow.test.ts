import { describe, expect, it } from 'vitest'
import { installCellEditorScrollFollow } from '../src/renderer/cell-editor-scroll-follow'

function harness(visible: boolean) {
  const calls: Array<boolean | 'fit'> = []
  const state = { visible }
  const handlers = new Set<() => void>()
  const frames: Array<() => void> = []
  const dispose = installCellEditorScrollFollow({
    bridge: {
      isVisible: () => ({ visible: state.visible }),
      refreshEditCellPosition: (resetSizeOnly) => void calls.push(resetSizeOnly),
    },
    resize: { fitTextSize: () => void calls.push('fit') },
    onScroll: (h) => {
      handlers.add(h)
      return () => handlers.delete(h)
    },
    schedule: (cb) => void frames.push(cb),
  })
  return {
    calls,
    state,
    scroll: () => handlers.forEach((h) => h()),
    flush: () => frames.splice(0).forEach((f) => f()),
    handlers,
    dispose,
  }
}

describe('cell editor follows grid scroll', () => {
  it('re-places the open editor from the new scroll offset, once per frame', () => {
    const h = harness(true)
    h.scroll()
    h.scroll()
    h.scroll()
    expect(h.calls).toEqual([])
    h.flush()
    expect(h.calls).toEqual([false, 'fit'])
    h.scroll()
    h.flush()
    expect(h.calls).toEqual([false, 'fit', false, 'fit'])
  })

  it('does nothing while no cell is being edited', () => {
    const h = harness(false)
    h.scroll()
    h.flush()
    expect(h.calls).toEqual([])
  })

  it('skips the frame when the editor closed after the scroll', () => {
    const h = harness(true)
    h.scroll()
    h.state.visible = false
    h.flush()
    expect(h.calls).toEqual([])
  })

  it('unsubscribes on dispose', () => {
    const h = harness(true)
    h.dispose()
    expect(h.handlers.size).toBe(0)
  })
})
