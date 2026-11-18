import { beforeEach, describe, expect, it, vi } from 'vitest'
import { handleGlobalKeydown, resetNudgeStateForTest } from '../src/renderer/keyboard-actions'
import * as clipboardActions from '../src/renderer/clipboard-actions'
import * as slideActions from '../src/renderer/slide-actions'
import * as showActions from '../src/renderer/show-actions'
import type { ActionCtx } from '../src/renderer/action-context'

vi.mock('../src/renderer/clipboard-actions', () => ({
  copySelected: vi.fn(),
  cutSelected: vi.fn(),
  copySlideAt: vi.fn(),
  copyFormat: vi.fn(),
  pasteFormat: vi.fn(),
  pasteClipboard: vi.fn(),
  duplicateSelected: vi.fn(),
  deleteSelected: vi.fn(),
}))
vi.mock('../src/renderer/slide-actions', () => ({
  cutSlideAt: vi.fn(),
  deleteSlideAt: vi.fn(),
  deleteSlidesAt: vi.fn(),
  duplicateSlideAt: vi.fn(),
}))
vi.mock('../src/renderer/arrange-actions', () => ({}))
vi.mock('../src/renderer/show-actions', () => ({ startSlideShow: vi.fn() }))

function makeCtx(over: Record<string, unknown> = {}): ActionCtx {
  return {
    slideShow: false,
    presenter: false,
    editing: null,
    selectedIds: [],
    slide: { nodes: [] },
    slides: [{}],
    current: 0,
    slideSelection: [],
    masterItems: null,
    inkTool: 'select',
    viewMode: 'normal',
    brushMode: null,
    enteredGroupId: null,
    findNodeCtx: () => null,
    ...over,
  } as unknown as ActionCtx
}

function keydown(key: string, init: KeyboardEventInit = {}): KeyboardEvent {
  return new KeyboardEvent('keydown', { key, metaKey: true, cancelable: true, ...init })
}

function selectText(): void {
  const div = document.createElement('div')
  div.textContent = 'answer from the AI panel'
  document.body.appendChild(div)
  const range = document.createRange()
  range.selectNodeContents(div)
  const sel = window.getSelection()!
  sel.removeAllRanges()
  sel.addRange(range)
}

describe('slide show shortcuts', () => {
  beforeEach(() => {
    vi.clearAllMocks()
    document.body.innerHTML = ''
  })

  it('starts from the current slide with Command+Enter on macOS', () => {
    const ctx = makeCtx({ current: 2, slides: [{}, {}, {}] })
    const e = keydown('Enter')

    handleGlobalKeydown(ctx, e, 'MacIntel')

    expect(e.defaultPrevented).toBe(true)
    expect(showActions.startSlideShow).toHaveBeenCalledWith(ctx, false)
  })

  it('starts from the current slide with Shift+F5 on Windows', () => {
    const ctx = makeCtx({ current: 2, slides: [{}, {}, {}] })
    const e = keydown('F5', { metaKey: false, shiftKey: true })

    handleGlobalKeydown(ctx, e, 'Win32')

    expect(e.defaultPrevented).toBe(true)
    expect(showActions.startSlideShow).toHaveBeenCalledWith(ctx, false)
  })

  it('does not repurpose Ctrl+Enter on Windows', () => {
    const ctx = makeCtx()
    const e = keydown('Enter', { metaKey: false, ctrlKey: true })

    handleGlobalKeydown(ctx, e, 'Win32')

    expect(e.defaultPrevented).toBe(false)
    expect(showActions.startSlideShow).not.toHaveBeenCalled()
  })

  it('does not start a show from a text field on macOS', () => {
    const input = document.createElement('textarea')
    document.body.appendChild(input)
    input.focus()
    const ctx = makeCtx()
    const e = keydown('Enter')

    handleGlobalKeydown(ctx, e, 'MacIntel')

    expect(e.defaultPrevented).toBe(false)
    expect(showActions.startSlideShow).not.toHaveBeenCalled()
  })

  it('does not start a show when another handler consumed Enter', () => {
    const ctx = makeCtx()
    const e = keydown('Enter')
    e.preventDefault()

    handleGlobalKeydown(ctx, e, 'MacIntel')

    expect(showActions.startSlideShow).not.toHaveBeenCalled()
  })

  it('does not start a show while confirming a crop', () => {
    const ctx = makeCtx({ cropTarget: {} })
    const e = keydown('Enter')

    handleGlobalKeydown(ctx, e, 'MacIntel')

    expect(e.defaultPrevented).toBe(false)
    expect(showActions.startSlideShow).not.toHaveBeenCalled()
  })
})

describe('copy shortcuts with a DOM text selection', () => {
  beforeEach(() => {
    vi.clearAllMocks()
    document.body.innerHTML = ''
    window.getSelection()?.removeAllRanges()
  })

  it('does not hijack ⌘C for the slide clipboard when nothing is selected on canvas', () => {
    selectText()
    const e = keydown('c')
    handleGlobalKeydown(makeCtx(), e)
    expect(e.defaultPrevented).toBe(false)
    expect(clipboardActions.copySlideAt).not.toHaveBeenCalled()
  })

  it('does not hijack ⌘C/⌘X for the element clipboard when shapes are selected', () => {
    selectText()
    const ctx = makeCtx({ selectedIds: ['s1'] })
    const c = keydown('c')
    handleGlobalKeydown(ctx, c)
    const x = keydown('x')
    handleGlobalKeydown(ctx, x)
    expect(c.defaultPrevented).toBe(false)
    expect(x.defaultPrevented).toBe(false)
    expect(clipboardActions.copySelected).not.toHaveBeenCalled()
    expect(clipboardActions.cutSelected).not.toHaveBeenCalled()
  })

  it('still copies the current slide / selection when the selection is collapsed', () => {
    const noSel = keydown('c')
    handleGlobalKeydown(makeCtx(), noSel)
    expect(noSel.defaultPrevented).toBe(true)
    expect(clipboardActions.copySlideAt).toHaveBeenCalledTimes(1)

    const withSel = keydown('x')
    handleGlobalKeydown(makeCtx({ selectedIds: ['s1'] }), withSel)
    expect(withSel.defaultPrevented).toBe(true)
    expect(clipboardActions.cutSelected).toHaveBeenCalledTimes(1)
  })

  it('keeps non-copy shortcuts unaffected by a text selection', () => {
    selectText()
    const del = new KeyboardEvent('keydown', { key: 'Delete', cancelable: true })
    handleGlobalKeydown(makeCtx({ selectedIds: ['s1'] }), del)
    expect(del.defaultPrevented).toBe(true)
    expect(clipboardActions.deleteSelected).toHaveBeenCalledTimes(1)
  })

  it('does not hijack ⌘X for slide cut while text is selected', () => {
    selectText()
    const cut = keydown('x')
    handleGlobalKeydown(makeCtx(), cut)
    expect(cut.defaultPrevented).toBe(false)
    expect(slideActions.cutSlideAt).not.toHaveBeenCalled()
  })
})

describe('Delete/Backspace on the thumbnail pane', () => {
  beforeEach(() => {
    vi.clearAllMocks()
    document.body.innerHTML = ''
    window.getSelection()?.removeAllRanges()
  })

  const plain = (key: string) => new KeyboardEvent('keydown', { key, cancelable: true })

  it('deletes the current slide when nothing is selected on canvas', () => {
    const ctx = makeCtx({ current: 1, slides: [{}, {}] })
    const e = plain('Backspace')
    handleGlobalKeydown(ctx, e)
    expect(e.defaultPrevented).toBe(true)
    expect(slideActions.deleteSlideAt).toHaveBeenCalledWith(ctx, 1)
  })

  it('deletes the whole thumbnail multi-selection in one call', () => {
    const ctx = makeCtx({ current: 3, slides: [{}, {}, {}, {}], slideSelection: [1, 3] })
    const e = plain('Delete')
    handleGlobalKeydown(ctx, e)
    expect(e.defaultPrevented).toBe(true)
    expect(slideActions.deleteSlidesAt).toHaveBeenCalledWith(ctx, [1, 3])
    expect(slideActions.deleteSlideAt).not.toHaveBeenCalled()
  })

  it('⌘⇧D duplicates the current slide (PowerPoint for Mac); ⌘D alone does not', () => {
    const ctx = makeCtx({ current: 1, slides: [{}, {}] })
    const dup = keydown('D', { shiftKey: true })
    handleGlobalKeydown(ctx, dup)
    expect(dup.defaultPrevented).toBe(true)
    expect(slideActions.duplicateSlideAt).toHaveBeenCalledWith(ctx, 1)
    vi.clearAllMocks()
    handleGlobalKeydown(ctx, keydown('d'))
    expect(slideActions.duplicateSlideAt).not.toHaveBeenCalled()
  })

  it('also deletes in the slide sorter view', () => {
    const e = plain('Delete')
    handleGlobalKeydown(makeCtx({ viewMode: 'sorter' }), e)
    expect(slideActions.deleteSlideAt).toHaveBeenCalledTimes(1)
  })

  it('does nothing while an ink tool is active', () => {
    const e = plain('Delete')
    handleGlobalKeydown(makeCtx({ inkTool: 'pen' }), e)
    expect(e.defaultPrevented).toBe(false)
    expect(slideActions.deleteSlideAt).not.toHaveBeenCalled()
  })

  it('does nothing in reading view', () => {
    const e = plain('Backspace')
    handleGlobalKeydown(makeCtx({ viewMode: 'reading' }), e)
    expect(e.defaultPrevented).toBe(false)
    expect(slideActions.deleteSlideAt).not.toHaveBeenCalled()
  })

  it('leaves the key to a plain-DOM text selection', () => {
    selectText()
    const e = plain('Delete')
    handleGlobalKeydown(makeCtx(), e)
    expect(e.defaultPrevented).toBe(false)
    expect(slideActions.deleteSlideAt).not.toHaveBeenCalled()
  })

  it('does not fire from master view or with modifiers', () => {
    const master = plain('Delete')
    handleGlobalKeydown(makeCtx({ masterItems: [] }), master)
    const alt = new KeyboardEvent('keydown', { key: 'Backspace', altKey: true, cancelable: true })
    handleGlobalKeydown(makeCtx(), alt)
    expect(slideActions.deleteSlideAt).not.toHaveBeenCalled()
  })
})

describe('arrow-key nudge of a multi-selection', () => {
  const boxOf = (x: number, y: number) => ({ x, y, w: 100, h: 50, rotationDeg: 0 })
  const nodes: Record<string, ReturnType<typeof boxOf>> = {
    a: boxOf(10, 10),
    b: boxOf(200, 40),
    c: boxOf(300, 80),
  }
  beforeEach(() => {
    resetNudgeStateForTest()
    vi.clearAllMocks()
  })

  const make = (onTransform: (...a: unknown[]) => Promise<void>) =>
    makeCtx({
      selectedIds: ['a', 'b', 'c'],
      findNodeCtx: (id: string) => (nodes[id] ? { node: { box: nodes[id] } } : null),
      onTransform,
    })

  it('hands every selected element to onTransform in the same keypress', () => {
    const onTransform = vi.fn((..._a: unknown[]) => Promise.resolve())
    handleGlobalKeydown(make(onTransform), keydown('ArrowRight', { metaKey: false }))
    expect(onTransform.mock.calls.map((c: unknown[]) => [c[0], (c[1] as { x: number }).x])).toEqual(
      [
        ['a', 11],
        ['b', 201],
        ['c', 301],
      ],
    )
  })

  it('shift moves by 10 and keeps size and rotation', () => {
    const onTransform = vi.fn((..._a: unknown[]) => Promise.resolve())
    handleGlobalKeydown(make(onTransform), keydown('ArrowDown', { metaKey: false, shiftKey: true }))
    expect(onTransform.mock.calls[0]![1]).toEqual({ x: 10, y: 20, w: 100, h: 50, rotationDeg: 0 })
  })

  it('a quick second press steps from the box already requested, not the stale render', async () => {
    let release: () => void = () => undefined
    const gate = new Promise<void>((r) => (release = r))
    const onTransform = vi.fn((..._a: unknown[]) => gate)
    const ctx = make(onTransform)
    handleGlobalKeydown(ctx, keydown('ArrowLeft', { metaKey: false }))
    handleGlobalKeydown(ctx, keydown('ArrowLeft', { metaKey: false }))
    const second = onTransform.mock.calls.slice(3).map((c: unknown[]) => (c[1] as { x: number }).x)
    expect(second).toEqual([8, 198, 298])
    release()
    await gate
    await Promise.resolve()
    // once the nudges landed, the next press starts from the rendered box again
    handleGlobalKeydown(ctx, keydown('ArrowLeft', { metaKey: false }))
    expect((onTransform.mock.calls[6]![1] as { x: number }).x).toBe(9)
  })

  it('optimistically updates the current slide nodes via applySlide for instant visual response', () => {
    const onTransform = vi.fn((..._a: unknown[]) => Promise.resolve())
    const applySlide = vi.fn()
    const curSlide = {
      id: 's0',
      widthPx: 1280,
      heightPx: 720,
      nodes: [
        { id: '1', sourceId: 'a', box: { ...nodes.a } },
        { id: '2', sourceId: 'b', box: { ...nodes.b } },
        { id: '3', sourceId: 'c', box: { ...nodes.c } },
      ],
    }
    const ctx = makeCtx({
      selectedIds: ['a', 'b'],
      findNodeCtx: (id: string) => (nodes[id] ? { node: { box: nodes[id] } } : null),
      onTransform,
      applySlide,
      current: 0,
      slides: [curSlide],
    })
    handleGlobalKeydown(ctx, keydown('ArrowRight', { metaKey: false }))
    expect(applySlide).toHaveBeenCalledTimes(1)
    const updated = applySlide.mock.calls[0]![1] as typeof curSlide
    expect(updated.nodes.find((n) => n.sourceId === 'a')!.box.x).toBe(11)
    expect(updated.nodes.find((n) => n.sourceId === 'b')!.box.x).toBe(201)
    expect(updated.nodes.find((n) => n.sourceId === 'c')!.box.x).toBe(300)
  })

  it('batches continuous arrow nudges inside a history batch', async () => {
    vi.useFakeTimers()
    const beginHistoryBatch = vi.fn().mockResolvedValue(true)
    const endHistoryBatch = vi.fn().mockResolvedValue(null)
    window.slidesApi = {
      beginHistoryBatch,
      endHistoryBatch,
    } as unknown as typeof window.slidesApi

    const onTransform = vi.fn((..._a: unknown[]) => Promise.resolve())
    const ctx = make(onTransform)

    handleGlobalKeydown(ctx, keydown('ArrowRight', { metaKey: false }))
    handleGlobalKeydown(ctx, keydown('ArrowRight', { metaKey: false }))
    expect(beginHistoryBatch).toHaveBeenCalledTimes(1)
    expect(endHistoryBatch).not.toHaveBeenCalled()

    vi.advanceTimersByTime(450)
    expect(endHistoryBatch).toHaveBeenCalledTimes(1)
    expect(endHistoryBatch).toHaveBeenCalledWith({ snapshot: false })
    vi.useRealTimers()
  })
})

describe('Esc drops the element selection', () => {
  beforeEach(() => {
    vi.clearAllMocks()
    document.body.innerHTML = ''
  })

  it('clears the selection when elements are selected', () => {
    const setSelectedIds = vi.fn()
    const ctx = makeCtx({ selectedIds: ['a', 'b'], setSelectedIds })
    const e = keydown('Escape', { metaKey: false })
    handleGlobalKeydown(ctx, e, 'MacIntel')
    expect(setSelectedIds).toHaveBeenCalledWith([])
    expect(e.defaultPrevented).toBe(true)
  })

  it('leaves in-group editing first (group stays selected)', () => {
    const setSelectedIds = vi.fn()
    const ctx = makeCtx({
      selectedIds: ['child'],
      enteredGroupId: 'g1',
      setSelectedIds,
      setEnteredGroupId: vi.fn(),
    })
    handleGlobalKeydown(ctx, keydown('Escape', { metaKey: false }), 'MacIntel')
    expect(setSelectedIds).toHaveBeenCalledWith(['g1'])
  })

  it('does nothing while a menu is open (Esc closes the menu only)', () => {
    const menu = document.createElement('div')
    menu.className = 'ctx-menu'
    document.body.appendChild(menu)
    const setSelectedIds = vi.fn()
    const ctx = makeCtx({ selectedIds: ['a'], setSelectedIds })
    handleGlobalKeydown(ctx, keydown('Escape', { metaKey: false }), 'MacIntel')
    expect(setSelectedIds).not.toHaveBeenCalled()
  })

  it('does nothing while editing text or typing in a field', () => {
    const setSelectedIds = vi.fn()
    handleGlobalKeydown(
      makeCtx({ selectedIds: ['a'], editing: { sourceId: 'a' }, setSelectedIds }),
      keydown('Escape', { metaKey: false }),
      'MacIntel',
    )
    const input = document.createElement('input')
    document.body.appendChild(input)
    input.focus()
    handleGlobalKeydown(
      makeCtx({ selectedIds: ['a'], setSelectedIds }),
      keydown('Escape', { metaKey: false }),
      'MacIntel',
    )
    expect(setSelectedIds).not.toHaveBeenCalled()
  })

  it('with nothing selected it is a no-op', () => {
    const setSelectedIds = vi.fn()
    const e = keydown('Escape', { metaKey: false })
    handleGlobalKeydown(makeCtx({ selectedIds: [], setSelectedIds }), e, 'MacIntel')
    expect(setSelectedIds).not.toHaveBeenCalled()
    expect(e.defaultPrevented).toBe(false)
  })
})
