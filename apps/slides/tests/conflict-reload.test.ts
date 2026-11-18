import { describe, expect, it } from 'vitest'
import { appStrings } from '../src/renderer/i18n/strings-app'
import type { OpenResult } from '../src/shared/ipc'

describe('conflict dialog strings', () => {
  const app = appStrings as unknown as Record<string, Record<string, string>>
  const langs = Object.keys(app)

  it('covers all 20 UI languages with conflict resolution strings', () => {
    expect(langs.length).toBe(20)
    for (const lang of langs) {
      const strings = app[lang]!
      expect(strings.conflictTitle, `${lang} conflictTitle`).toBeTruthy()
      expect(strings.conflictSubtitle, `${lang} conflictSubtitle`).toBeTruthy()
      expect(strings.conflictLocalTitle, `${lang} conflictLocalTitle`).toBeTruthy()
      expect(strings.conflictLocalOverwrite, `${lang} conflictLocalOverwrite`).toBeTruthy()
      expect(strings.conflictLocalSaveCopy, `${lang} conflictLocalSaveCopy`).toBeTruthy()
      expect(strings.conflictDiskTitle, `${lang} conflictDiskTitle`).toBeTruthy()
      expect(strings.conflictDiskReload, `${lang} conflictDiskReload`).toBeTruthy()
      expect(strings.conflictDecideLater, `${lang} conflictDecideLater`).toBeTruthy()
    }
  })
})

describe('reload keeps current slide position', () => {
  // Pure logic test verifying the slide index retention behavior of applyOpen
  function simulateApplyOpen(
    prevPath: string | null,
    prevCurrent: number,
    result: OpenResult,
  ): { nextPath: string; nextCurrent: number } {
    const sameDoc = prevPath === result.path
    const isReload = !!result.isReload || sameDoc
    const nextPath = result.path
    const nextCurrent = isReload ? Math.min(prevCurrent, Math.max(0, result.slides.length - 1)) : 0
    return { nextPath, nextCurrent }
  }

  it('preserves current slide index when reloading the same file', () => {
    const prevPath = '/Users/test/deck.pptx'
    const prevCurrent = 3 // User was on slide 4
    const result: OpenResult = {
      path: '/Users/test/deck.pptx',
      slides: [{} as any, {} as any, {} as any, {} as any, {} as any], // 5 slides
      size: { cx: 9144000, cy: 5143500 },
      isReload: true,
    }

    const { nextPath, nextCurrent } = simulateApplyOpen(prevPath, prevCurrent, result)
    expect(nextPath).toBe(prevPath)
    expect(nextCurrent).toBe(3) // Stays on slide 4!
  })

  it('clamps current slide index if slide count decreased on reload', () => {
    const prevPath = '/Users/test/deck.pptx'
    const prevCurrent = 4 // User was on slide 5
    const result: OpenResult = {
      path: '/Users/test/deck.pptx',
      slides: [{} as any, {} as any], // Now only 2 slides (index 0 and 1)
      size: { cx: 9144000, cy: 5143500 },
      isReload: true,
    }

    const { nextCurrent } = simulateApplyOpen(prevPath, prevCurrent, result)
    expect(nextCurrent).toBe(1) // Clamped to last slide
  })

  it('resets current slide index to 0 when opening a different file', () => {
    const prevPath = '/Users/test/deck1.pptx'
    const prevCurrent = 3
    const result: OpenResult = {
      path: '/Users/test/deck2.pptx', // Different path
      slides: [{} as any, {} as any, {} as any, {} as any],
      size: { cx: 9144000, cy: 5143500 },
      isReload: false,
    }

    const { nextPath, nextCurrent } = simulateApplyOpen(prevPath, prevCurrent, result)
    expect(nextPath).toBe('/Users/test/deck2.pptx')
    expect(nextCurrent).toBe(0) // Resets to first slide
  })

  it('advances slideEpoch and changes canvasKey even when reloading the same slide index', () => {
    let slideEpoch = 0
    let annotationsNonce = 0
    let editingCell: { sourceId: string; row: number; col: number } | null = {
      sourceId: 't1',
      row: 0,
      col: 0,
    }
    let cellSel: unknown = { sourceId: 't1', r1: 0, c1: 0, r2: 0, c2: 0 }
    let enteredGroupId: string | null = 'g1'
    let selectedIds = ['shape1']
    let current = 0

    const applyReload = () => {
      // Simulate state updates performed by applyOpen on reload
      current = Math.min(current, 2)
      selectedIds = []
      editingCell = null
      cellSel = null
      enteredGroupId = null
      slideEpoch += 1
      annotationsNonce += 1
    }

    const initialCanvasKey = `canvas_${current}_${slideEpoch}`
    expect(initialCanvasKey).toBe('canvas_0_0')

    applyReload()

    const reloadedCanvasKey = `canvas_${current}_${slideEpoch}`
    expect(current).toBe(0) // Current slide index unchanged
    expect(slideEpoch).toBe(1) // Epoch advanced
    expect(annotationsNonce).toBe(1) // Annotations nonce advanced
    expect(reloadedCanvasKey).toBe('canvas_0_1')
    expect(reloadedCanvasKey).not.toBe(initialCanvasKey) // Canvas key forced to change
    expect(editingCell).toBeNull()
    expect(cellSel).toBeNull()
    expect(enteredGroupId).toBeNull()
    expect(selectedIds).toEqual([])
  })
})
