/**
 * Mouse selection must never rewrite another cell: the cell editor DOM is reused when the edit
 * target changes, so a late commit has to be written back to the cell the text came from, and
 * must not close the session of the cell that replaced it.
 */
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import type { ShapeRenderNode } from '@genoffice/pptx-render'
import { TextEditOverlay } from '../src/renderer/TextEditOverlay'
import {
  cellEditKey,
  commitCellEdit,
  endCellEdit,
  parseCellEditKey,
} from '../src/renderer/table-actions'
import type { ActionCtx, EditingCellState } from '../src/renderer/action-context'

;(globalThis as unknown as { IS_REACT_ACT_ENVIRONMENT: boolean }).IS_REACT_ACT_ENVIRONMENT = true

describe('cell edit keys', () => {
  it('round-trips ids that contain underscores', () => {
    const c = { sourceId: 'tbl_2', row: 3, col: 11 }
    expect(cellEditKey(c)).toBe('cellEdit_tbl_2_3_11')
    expect(parseCellEditKey(cellEditKey(c))).toEqual(c)
    expect(parseCellEditKey('n1')).toBeNull()
    expect(parseCellEditKey(undefined)).toBeNull()
  })
})

describe('commitCellEdit writes to the owner cell', () => {
  const editTableCell = vi.fn()
  beforeEach(() => {
    editTableCell.mockReset().mockResolvedValue({ fake: 'slide' })
    ;(globalThis as unknown as { window: unknown }).window = Object.assign(globalThis.window, {
      slidesApi: { editTableCell },
    })
  })
  const ctxWith = (editing: EditingCellState | null) => {
    let state: EditingCellState | null = editing
    const ctx = {
      current: 0,
      get editingCell() {
        return state
      },
      setEditingCell: (v: unknown) => {
        state =
          typeof v === 'function' ? (v as (c: typeof state) => typeof state)(state) : (v as never)
      },
      applySlide: vi.fn(),
    } as unknown as ActionCtx
    return { ctx, get: () => state }
  }
  const paras = [{ runs: [{ text: 'A text' }] }] as never

  it('a late commit of cell A while the editor already moved to cell B lands in A and keeps B editing', async () => {
    const { ctx, get } = ctxWith({ sourceId: 'tbl_2', row: 2, col: 2 })
    await commitCellEdit(ctx, paras, cellEditKey({ sourceId: 'tbl_2', row: 1, col: 0 }))
    expect(editTableCell).toHaveBeenCalledTimes(1)
    expect(editTableCell.mock.calls[0]![0]).toMatchObject({ sourceId: 'tbl_2', row: 1, col: 0 })
    expect(get()).toEqual({ sourceId: 'tbl_2', row: 2, col: 2 })
  })

  it('a normal commit ends the session', async () => {
    const { ctx, get } = ctxWith({ sourceId: 'tbl_2', row: 1, col: 0 })
    await commitCellEdit(ctx, paras, cellEditKey({ sourceId: 'tbl_2', row: 1, col: 0 }))
    expect(get()).toBeNull()
  })

  it('without an owner it falls back to the current editing cell', async () => {
    const { ctx } = ctxWith({ sourceId: 'tbl_2', row: 0, col: 1 })
    await commitCellEdit(ctx, paras)
    expect(editTableCell.mock.calls[0]![0]).toMatchObject({ row: 0, col: 1 })
  })

  it('endCellEdit leaves another cell alone', () => {
    const { ctx, get } = ctxWith({ sourceId: 'tbl_2', row: 2, col: 2 })
    endCellEdit(ctx, { sourceId: 'tbl_2', row: 1, col: 0 })
    expect(get()).not.toBeNull()
    endCellEdit(ctx, { sourceId: 'tbl_2', row: 2, col: 2 })
    expect(get()).toBeNull()
  })
})

describe('TextEditOverlay hands the owning cell to its commit', () => {
  let host: HTMLDivElement
  let root: Root
  beforeEach(() => {
    host = document.createElement('div')
    document.body.appendChild(host)
    root = createRoot(host)
  })
  afterEach(() => {
    act(() => root.unmount())
    host.remove()
  })
  const cellNode = (row: number, col: number) =>
    ({
      id: `cellEdit_tbl_2_${row}_${col}`,
      type: 'shape',
      sourceId: 'tbl_2',
      box: { x: 0, y: 0, w: 200, h: 40, rotationDeg: 0 },
      fill: { kind: 'none' },
    }) as unknown as ShapeRenderNode

  it('dirty text of cell A, re-targeted to B, commits as B only once repopulated; A commits as A', () => {
    const onCommit = vi.fn()
    const onCancel = vi.fn()
    const render = (n: ShapeRenderNode) =>
      act(() =>
        root.render(createElement(TextEditOverlay, { node: n, scale: 1, onCommit, onCancel })),
      )
    render(cellNode(1, 0))
    const editor = host.querySelector('.slide-text-editor') as HTMLElement
    Object.defineProperty(editor, 'isContentEditable', { value: true })
    editor.focus()
    editor.textContent = 'edited A'
    act(() => {
      editor.blur()
    })
    expect(onCommit).toHaveBeenCalledTimes(1)
    expect(onCommit.mock.calls[0]![1]).toBe('cellEdit_tbl_2_1_0')
    // the same editor is now pointed at cell B (empty): an unchanged close is B's, not A's
    render(cellNode(2, 2))
    editor.focus()
    act(() => editor.blur())
    expect(onCancel).toHaveBeenCalled()
    expect(onCancel.mock.calls.at(-1)![0]).toBe('cellEdit_tbl_2_2_2')
  })

  it('swallows the native text drag inside a cell editor', () => {
    act(() =>
      root.render(
        createElement(TextEditOverlay, {
          node: cellNode(0, 0),
          scale: 1,
          onCommit: vi.fn(),
          onCancel: vi.fn(),
        }),
      ),
    )
    const editor = host.querySelector('.slide-text-editor') as HTMLElement
    const ev = new Event('dragstart', { bubbles: true, cancelable: true })
    editor.dispatchEvent(ev)
    expect(ev.defaultPrevented).toBe(true)
  })
})
