/**
 * A table border drag is ONE undo step however many preview frames it streams, leaves no step when
 * cancelled, and never damages the history (the main-process half of the live border preview).
 */
import { describe, expect, it, vi } from 'vitest'
import { createBlankPptx, openPptx } from '@genoffice/pptx-engine'
import { runTxn, type TxnRequest } from '@genoffice/pptx-ops'
import { buildRenderSlide, type TableRenderNode } from '@genoffice/pptx-render'
import type { Session } from '../src/main/session-state'
import { redoStep, undoStep } from '../src/main/session-state'
import { applyTableGridTxn } from '../src/main/table-grid-txn'

vi.mock('electron', () => ({
  BrowserWindow: { getFocusedWindow: () => null },
  webContents: { fromId: () => null },
}))
vi.mock('../src/main/fonts', () => ({
  createSystemFontMetrics: () => ({}),
}))

const FIT = 1280
const emu = (px: number) => px * 9525

async function sessionWithTable(): Promise<{ session: Session; id: string }> {
  const opened = await openPptx(await createBlankPptx())
  const session = { path: '', opened, fitWidthPx: FIT, undoStack: [], redoStack: [] } as Session
  const r = runTxn(opened, {
    ops: [
      {
        op: 'addTable',
        target: { slide: 0 },
        rows: 2,
        cols: 3,
        offset: { x: emu(100), y: emu(300), cx: emu(600), cy: emu(120) },
      },
    ],
  })
  expect(r.applied).toBe(true)
  const t = tableNode(session)
  return { session, id: t.sourceId }
}

function tableNode(session: Session): TableRenderNode {
  const rendered = buildRenderSlide(session.opened.deck.slides[0]!, session.opened.deck.size, {
    fitWidthPx: FIT,
  })
  return rendered.nodes.find((n) => n.type === 'table') as TableRenderNode
}

const widths = (session: Session) => {
  const g = tableNode(session).gridX
  return g.slice(1).map((v, i) => Math.round(v - g[i]!))
}

const colOps = (id: string, a: number, b: number): TxnRequest => ({
  ops: [
    { op: 'setTableColWidth', target: { slide: 0, el: id }, col: 0, wEmu: emu(a) },
    { op: 'setTableColWidth', target: { slide: 0, el: id }, col: 1, wEmu: emu(b) },
  ],
})

// the real journaling runner of slides-main, reduced to what the gesture needs
const journaled = (session: Session, _source: 'edit', req: TxnRequest) =>
  runTxn(session.opened, req)

describe('table border gesture history', () => {
  it('many preview frames plus the settled commit are one undo step', async () => {
    const { session, id } = await sessionWithTable()
    const before = widths(session)
    for (const w of [190, 180, 170, 160]) {
      expect(applyTableGridTxn(session, colOps(id, w, 400 - w), { preview: true }, journaled)).toBe(
        true,
      )
      expect(session.transformPreview).toBe(true)
      expect(session.undoStack.length).toBe(1)
    }
    expect(widths(session).slice(0, 2)).toEqual([160, 240])
    expect(applyTableGridTxn(session, colOps(id, 150, 250), {}, journaled)).toBe(true)
    expect(session.transformPreview).toBe(false)
    expect(session.undoStack.length).toBe(1)
    expect(widths(session).slice(0, 2)).toEqual([150, 250])
    expect(undoStep(session)).toBe(true)
    expect(widths(session)).toEqual(before)
    expect(undoStep(session)).toBe(false)
    expect(redoStep(session)).toBe(true)
    expect(widths(session).slice(0, 2)).toEqual([150, 250])
  })

  it('a plain commit with no frames before it is one step too', async () => {
    const { session, id } = await sessionWithTable()
    expect(applyTableGridTxn(session, colOps(id, 150, 250), {}, journaled)).toBe(true)
    expect(session.undoStack.length).toBe(1)
    expect(session.transformPreview).toBeFalsy()
  })

  it('cancel rolls the frames back and leaves no undo step', async () => {
    const { session, id } = await sessionWithTable()
    const before = widths(session)
    applyTableGridTxn(session, colOps(id, 180, 220), { preview: true }, journaled)
    applyTableGridTxn(session, colOps(id, 170, 230), { preview: true }, journaled)
    expect(widths(session)).not.toEqual(before)
    expect(applyTableGridTxn(session, colOps(id, 0, 0), { cancel: true }, journaled)).toBe(true)
    expect(widths(session)).toEqual(before)
    expect(session.transformPreview).toBe(false)
    expect(session.undoStack.length).toBe(0)
    // a stray cancel (nothing in flight) is harmless
    expect(applyTableGridTxn(session, colOps(id, 0, 0), { cancel: true }, journaled)).toBe(true)
    expect(session.undoStack.length).toBe(0)
  })

  it('a refused first frame sets no flag and keeps the redo stack', async () => {
    const { session } = await sessionWithTable()
    session.redoStack.push({} as never)
    const bad = colOps('no-such-table', 100, 100)
    expect(applyTableGridTxn(session, bad, { preview: true }, journaled)).toBe(false)
    expect(session.transformPreview).toBeFalsy()
    expect(session.undoStack.length).toBe(0)
    expect(session.redoStack.length).toBe(1)
  })

  it('a cancelled drag after an earlier edit keeps that edit undoable as its own step', async () => {
    const { session, id } = await sessionWithTable()
    const orig = widths(session)
    expect(applyTableGridTxn(session, colOps(id, 120, 280), {}, journaled)).toBe(true)
    const edited = widths(session)
    applyTableGridTxn(session, colOps(id, 90, 310), { preview: true }, journaled)
    applyTableGridTxn(session, colOps(id, 0, 0), { cancel: true }, journaled)
    expect(widths(session)).toEqual(edited)
    expect(undoStep(session)).toBe(true)
    expect(widths(session)).toEqual(orig)
  })
})
