/**
 * Moving a table must not rewrite its stored grid. A row grows to fit its text, so the DRAWN table is
 * taller than the stored a:tr heights; committing a move with the drawn size as if it were a resize
 * re-scaled the stored heights up a little more on every move (the row looked like it kept growing).
 */
import { describe, expect, it, vi } from 'vitest'
import { addTable, createBlankPptx, openPptx, type TableElement } from '@genoffice/pptx-engine'
import { runTxn } from '@genoffice/pptx-ops'
import { buildRenderSlide } from '@genoffice/pptx-render'
import type { Session } from '../src/main/session-state'
import { applyTransformBatch } from '../src/main/batch-transform'

vi.mock('electron', () => ({
  BrowserWindow: { getFocusedWindow: () => null },
  webContents: { fromId: () => null },
}))
vi.mock('../src/main/fonts', () => ({
  createSystemFontMetrics: () => ({}),
}))

const FIT = 1280
const EMU = 9525

async function tableWithGrownRows() {
  const opened = await openPptx(await createBlankPptx())
  const session = { path: '', opened, fitWidthPx: FIT, undoStack: [], redoStack: [] } as Session
  const added = addTable(opened, 0, {
    rows: 2,
    cols: 2,
    offset: { x: 100 * EMU, y: 200 * EMU, cx: 400 * EMU, cy: 20 * EMU },
    rowHeightsEmu: [10 * EMU, 10 * EMU],
  })!
  for (const row of [0, 1]) {
    runTxn(opened, {
      ops: [
        {
          op: 'setTableCell',
          target: { slide: 0, el: added.elementId },
          row,
          col: 0,
          paragraphs: [{ runs: [{ text: 'Row text that needs height' }] }],
        },
      ],
    })
  }
  const slide = opened.deck.slides[0]!
  const table = () => slide.elements.find((e) => e.type === 'table') as TableElement
  const drawn = () => {
    const node = buildRenderSlide(slide, opened.deck.size, { fitWidthPx: FIT }).nodes.find(
      (n) => n.type === 'table',
    )!
    return node.box
  }
  return { session, table, drawn, id: table().id }
}

describe('moving a table whose rows were grown by text', () => {
  it('the drawn table is taller than the stored grid (the situation that used to inflate it)', async () => {
    const { table, drawn } = await tableWithGrownRows()
    const stored = table().rowHeights.reduce((a, b) => a + b, 0) / EMU
    expect(drawn().h).toBeGreaterThan(stored + 5)
  })

  it('a move flagged keepTableSize leaves the stored row heights and frame size alone', async () => {
    const { session, table, drawn, id } = await tableWithGrownRows()
    const before = [...table().rowHeights]
    const ext = { ...table().transform.offset }
    for (let n = 1; n <= 5; n++) {
      const b = drawn()
      expect(
        applyTransformBatch(session, {
          slideIndex: 0,
          fitWidthPx: FIT,
          items: [
            {
              sourceId: id,
              xPx: b.x + 30,
              yPx: b.y + 10,
              wPx: b.w,
              hPx: b.h,
              rotationDeg: 0,
              resizeTableGrid: true,
              keepTableSize: true,
            },
          ],
        }),
      ).toBe(true)
    }
    expect(table().rowHeights).toEqual(before)
    expect(table().transform.offset.cx).toBe(ext.cx)
    expect(table().transform.offset.cy).toBe(ext.cy)
    expect(table().transform.offset.x).toBe(ext.x + 150 * EMU)
  })

  it('without the flag the same commit re-scales the stored rows (what the flag prevents)', async () => {
    const { session, table, drawn, id } = await tableWithGrownRows()
    const before = [...table().rowHeights]
    const b = drawn()
    applyTransformBatch(session, {
      slideIndex: 0,
      fitWidthPx: FIT,
      items: [
        {
          sourceId: id,
          xPx: b.x + 30,
          yPx: b.y,
          wPx: b.w,
          hPx: b.h,
          rotationDeg: 0,
          resizeTableGrid: true,
        },
      ],
    })
    expect(table().rowHeights).not.toEqual(before)
  })

  it('a real resize still redistributes the grid', async () => {
    const { session, table, drawn, id } = await tableWithGrownRows()
    const b = drawn()
    applyTransformBatch(session, {
      slideIndex: 0,
      fitWidthPx: FIT,
      items: [
        {
          sourceId: id,
          xPx: b.x,
          yPx: b.y,
          wPx: b.w * 1.5,
          hPx: b.h,
          rotationDeg: 0,
          resizeTableGrid: true,
        },
      ],
    })
    expect(table().colWidths.reduce((a, b2) => a + b2, 0) / EMU).toBeCloseTo(b.w * 1.5, 0)
  })
})
