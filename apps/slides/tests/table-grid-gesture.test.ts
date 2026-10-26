import { describe, expect, it } from 'vitest'
import type { RenderSlide, TableCellRender, TableRenderNode } from '@genoffice/pptx-render'
import { createTableGridGesture } from '../src/renderer/table-grid-gesture'

function tableNode(widths: number[], rows: number[]): TableRenderNode {
  const gridX = [0]
  for (const w of widths) gridX.push(gridX[gridX.length - 1]! + w)
  const gridY = [0]
  for (const h of rows) gridY.push(gridY[gridY.length - 1]! + h)
  const cells: TableCellRender[] = []
  rows.forEach((_, r) =>
    widths.forEach((_, c) =>
      cells.push({
        row: r,
        col: c,
        x: gridX[c]!,
        y: gridY[r]!,
        w: widths[c]!,
        h: rows[r]!,
        fill: { kind: 'none' },
      } as TableCellRender),
    ),
  )
  return {
    id: 'r_t',
    type: 'table',
    sourceId: 't',
    box: { x: 0, y: 0, w: gridX[gridX.length - 1]!, h: gridY[gridY.length - 1]!, rotationDeg: 0 },
    cells,
    gridX,
    gridY,
  } as unknown as TableRenderNode
}

const slideOf = (t: TableRenderNode): RenderSlide => ({ nodes: [t] }) as unknown as RenderSlide

function setup() {
  const shown: Array<{ slide: RenderSlide; settled: boolean }> = []
  let current = slideOf(tableNode([100, 100], [40]))
  const calls: Array<Record<string, unknown>> = []
  const waiting: Array<(s: RenderSlide | null) => void> = []
  const api = {
    setTableColWidth: (op: Record<string, unknown>) => {
      calls.push(op)
      return new Promise<RenderSlide | null>((res) => waiting.push(res))
    },
    setTableRowHeight: (op: Record<string, unknown>) => {
      calls.push(op)
      return new Promise<RenderSlide | null>((res) => waiting.push(res))
    },
  }
  const g = createTableGridGesture(
    {
      getSlide: () => current,
      setSlide: (_i, slide, settled) => {
        current = slide
        shown.push({ slide, settled })
      },
    },
    () => api as never,
  )
  const widthsShown = () => {
    const t = current.nodes[0] as TableRenderNode
    return t.gridX.slice(1).map((v, i) => v - t.gridX[i]!)
  }
  return { g, shown, calls, waiting, widthsShown }
}

const move = (w: number) =>
  ({ kind: 'col', col: 0, wPx: w, neighbor: { col: 1, wPx: 200 - w } }) as const

describe('table border drag gesture', () => {
  it('shows every pointer frame at once, one preview request in flight at a time', () => {
    const { g, calls, widthsShown } = setup()
    g.move(0, 't', move(90))
    expect(widthsShown()).toEqual([90, 110])
    g.move(0, 't', move(80))
    g.move(0, 't', move(70))
    expect(widthsShown()).toEqual([70, 130])
    // only the first frame went out; the rest wait for its reply
    expect(calls).toHaveLength(1)
    expect(calls[0]).toMatchObject({ preview: true, wPx: 90 })
  })

  it('a reply is re-gridded onto the pointer, then the latest position goes out', async () => {
    const { g, calls, waiting, widthsShown } = setup()
    g.move(0, 't', move(90))
    g.move(0, 't', move(60))
    // the main process answers for 90 (rows grew to 55 on the way)
    const reply = slideOf(tableNode([90, 110], [55]))
    waiting.shift()!(reply)
    await Promise.resolve()
    await Promise.resolve()
    expect(widthsShown()).toEqual([60, 140])
    expect(calls).toHaveLength(2)
    expect(calls[1]).toMatchObject({ preview: true, wPx: 60 })
  })

  it('commit waits for the frame in flight, then settles without the preview flag', async () => {
    const { g, calls, waiting, shown } = setup()
    g.move(0, 't', move(90))
    const done = g.commit(0, 't', move(75))
    await Promise.resolve()
    expect(calls).toHaveLength(1)
    waiting.shift()!(slideOf(tableNode([90, 110], [40])))
    await new Promise((r) => setTimeout(r, 0))
    expect(calls).toHaveLength(2)
    expect(calls[1]).toMatchObject({ wPx: 75 })
    expect(calls[1]!.preview).toBeUndefined()
    waiting.shift()!(slideOf(tableNode([75, 125], [40])))
    await done
    expect(shown[shown.length - 1]!.settled).toBe(true)
    expect(g.active()).toBe(false)
  })

  it('cancel waits for the frame in flight, then rolls the model back', async () => {
    const { g, calls, waiting, widthsShown } = setup()
    g.move(0, 't', move(90))
    const cancelled = g.cancel()
    await Promise.resolve()
    expect(g.active()).toBe(false)
    expect(calls).toHaveLength(1)
    waiting.shift()!(slideOf(tableNode([90, 110], [40])))
    await new Promise((r) => setTimeout(r, 0))
    expect(calls[1]).toMatchObject({ cancel: true })
    waiting.shift()!(slideOf(tableNode([100, 100], [40])))
    await cancelled
    expect(widthsShown()).toEqual([100, 100])
  })

  it('cancel after a reply asks the main process to roll the frames back', async () => {
    const { g, calls, waiting, shown, widthsShown } = setup()
    g.move(0, 't', move(90))
    waiting.shift()!(slideOf(tableNode([90, 110], [40])))
    await new Promise((r) => setTimeout(r, 0))
    const cancelled = g.cancel()
    await new Promise((r) => setTimeout(r, 0))
    expect(calls[calls.length - 1]).toMatchObject({ cancel: true })
    waiting.shift()!(slideOf(tableNode([100, 100], [40])))
    await cancelled
    expect(widthsShown()).toEqual([100, 100])
    expect(shown[shown.length - 1]!.settled).toBe(false)
  })

  it('a failed settled commit rolls the preview frames back', async () => {
    const { g, calls, waiting, widthsShown } = setup()
    g.move(0, 't', move(90))
    waiting.shift()!(slideOf(tableNode([90, 110], [40])))
    await new Promise((r) => setTimeout(r, 0))
    const done = g.commit(0, 't', move(75))
    await new Promise((r) => setTimeout(r, 0))
    waiting.shift()!(null) // the commit was refused
    await new Promise((r) => setTimeout(r, 0))
    expect(calls[calls.length - 1]).toMatchObject({ cancel: true })
    waiting.shift()!(slideOf(tableNode([100, 100], [40])))
    await done
    expect(widthsShown()).toEqual([100, 100])
  })
})
