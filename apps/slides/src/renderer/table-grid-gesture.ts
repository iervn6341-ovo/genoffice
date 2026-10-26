/**
 * Live column / row border drag of a table. Every pointer frame re-grids the table locally (no IPC
 * latency), and one request at a time asks the main process for the authoritative preview (text
 * re-wrapped at the new widths, rows grown to fit). The reply is re-gridded onto the pointer's
 * latest position, so the dragged border always sits under the pointer while the reflow trails a
 * frame or two behind. The main process treats the frames as one gesture: a single undo step on the
 * settled commit, nothing at all when the drag is cancelled.
 */
import type { RenderSlide, TableRenderNode } from '@genoffice/pptx-render'
import type { SlidesApi } from '../shared/ipc'
import { FIT_WIDTH } from './app-constants'
import { changedGrid, regridTable, type GridChange } from './table-grid-preview'

export interface GridGestureHost {
  /** The slide as the canvas currently shows it */
  getSlide: (slideIndex: number) => RenderSlide | undefined
  /** Put a slide on the canvas; `settled` = an authoritative commit (marks the document dirty) */
  setSlide: (slideIndex: number, slide: RenderSlide, settled: boolean) => void
}

type Api = Pick<SlidesApi, 'setTableColWidth' | 'setTableRowHeight'>

interface Live {
  slideIndex: number
  sourceId: string
  /** The slide the drag started from */
  orig: RenderSlide
  /** Latest authoritative slide / table (the original until a preview reply lands) */
  auth: RenderSlide
  tbl: TableRenderNode
  want: GridChange
  inFlight: Promise<void> | null
  pending: boolean
}

function tableOf(slide: RenderSlide | undefined, sourceId: string): TableRenderNode | null {
  const n = slide?.nodes.find((x) => x.sourceId === sourceId)
  return n && n.type === 'table' ? (n as TableRenderNode) : null
}

function withTable(slide: RenderSlide, tbl: TableRenderNode): RenderSlide {
  return { ...slide, nodes: slide.nodes.map((n) => (n.sourceId === tbl.sourceId ? tbl : n)) }
}

export function createTableGridGesture(host: GridGestureHost, api: () => Api) {
  let live: Live | null = null

  const request = (
    at: { slideIndex: number; sourceId: string },
    c: GridChange,
    extra: { preview?: boolean; cancel?: boolean },
  ) =>
    c.kind === 'col'
      ? api().setTableColWidth({
          slideIndex: at.slideIndex,
          sourceId: at.sourceId,
          col: c.col,
          wPx: c.wPx,
          fitWidthPx: FIT_WIDTH,
          ...(c.neighbor ? { neighbor: c.neighbor } : {}),
          ...extra,
        })
      : api().setTableRowHeight({
          slideIndex: at.slideIndex,
          sourceId: at.sourceId,
          row: c.row,
          hPx: c.hPx,
          fitWidthPx: FIT_WIDTH,
          ...extra,
        })

  /** Show the pointer's grid on top of the latest authoritative layout */
  const show = (l: Live) => {
    const g = changedGrid(l.tbl, l.want)
    host.setSlide(l.slideIndex, withTable(l.auth, regridTable(l.tbl, g.gridX, g.gridY)), false)
  }

  const pump = (l: Live) => {
    if (l.inFlight) {
      l.pending = true
      return
    }
    l.pending = false
    l.inFlight = request(l, l.want, { preview: true })
      .then((slide) => {
        l.inFlight = null
        if (live !== l || !slide) return
        const tbl = tableOf(slide, l.sourceId)
        if (tbl) {
          l.auth = slide
          l.tbl = tbl
          show(l)
        }
        if (l.pending) pump(l)
      })
      .catch(() => {
        l.inFlight = null
      })
  }

  /** Put the model and the canvas back to where the drag started */
  const rollBack = async (l: Live): Promise<void> => {
    const slide = await request(l, l.want, { cancel: true })
    host.setSlide(l.slideIndex, slide ?? l.orig, false)
  }

  return {
    /** The pointer asks for this border position */
    move(slideIndex: number, sourceId: string, change: GridChange): void {
      if (!live || live.sourceId !== sourceId || live.slideIndex !== slideIndex) {
        const auth = host.getSlide(slideIndex)
        const tbl = tableOf(auth, sourceId)
        if (!auth || !tbl) return
        live = {
          slideIndex,
          sourceId,
          orig: auth,
          auth,
          tbl,
          want: change,
          inFlight: null,
          pending: false,
        }
      }
      live.want = change
      show(live)
      pump(live)
    },
    /** Settle the drag at `change`: one undo step, the authoritative layout replaces the preview */
    async commit(slideIndex: number, sourceId: string, change: GridChange): Promise<void> {
      const l = live
      live = null
      if (l?.inFlight) await l.inFlight
      const slide = await request({ slideIndex, sourceId }, change, {})
      if (slide) host.setSlide(slideIndex, slide, true)
      else if (l) await rollBack(l)
    },
    /** Abandon the drag: nothing stays in the model or on the undo stack */
    async cancel(): Promise<void> {
      const l = live
      live = null
      if (!l) return
      if (l.inFlight) await l.inFlight
      await rollBack(l)
    },
    /** True while a border drag owns the table */
    active(): boolean {
      return live != null
    },
  }
}
