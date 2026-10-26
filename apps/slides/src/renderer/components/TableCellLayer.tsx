/**
 * Cell-range selection layer for a selected table: invisible hit areas over the cells (press to
 * select a cell, drag to select a block, Shift+press to extend, double-click to edit), press-to-select
 * gutters for whole rows / columns / the table, and the highlight of the selected block. It lives in
 * table-local coordinates inside a div that carries the table's rotation and flips, so rotated and
 * mirrored tables behave like plain ones. Cells keep a gap around their borders so the Konva column /
 * row resize grips and the transform handles underneath stay reachable.
 */
import { useEffect, useRef } from 'react'
import type { TableRenderNode } from '@genoffice/pptx-render'
import { tableCellSelColors } from '../canvas-colors'
import { onNodeGesture } from '../node-gesture'
import { TABLE_BORDER_GRIP_PX, tableCellAtPoint, tableLocalPointFromStage } from '../table-hit'
import {
  colSel,
  gridCols,
  gridRows,
  rectBounds,
  rowSel,
  selectionRect,
  singleCellSel,
  type CellAddr,
  type TableCellSel,
} from '../table-selection'

interface Props {
  table: TableRenderNode
  sel: TableCellSel | null
  zoom: number
  /** Frame color chosen for this spot of the slide (see selectionChromeColor) */
  chrome: string
  onSelect: (sel: TableCellSel) => void
  onEditCell: (at: CellAddr) => void
  onContextMenu: (at: CellAddr, x: number, y: number) => void
  /** The press that selected this table landed on this cell and the button is still down: keep extending from it */
  dragFrom?: CellAddr | null
}

type Drag =
  { kind: 'cell'; anchor: CellAddr } | { kind: 'row'; from: number } | { kind: 'col'; from: number }

export function TableCellLayer({
  table,
  sel,
  zoom,
  chrome,
  onSelect,
  onEditCell,
  onContextMenu,
  dragFrom,
}: Props) {
  const rootRef = useRef<HTMLDivElement>(null)
  const dragRef = useRef<Drag | null>(null)
  const latest = useRef({ table, onSelect })
  latest.current = { table, onSelect }
  const box = table.box
  const u = 1 / Math.max(zoom, 0.1)
  const gap = TABLE_BORDER_GRIP_PX * u
  const colors = tableCellSelColors(chrome)

  // While the table is moved on the canvas the model still holds its old box until release: follow
  // the live offset (resizing hides the block, whose cells are about to re-flow). A resize stays
  // hidden past the gesture's end until the committed table arrives — showing the block at the
  // pre-release geometry for the length of the IPC round trip is a visible jump — with a timer as
  // the fallback when the commit changes nothing.
  const resizeRef = useRef<{ ended: boolean; timer: number | null } | null>(null)
  const unhide = () => {
    const el = rootRef.current
    if (resizeRef.current?.timer != null) window.clearTimeout(resizeRef.current.timer)
    resizeRef.current = null
    if (el) el.style.visibility = ''
  }
  useEffect(
    () =>
      onNodeGesture((g) => {
        const el = rootRef.current
        if (!el || g.sourceId !== latest.current.table.sourceId) return
        if (g.kind === 'move') {
          el.style.visibility = ''
          el.style.translate = `${g.dx ?? 0}px ${g.dy ?? 0}px`
        } else if (g.kind === 'resize') {
          el.style.visibility = 'hidden'
          resizeRef.current = { ended: false, timer: null }
        } else if (g.kind === 'end') {
          el.style.translate = ''
          if (resizeRef.current) {
            resizeRef.current.ended = true
            resizeRef.current.timer = window.setTimeout(unhide, 1200)
          } else el.style.visibility = ''
        }
      }),
    [],
  )
  // the committed table landed after the resize ended: show the block at its real place
  useEffect(() => {
    if (resizeRef.current?.ended) unhide()
  }, [table])
  useEffect(
    () => () => {
      if (resizeRef.current?.timer != null) window.clearTimeout(resizeRef.current.timer)
    },
    [],
  )

  // a drag keeps extending the block while the button is down, anywhere on the page
  useEffect(() => {
    // A press on an unselected table mounts this layer mid-gesture: resume the block selection from there
    if (dragFrom) dragRef.current = { kind: 'cell', anchor: dragFrom }
    const move = (e: MouseEvent) => {
      const drag = dragRef.current
      if (!drag || !rootRef.current) return
      if (!(e.buttons & 1)) {
        // the release happened outside our listeners (e.g. before mount): never extend with the button up
        dragRef.current = null
        return
      }
      const { table: tbl, onSelect: select } = latest.current
      const parent = rootRef.current.parentElement
      if (!parent) return
      const r = parent.getBoundingClientRect()
      const k = parent.offsetWidth ? r.width / parent.offsetWidth : 1
      const p = tableLocalPointFromStage(
        { x: (e.clientX - r.left) / k, y: (e.clientY - r.top) / k },
        tbl.box,
        0,
      )
      if (drag.kind === 'cell') {
        // outside the table the pointer counts as its nearest edge cell
        const q = {
          x: Math.min(Math.max(p.x, 0), tbl.box.w - 0.01),
          y: Math.min(Math.max(p.y, 0), tbl.box.h - 0.01),
        }
        const hit = tableCellAtPoint(tbl, q)
        if (hit) {
          select({
            sourceId: tbl.sourceId,
            anchor: drag.anchor,
            focus: { row: hit.row, col: hit.col },
          })
        }
      } else if (drag.kind === 'row') {
        const g = tbl.gridY
        let i = g.length - 2
        for (let n = 0; n < g.length - 1; n++) {
          if (p.y < g[n + 1]!) {
            i = n
            break
          }
        }
        select(rowSel(tbl, tbl.sourceId, drag.from, Math.max(0, i)))
      } else {
        const w = tbl.gridX[tbl.gridX.length - 1] ?? tbl.box.w
        const lx = tbl.rtl ? w - p.x : p.x
        const g = tbl.gridX
        let i = g.length - 2
        for (let n = 0; n < g.length - 1; n++) {
          if (lx < g[n + 1]!) {
            i = n
            break
          }
        }
        select(colSel(tbl, tbl.sourceId, drag.from, Math.max(0, i)))
      }
    }
    const up = () => {
      dragRef.current = null
    }
    window.addEventListener('mousemove', move)
    window.addEventListener('mouseup', up)
    return () => {
      window.removeEventListener('mousemove', move)
      window.removeEventListener('mouseup', up)
    }
    // dragFrom is only the mount-time seed
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [])

  const press = (e: React.MouseEvent) => {
    // the ribbon field that had focus must let go, or keys keep going to it instead of the cells
    ;(document.activeElement as HTMLElement | null)?.blur?.()
    e.preventDefault()
  }

  const rect = sel ? selectionRect(table, sel) : null
  const bounds = rect ? rectBounds(table, rect) : null
  const tw = table.gridX[table.gridX.length - 1] ?? box.w
  const flipX = box.flipH ? -1 : 1
  const flipY = box.flipV ? -1 : 1

  return (
    <div
      ref={rootRef}
      className="tcl-root"
      data-testid="table-cell-layer"
      style={{
        position: 'absolute',
        left: box.x,
        top: box.y,
        width: box.w,
        height: box.h,
        transform: `rotate(${box.rotationDeg ?? 0}deg) scale(${flipX}, ${flipY})`,
        transformOrigin: 'center center',
        zIndex: 12,
        pointerEvents: 'none',
      }}
    >
      {bounds && (
        <div
          data-testid="table-cell-selection"
          style={{
            position: 'absolute',
            left: bounds.x,
            top: bounds.y,
            width: bounds.w,
            height: bounds.h,
            background: colors.fill,
            outline: `${2 * u}px solid ${colors.stroke}`,
            outlineOffset: -1 * u,
            pointerEvents: 'none',
          }}
        />
      )}
      {table.cells.map((c) => (
        <div
          key={`${c.row}_${c.col}`}
          className="tcl-cell"
          data-cell={`${c.row},${c.col}`}
          style={{
            position: 'absolute',
            left: c.x + gap,
            top: c.y + gap,
            width: Math.max(0, c.w - 2 * gap),
            height: Math.max(0, c.h - 2 * gap),
            pointerEvents: 'auto',
          }}
          onMouseDown={(e) => {
            if (e.button !== 0) return
            press(e)
            const at = { row: c.row, col: c.col }
            if (e.shiftKey && sel) {
              onSelect({ ...sel, focus: at })
              dragRef.current = { kind: 'cell', anchor: sel.anchor }
            } else {
              onSelect(singleCellSel(table.sourceId, at))
              dragRef.current = { kind: 'cell', anchor: at }
            }
          }}
          onDoubleClick={(e) => {
            e.preventDefault()
            onEditCell({ row: c.row, col: c.col })
          }}
          onContextMenu={(e) => {
            e.preventDefault()
            onContextMenu({ row: c.row, col: c.col }, e.clientX, e.clientY)
          }}
        />
      ))}
      {/* whole-row handles on the left edge, whole-column handles above the top edge, the corner selects the table */}
      {Array.from({ length: gridRows(table) }, (_, r) => (
        <div
          key={`row_${r}`}
          className="tcl-gutter tcl-gutter-row"
          data-gutter={`row-${r}`}
          style={{
            position: 'absolute',
            left: -12 * u,
            top: table.gridY[r]! + u,
            width: 8 * u,
            height: Math.max(0, table.gridY[r + 1]! - table.gridY[r]! - 2 * u),
            pointerEvents: 'auto',
            ['--tcl-gutter' as string]: colors.gutter,
          }}
          onMouseDown={(e) => {
            if (e.button !== 0) return
            press(e)
            if (e.shiftKey && sel) {
              const from = Math.min(sel.anchor.row, sel.focus.row)
              onSelect(rowSel(table, table.sourceId, from, r))
              dragRef.current = { kind: 'row', from }
            } else {
              onSelect(rowSel(table, table.sourceId, r))
              dragRef.current = { kind: 'row', from: r }
            }
          }}
        />
      ))}
      {Array.from({ length: gridCols(table) }, (_, c) => {
        const x0 = table.rtl ? tw - table.gridX[c + 1]! : table.gridX[c]!
        return (
          <div
            key={`col_${c}`}
            className="tcl-gutter tcl-gutter-col"
            data-gutter={`col-${c}`}
            style={{
              position: 'absolute',
              left: x0 + u,
              top: -12 * u,
              width: Math.max(0, table.gridX[c + 1]! - table.gridX[c]! - 2 * u),
              height: 8 * u,
              pointerEvents: 'auto',
              ['--tcl-gutter' as string]: colors.gutter,
            }}
            onMouseDown={(e) => {
              if (e.button !== 0) return
              press(e)
              if (e.shiftKey && sel) {
                const from = Math.min(sel.anchor.col, sel.focus.col)
                onSelect(colSel(table, table.sourceId, from, c))
                dragRef.current = { kind: 'col', from }
              } else {
                onSelect(colSel(table, table.sourceId, c))
                dragRef.current = { kind: 'col', from: c }
              }
            }}
          />
        )
      })}
      <div
        className="tcl-gutter tcl-gutter-corner"
        data-gutter="all"
        style={{
          position: 'absolute',
          left: -12 * u,
          top: -12 * u,
          width: 8 * u,
          height: 8 * u,
          pointerEvents: 'auto',
          ['--tcl-gutter' as string]: colors.gutter,
        }}
        onMouseDown={(e) => {
          if (e.button !== 0) return
          press(e)
          onSelect({
            sourceId: table.sourceId,
            anchor: { row: 0, col: 0 },
            focus: { row: gridRows(table) - 1, col: gridCols(table) - 1 },
          })
        }}
      />
    </div>
  )
}
