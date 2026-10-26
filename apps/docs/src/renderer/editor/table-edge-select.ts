import { Extension } from '@tiptap/core'
import { Plugin, PluginKey } from '@tiptap/pm/state'
import { CellSelection } from '@tiptap/pm/tables'
import type { EditorView } from '@tiptap/pm/view'
import { posInsideSelection } from './selection-hit'

/**
 * Word's table edge selectors: a click just LEFT of a row selects the whole row (drag down to
 * take more rows); a click just ABOVE a column selects the whole column (drag sideways for more).
 * Both produce an ordinary CellSelection, so every range operation (shading, merge, delete,
 * insert, paragraph format) works on it.
 */

/** width of the selector band outside the table (px) */
export const EDGE_BAND_PX = 16

export type EdgeKind = 'row' | 'column'
export interface EdgeHit {
  kind: EdgeKind
  table: HTMLTableElement
  /** the cell the selection anchors on (first cell of the row / first-row cell of the column) */
  cell: HTMLTableCellElement
}

const inside = (v: number, lo: number, hi: number) => v >= lo && v <= hi

export function edgeHitAt(
  tables: Iterable<HTMLTableElement>,
  x: number,
  y: number,
  band = EDGE_BAND_PX,
): EdgeHit | null {
  for (const table of tables) {
    const r = table.getBoundingClientRect()
    if (r.width === 0 && r.height === 0) continue
    if (inside(x, r.left - band, r.left - 1) && inside(y, r.top, r.bottom)) {
      const row = Array.from(table.rows).find((tr) => {
        const b = tr.getBoundingClientRect()
        return inside(y, b.top, b.bottom)
      })
      const cell = row?.cells[0]
      if (cell) return { kind: 'row', table, cell }
    }
    if (inside(y, r.top - band, r.top - 1) && inside(x, r.left, r.right)) {
      const first = table.rows[0]
      const cell = first
        ? Array.from(first.cells).find((c) => {
            const b = c.getBoundingClientRect()
            return inside(x, b.left, b.right)
          })
        : undefined
      if (cell) return { kind: 'column', table, cell }
    }
  }
  return null
}

/** document position of the table cell element (the position BEFORE the cell node) */
function cellPos(view: EditorView, cell: HTMLElement): number | null {
  const $pos = view.state.doc.resolve(view.posAtDOM(cell, 0))
  for (let depth = $pos.depth; depth > 0; depth--) {
    const role = $pos.node(depth).type.spec.tableRole
    if (role === 'cell' || role === 'header_cell') return $pos.before(depth)
  }
  return null
}

/** the cell of `hit.table` under (x, y) along the hit's axis, for drag extension */
function headCellAt(hit: EdgeHit, x: number, y: number): HTMLTableCellElement | null {
  if (hit.kind === 'row') {
    const row = Array.from(hit.table.rows).find((tr) => {
      const b = tr.getBoundingClientRect()
      return inside(y, b.top, b.bottom)
    })
    return row?.cells[0] ?? null
  }
  const first = hit.table.rows[0]
  return (
    (first &&
      Array.from(first.cells).find((c) => {
        const b = c.getBoundingClientRect()
        return inside(x, b.left, b.right)
      })) ||
    null
  )
}

function selectEdge(view: EditorView, hit: EdgeHit, head: HTMLElement): boolean {
  const a = cellPos(view, hit.cell)
  const h = cellPos(view, head)
  if (a === null || h === null) return false
  const $a = view.state.doc.resolve(a)
  const $h = view.state.doc.resolve(h)
  const selection =
    hit.kind === 'row' ? CellSelection.rowSelection($a, $h) : CellSelection.colSelection($a, $h)
  view.dispatch(view.state.tr.setSelection(selection))
  return true
}

function edgeSelectPlugin(): Plugin {
  return new Plugin({
    key: new PluginKey('tableEdgeSelect'),
    props: {
      handleDOMEvents: {
        mousedown(view, event) {
          // A right press inside a cell range must not move the native caret (Chromium puts it, or
          // a word selection, under the pointer and ProseMirror then collapses the range): swallow
          // the press, the context menu still opens and acts on the whole range.
          if (event.button === 2 && view.state.selection instanceof CellSelection) {
            const at = view.posAtCoords({ left: event.clientX, top: event.clientY })
            if (at && posInsideSelection(view.state.selection, at.pos)) {
              event.preventDefault()
              return true
            }
            return false
          }
          if (event.button !== 0 || !view.editable) return false
          const tables = view.dom.querySelectorAll<HTMLTableElement>('table.doc-table')
          const hit = edgeHitAt(tables, event.clientX, event.clientY)
          if (!hit) return false
          if (!selectEdge(view, hit, hit.cell)) return false
          event.preventDefault()
          view.focus()
          const onMove = (e: MouseEvent) => {
            if (view.isDestroyed) return onUp()
            const head = headCellAt(hit, e.clientX, e.clientY)
            if (head) selectEdge(view, hit, head)
          }
          const onUp = () => {
            window.removeEventListener('mousemove', onMove)
            window.removeEventListener('mouseup', onUp)
          }
          window.addEventListener('mousemove', onMove)
          window.addEventListener('mouseup', onUp)
          return true
        },
      },
    },
  })
}

export const TableEdgeSelect = Extension.create({
  name: 'tableEdgeSelect',
  addProseMirrorPlugins() {
    return [edgeSelectPlugin()]
  },
})
