/**
 * Drag-and-drop guard for ProseMirror-style editors that contain tables (Docs, Markdown).
 *
 * Why: pressing on already-selected text and dragging starts a native drag. Inside a table that
 * is rarely intended — the user is usually extending a cell selection with the mouse — and the
 * default handling either moves the text into another cell, or, for a cell-range selection,
 * lifts the whole row/cell slice and pastes it into the target cell, which tears the table apart.
 *
 * What it blocks (and nothing else — dragging text inside one paragraph, between paragraphs, or
 * into / out of a table still behaves like Word):
 *  1. `dragstart` while the selection is a cell-range selection;
 *  2. a drop whose slice carries table structure (rows / cells) onto a spot inside a table;
 *  3. a drop that MOVES text out of one table cell into a different cell of a table.
 * A blocked drop changes nothing, so there is no undo step to clean up.
 *
 * The helper is structural (no prosemirror import) so `@genoffice/ui` stays dependency-free;
 * pass it the editor's own node names.
 */

interface NodeTypeLike {
  name: string
}
interface NodeLike {
  type: NodeTypeLike
}
interface FragmentLike {
  childCount: number
  firstChild: { content: FragmentLike } | null
  descendants(f: (node: NodeLike) => boolean | void): void
}
interface SliceLike {
  content: FragmentLike
  openStart: number
  openEnd: number
}
interface ResolvedPosLike {
  depth: number
  node(depth: number): NodeLike
  before(depth: number): number
}
interface SelectionLike {
  $from: ResolvedPosLike
  $to: ResolvedPosLike
}
export interface DragGuardViewLike {
  state: {
    selection: SelectionLike
    doc: { resolve(pos: number): ResolvedPosLike }
  }
  posAtCoords(coords: { left: number; top: number }): { pos: number; inside: number } | null
}

export interface TableDragGuardOptions {
  /** Cell node type names (data cells and header cells) */
  cellNames: readonly string[]
  /** Row node type names — a slice holding any of rows / cells counts as "table structure" */
  rowNames: readonly string[]
}

export interface TableDragGuard {
  /** `handleDOMEvents.dragstart`: true (and cancels the native drag) for a cell-range selection */
  dragstart(view: DragGuardViewLike, event: DragEvent): boolean
  /** `handleDrop`: true when the drop must be swallowed */
  drop(view: DragGuardViewLike, event: DragEvent, slice: SliceLike | null, moved: boolean): boolean
}

/** A cell-range selection (prosemirror-tables' CellSelection) — recognised by its anchor/head cells */
export function isCellRangeSelection(selection: unknown): boolean {
  if (typeof selection !== 'object' || selection === null) return false
  const s = selection as { $anchorCell?: unknown; $headCell?: unknown }
  return s.$anchorCell !== undefined && s.$headCell !== undefined
}

/** Document position of the cell containing `$pos`, or null when it is not inside a cell */
export function enclosingCellPos(
  $pos: ResolvedPosLike,
  cellNames: readonly string[],
): number | null {
  for (let d = $pos.depth; d > 0; d--) {
    if (cellNames.includes($pos.node(d).type.name)) return $pos.before(d)
  }
  return null
}

/**
 * Does the slice carry rows or cells (a dragged cell range)? A slice of plain text taken inside a
 * cell is wrapped in the open parents (table > row > cell > paragraph); those open levels are not
 * dragged structure, so they are peeled off first.
 */
export function sliceHasTableStructure(
  slice: SliceLike | null,
  structureNames: readonly string[],
): boolean {
  if (!slice) return false
  let content = slice.content
  let open = Math.min(slice.openStart, slice.openEnd)
  while (open > 0 && content.childCount === 1 && content.firstChild) {
    content = content.firstChild.content
    open--
  }
  let found = false
  content.descendants((node) => {
    if (found) return false
    if (structureNames.includes(node.type.name)) found = true
    return !found
  })
  return found
}

export function createTableDragGuard(options: TableDragGuardOptions): TableDragGuard {
  const structure = [...options.cellNames, ...options.rowNames]
  return {
    dragstart(view, event) {
      if (!isCellRangeSelection(view.state.selection)) return false
      event.preventDefault()
      return true
    },
    drop(view, event, slice, moved) {
      const hit = view.posAtCoords({ left: event.clientX, top: event.clientY })
      if (!hit) return false
      const targetCell = enclosingCellPos(view.state.doc.resolve(hit.pos), options.cellNames)
      if (targetCell === null) return false
      if (sliceHasTableStructure(slice, structure)) {
        event.preventDefault()
        return true
      }
      if (moved) {
        // the dragged selection is still the live selection during a drop
        const sourceCell = enclosingCellPos(view.state.selection.$from, options.cellNames)
        if (sourceCell !== null && sourceCell !== targetCell) {
          event.preventDefault()
          return true
        }
      }
      return false
    },
  }
}
