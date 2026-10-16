import { Fragment, type Node as PmNode, type NodeType } from '@tiptap/pm/model'
import type { Command, EditorState, Transaction } from '@tiptap/pm/state'
import { TextSelection } from '@tiptap/pm/state'
import { CellSelection } from '@tiptap/pm/tables'
import type { EditorView } from '@tiptap/pm/view'
import { selectedTable } from './table-properties'
import { tableModelToPmNode, type PmNode as JsonNode } from './convert'

/**
 * Word Table Layout / Table Design extras that are not covered by the basic row, column,
 * merge and AutoFit commands: alternative text, cell spacing, row break rule, text
 * direction, distribute rows, sort, convert table <-> text and selection helpers. Every
 * command edits native table attributes (with the matching `*Edited` flag) so the edit
 * survives Undo/Redo and is written to the DOCX table parts on save.
 */

/* ------------------------------------------------------------ table attrs */

/** Table Properties > Alt Text (w:tblCaption / w:tblDescription). Empty strings clear them. */
export function setTableAltText(title: string, description: string): Command {
  return (state, dispatch) => {
    const selected = selectedTable(state)
    if (!selected) return false
    dispatch?.(
      state.tr.setNodeMarkup(selected.pos, undefined, {
        ...selected.node.attrs,
        altTitle: title.trim() || null,
        altDescription: description.trim() || null,
        altEdited: true,
      }),
    )
    return true
  }
}

/** Table Properties > Options > Allow spacing between cells (twips of w:tblCellSpacing; 0/null clears). */
export function setTableCellSpacing(twips: number | null): Command {
  return (state, dispatch) => {
    const selected = selectedTable(state)
    if (!selected) return false
    const value = twips && twips > 0 ? Math.min(31680, Math.round(twips)) : null
    dispatch?.(
      state.tr.setNodeMarkup(selected.pos, undefined, {
        ...selected.node.attrs,
        cellSpacingTwips: value,
        cellSpacingEdited: true,
      }),
    )
    return true
  }
}

/* ------------------------------------------------------------- row / cell */

/** Rows covered by the selection (the caret's row for a collapsed selection). */
function selectedRowRange(state: EditorState): { top: number; bottom: number } | null {
  const selected = selectedTable(state)
  if (!selected) return null
  return { top: selected.rect.top, bottom: selected.rect.bottom }
}

export function rowsMayBreak(state: EditorState): boolean | null {
  const selected = selectedTable(state)
  const range = selectedRowRange(state)
  if (!selected || !range) return null
  let allow = true
  for (let row = range.top; row < range.bottom; row++) {
    if (selected.node.child(row).attrs.cantSplit) allow = false
  }
  return allow
}

/** Table Properties > Row > Allow row to break across pages (w:cantSplit when off). */
export function setRowsMayBreak(allow: boolean): Command {
  return (state, dispatch) => {
    const selected = selectedTable(state)
    const range = selectedRowRange(state)
    if (!selected || !range) return false
    let tr = state.tr
    selected.node.forEach((row, offset, index) => {
      if (index < range.top || index >= range.bottom) return
      tr = tr.setNodeMarkup(selected.tableStart + offset, undefined, {
        ...row.attrs,
        cantSplit: !allow,
        cantSplitEdited: true,
      })
    })
    dispatch?.(tr)
    return true
  }
}

export type CellTextDirection = null | 'tbRl' | 'btLr'

/** Cells covered by the selection, as absolute positions. */
function selectedCellPositions(state: EditorState): { pos: number; node: PmNode }[] {
  const selected = selectedTable(state)
  if (!selected) return []
  const { rect, tableStart } = selected
  const seen = new Set<number>()
  const out: { pos: number; node: PmNode }[] = []
  for (const relative of rect.map.cellsInRect(rect)) {
    if (seen.has(relative)) continue
    seen.add(relative)
    const node = state.doc.nodeAt(tableStart + relative)
    if (node) out.push({ pos: tableStart + relative, node })
  }
  return out
}

export function activeCellTextDirection(state: EditorState): CellTextDirection {
  const first = selectedCellPositions(state)[0]
  const value = first?.node.attrs.textDirection
  return value === 'tbRl' || value === 'btLr' ? value : null
}

/** Table Layout > Text Direction. Cycles horizontal -> rotate 90 -> rotate 270 when `next` is omitted. */
export function setCellTextDirection(next?: CellTextDirection): Command {
  return (state, dispatch) => {
    const cells = selectedCellPositions(state)
    if (cells.length === 0) return false
    const current = activeCellTextDirection(state)
    const target: CellTextDirection =
      next !== undefined ? next : current === null ? 'tbRl' : current === 'tbRl' ? 'btLr' : null
    let tr = state.tr
    for (const { pos, node } of cells) {
      tr = tr.setNodeMarkup(pos, undefined, {
        ...node.attrs,
        textDirection: target,
        textDirectionEdited: true,
      })
    }
    dispatch?.(tr)
    return true
  }
}

/**
 * Table Layout > Distribute Rows: selected rows (the whole table for fewer than two) take their
 * average rendered height. `zoom` is the canvas zoom factor so unzoomed pixels are written back.
 */
export function distributeSelectedRows(view: EditorView, zoom = 1): boolean {
  const state = view.state
  const selected = selectedTable(state)
  if (!selected) return false
  const rows: { offset: number; height: number }[] = []
  let { top, bottom } = selected.rect
  if (bottom - top < 2) {
    top = 0
    bottom = selected.node.childCount
  }
  if (bottom - top < 2) return false
  selected.node.forEach((_row, offset, index) => {
    if (index < top || index >= bottom) return
    const dom = view.nodeDOM(selected.tableStart + offset) as HTMLElement | null
    const px = (dom?.getBoundingClientRect().height ?? 0) / (zoom || 1)
    rows.push({ offset, height: px })
  })
  if (rows.length < 2 || rows.some((row) => !(row.height > 0))) return false
  const average = rows.reduce((sum, row) => sum + row.height, 0) / rows.length
  const twips = Math.max(1, Math.round(average * 15))
  let tr = state.tr
  for (const { offset } of rows) {
    const pos = selected.tableStart + offset
    const row = state.doc.nodeAt(pos)
    if (!row) continue
    tr = tr.setNodeMarkup(pos, undefined, {
      ...row.attrs,
      heightTwips: twips,
      heightRule: 'atLeast',
    })
  }
  view.dispatch(tr)
  return true
}

/* ------------------------------------------------------------------- sort */

export type SortKind = 'text' | 'number' | 'date'

export interface SortKey {
  /** zero-based grid column */
  column: number
  kind: SortKind
  descending: boolean
}

const collator = new Intl.Collator(undefined, { sensitivity: 'base' })

export function numberOf(text: string): number | null {
  const cleaned = text.replace(/[,\s\u00a0$€£¥%]/g, '').replace(/^\((.*)\)$/, '-$1')
  if (!/^[-+]?(\d+\.?\d*|\.\d+)(e[-+]?\d+)?$/i.test(cleaned)) return null
  const value = Number(cleaned)
  return Number.isFinite(value) ? value : null
}

function dateOf(text: string): number | null {
  const trimmed = text.trim()
  if (!trimmed) return null
  const value = Date.parse(trimmed.replace(/\./g, '/'))
  return Number.isNaN(value) ? null : value
}

/** Comparison of two cell texts; empty / unparsable values sort first, like Word's ascending order. */
export function compareCellText(a: string, b: string, kind: SortKind): number {
  if (kind === 'text') return collator.compare(a.trim(), b.trim())
  const parse = kind === 'number' ? numberOf : dateOf
  const left = parse(a)
  const right = parse(b)
  if (left === null && right === null) return collator.compare(a.trim(), b.trim())
  if (left === null) return -1
  if (right === null) return 1
  return left - right
}

export interface SortOptions {
  keys: SortKey[]
  /** first row is a header that stays on top (Word "My list has: Header row") */
  header: boolean
}

function plainRowSelection(state: EditorState): boolean {
  const selected = selectedTable(state)
  if (!selected) return false
  let plain = true
  selected.node.descendants((node) => {
    if (node.type.spec.tableRole === 'cell' || node.type.spec.tableRole === 'header_cell') {
      if ((node.attrs.colspan ?? 1) > 1 || (node.attrs.rowspan ?? 1) > 1) plain = false
      return false
    }
    return true
  })
  return plain
}

/** True when Sort can run: a table without merged cells holding the caret/selection. */
export function canSortTable(state: EditorState): boolean {
  return plainRowSelection(state)
}

/**
 * Table Layout > Sort. Sorts the selected rows (the whole table when fewer than two rows are
 * selected), optionally keeping the header row in place. Row properties travel with their row;
 * cell node kinds (header/data) stay by position. Refuses tables with merged cells.
 */
export function sortTableRows(options: SortOptions): Command {
  return (state, dispatch) => {
    if (!plainRowSelection(state)) return false
    const selected = selectedTable(state)
    if (!selected || options.keys.length === 0) return false
    const { node: table, tableStart } = selected
    let first = selected.rect.top
    let last = selected.rect.bottom
    if (last - first < 2) {
      first = 0
      last = table.childCount
    }
    if (options.header && first === 0) first = 1
    if (last - first < 2) return false
    const rows: PmNode[] = []
    const offsets: number[] = []
    table.forEach((row, offset, index) => {
      if (index >= first && index < last) {
        rows.push(row)
        offsets.push(offset)
      }
    })
    const textAt = (row: PmNode, column: number): string =>
      row.maybeChild(column)?.textContent ?? ''
    const order = rows
      .map((row, index) => ({ row, index }))
      .sort((a, b) => {
        for (const key of options.keys) {
          const result = compareCellText(
            textAt(a.row, key.column),
            textAt(b.row, key.column),
            key.kind,
          )
          if (result !== 0) return key.descending ? -result : result
        }
        return a.index - b.index
      })
    if (order.every((entry, index) => entry.index === index)) return true
    // cell kinds follow the slot, not the moved row
    const kindsOf = (row: PmNode): NodeType[] => {
      const kinds: NodeType[] = []
      row.forEach((cell) => kinds.push(cell.type))
      return kinds
    }
    const slotKinds = rows.map(kindsOf)
    const rebuilt = order.map(({ row }, slot) => {
      const cells: PmNode[] = []
      row.forEach((cell, _offset, column) => {
        const kind = slotKinds[slot][column] ?? cell.type
        cells.push(kind === cell.type ? cell : kind.create(cell.attrs, cell.content, cell.marks))
      })
      return row.type.create(row.attrs, cells, row.marks)
    })
    const from = tableStart + offsets[0]
    const to = tableStart + offsets[offsets.length - 1] + rows[rows.length - 1].nodeSize
    if (dispatch) {
      const tr = state.tr.replaceWith(from, to, Fragment.from(rebuilt))
      tr.setSelection(TextSelection.near(tr.doc.resolve(from + 2)))
      dispatch(tr)
    }
    return true
  }
}

/* ------------------------------------------------- convert table <-> text */

export type TextSeparator = 'paragraph' | 'tab' | 'comma' | { custom: string }

const separatorText = (separator: TextSeparator): string | null =>
  separator === 'paragraph'
    ? null
    : separator === 'tab'
      ? '\t'
      : separator === 'comma'
        ? ','
        : separator.custom

function plainParagraph(node: PmNode, content: Fragment): PmNode {
  return node.type.create({ ...node.attrs, docxIndex: null }, content)
}

/**
 * Table Layout > Convert to Text. `paragraph` writes every cell paragraph as its own paragraph;
 * other separators write one paragraph per row with the cells joined by the separator. Tables with
 * nested tables or anchored boxes in cells are refused (nothing changes).
 */
export function convertTableToText(separator: TextSeparator): Command {
  return (state, dispatch) => {
    const selected = selectedTable(state)
    if (!selected) return false
    const sepText = separatorText(separator)
    if (sepText === '' || (typeof separator === 'object' && !separator.custom)) return false
    const out: PmNode[] = []
    let supported = true
    selected.node.forEach((row) => {
      const cellParagraphs: PmNode[][] = []
      row.forEach((cell) => {
        if (cell.attrs.gridGap) return
        const paragraphs: PmNode[] = []
        cell.forEach((child) => {
          if (child.type.name === 'docParagraph' || child.type.name === 'docListItem')
            paragraphs.push(child)
          else supported = false
        })
        cellParagraphs.push(paragraphs)
      })
      if (sepText === null) {
        for (const paragraphs of cellParagraphs)
          for (const paragraph of paragraphs) out.push(plainParagraph(paragraph, paragraph.content))
        return
      }
      const schema = state.schema
      const parts: PmNode[] = []
      let template: PmNode | null = null
      cellParagraphs.forEach((paragraphs, index) => {
        if (index > 0) parts.push(schema.text(sepText))
        paragraphs.forEach((paragraph, at) => {
          if (!template) template = paragraph
          if (at > 0) parts.push(schema.text(' '))
          paragraph.content.forEach((child) => parts.push(child))
        })
      })
      out.push(
        plainParagraph(
          (template as PmNode | null) ?? state.schema.nodes.docParagraph.create(null),
          Fragment.fromArray(parts),
        ),
      )
    })
    if (!supported || out.length === 0) return false
    if (dispatch) {
      const from = selected.pos
      const to = from + selected.node.nodeSize
      const tr = state.tr.replaceWith(from, to, Fragment.fromArray(out))
      tr.setSelection(TextSelection.near(tr.doc.resolve(from)))
      dispatch(tr)
    }
    return true
  }
}

const SINGLE_LINE = { style: 'single', szEighths: 4, color: 'auto' }

/** Split a textblock's inline content at every occurrence of `sep` (marks are kept). */
export function splitParagraphAt(paragraph: PmNode, sep: string): Fragment[] {
  const cuts: { from: number; to: number }[] = []
  paragraph.descendants((node, pos) => {
    if (!node.isText || !node.text) return true
    let at = node.text.indexOf(sep)
    while (at >= 0) {
      cuts.push({ from: pos + at, to: pos + at + sep.length })
      at = node.text.indexOf(sep, at + sep.length)
    }
    return true
  })
  const pieces: Fragment[] = []
  let cursor = 0
  for (const cut of cuts) {
    pieces.push(paragraph.content.cut(cursor, cut.from))
    cursor = cut.to
  }
  pieces.push(paragraph.content.cut(cursor, paragraph.content.size))
  return pieces
}

export interface TextToTableOptions {
  separator: TextSeparator
  /** columns for the `paragraph` separator (1 when omitted) */
  columns?: number
}

/**
 * Insert > Table > Convert Text to Table. The selected top-level paragraphs become the rows
 * (tab / comma / custom character) or the cells (paragraph separator filled row by row).
 * Returns false when the selection is not made of plain top-level paragraphs.
 */
export function convertTextToTable(options: TextToTableOptions): Command {
  return (state, dispatch) => {
    const { $from, $to } = state.selection
    if (state.selection instanceof CellSelection || $from.depth < 1) return false
    const startIndex = $from.index(0)
    const endIndex = $to.index(0)
    const blocks: PmNode[] = []
    for (let index = startIndex; index <= endIndex; index++) {
      const block = state.doc.child(index)
      if (block.type.name !== 'docParagraph' && block.type.name !== 'docListItem') return false
      blocks.push(block)
    }
    const sepText = separatorText(options.separator)
    if (typeof options.separator === 'object' && !options.separator.custom) return false
    let grid: Fragment[][]
    if (sepText === null) {
      const columns = Math.max(1, Math.min(63, Math.round(options.columns ?? 1)))
      grid = []
      blocks.forEach((block, index) => {
        if (index % columns === 0) grid.push([])
        grid[grid.length - 1].push(block.content)
      })
    } else {
      grid = blocks.map((block) => splitParagraphAt(block, sepText))
    }
    const columnCount = Math.max(1, ...grid.map((row) => row.length))
    if (columnCount > 63) return false
    const template = blocks[0]
    const model = {
      rows: grid.map(() => Array.from({ length: columnCount }, () => ({ paras: [''] }))),
      colWidthsPct: Array.from({ length: columnCount }, () => 100 / columnCount),
      widthPct: 100,
      autoFit: 'window' as const,
      borders: {
        top: SINGLE_LINE,
        bottom: SINGLE_LINE,
        left: SINGLE_LINE,
        right: SINGLE_LINE,
        insideH: SINGLE_LINE,
        insideV: SINGLE_LINE,
      },
    }
    const json = tableModelToPmNode(model) as JsonNode
    const tableNode = state.schema.nodeFromJSON(json)
    // swap each empty cell paragraph for the converted content (formatting of the first source paragraph is kept)
    const rowNodes: PmNode[] = []
    tableNode.forEach((row, _o, rowIndex) => {
      const cells: PmNode[] = []
      row.forEach((cell, _co, columnIndex) => {
        const content = grid[rowIndex][columnIndex] ?? Fragment.empty
        const paragraph = state.schema.nodes.docParagraph.create(
          { ...(template.attrs as Record<string, unknown>), docxIndex: null },
          content,
        )
        cells.push(cell.type.create(cell.attrs, [paragraph], cell.marks))
      })
      rowNodes.push(row.type.create(row.attrs, cells, row.marks))
    })
    const finalTable = tableNode.type.create(tableNode.attrs, rowNodes, tableNode.marks)
    if (dispatch) {
      const from = $from.before(1)
      const last = state.doc.child(endIndex)
      const to = $to.before(1) + last.nodeSize
      const atEnd = to >= state.doc.content.size
      const nodes: PmNode[] = [finalTable]
      if (atEnd) nodes.push(state.schema.nodes.docParagraph.create())
      const tr: Transaction = state.tr.replaceWith(from, to, Fragment.fromArray(nodes))
      tr.setSelection(TextSelection.near(tr.doc.resolve(from + 3)))
      dispatch(tr)
    }
    return true
  }
}

/* ------------------------------------------------------------- selections */

export type TableSelectionTarget = 'table' | 'row' | 'column' | 'cell'

/** Table Layout > Select: expands the caret to a cell, row, column or the whole table. */
export function selectInTable(target: TableSelectionTarget): Command {
  return (state, dispatch) => {
    const selected = selectedTable(state)
    if (!selected) return false
    const { rect, tableStart } = selected
    const map = rect.map
    const at = (row: number, column: number) => tableStart + map.map[row * map.width + column]
    let anchor: number
    let head: number
    if (target === 'cell') {
      anchor = head = at(rect.top, rect.left)
    } else if (target === 'row') {
      anchor = at(rect.top, 0)
      head = at(rect.bottom - 1, map.width - 1)
    } else if (target === 'column') {
      anchor = at(0, rect.left)
      head = at(map.height - 1, rect.right - 1)
    } else {
      anchor = at(0, 0)
      head = at(map.height - 1, map.width - 1)
    }
    if (dispatch) {
      const tr = state.tr
      tr.setSelection(CellSelection.create(tr.doc, anchor, head))
      dispatch(tr)
    }
    return true
  }
}
