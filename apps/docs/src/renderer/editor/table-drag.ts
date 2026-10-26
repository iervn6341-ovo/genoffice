import { createTableDragGuard } from '@genoffice/ui'

/**
 * Drag safety inside Docs tables: no native drag of a cell-range selection, no cross-cell text
 * move, no row/cell slice dropped into a table. Everything else (moving text within a paragraph,
 * between paragraphs, into or out of a table) keeps the Word behaviour. See
 * `packages/ui/src/table-drag-guard.ts`.
 */
export const tableDragGuard = createTableDragGuard({
  cellNames: ['docTableCell', 'docTableHeader'],
  rowNames: ['docTableRow'],
})
