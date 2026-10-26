import { Extension } from '@tiptap/core'
import { Plugin, PluginKey } from '@tiptap/pm/state'
import { createTableDragGuard } from '@genoffice/ui'

/**
 * Drag safety inside Markdown tables: pressing on selected cell text and dragging used to move the
 * text into another cell, and dragging a cell-range selection tore the table apart (GFM save then
 * wrote a broken table). Blocks only those cases; paragraph-level text dragging is untouched.
 */
export const tableDragGuard = createTableDragGuard({
  cellNames: ['tableCell', 'tableHeader'],
  rowNames: ['tableRow'],
})

export const TableDragGuard = Extension.create({
  name: 'tableDragGuard',
  // ahead of the image-drop handler so a blocked table drop is never reinterpreted
  priority: 1100,
  addProseMirrorPlugins() {
    return [
      new Plugin({
        key: new PluginKey('tableDragGuard'),
        props: {
          handleDOMEvents: {
            dragstart: (view, event) => tableDragGuard.dragstart(view, event),
          },
          handleDrop: (view, event, slice, moved) => tableDragGuard.drop(view, event, slice, moved),
        },
      }),
    ]
  },
})
