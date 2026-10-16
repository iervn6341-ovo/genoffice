import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'

import { applyCellEditsToXlsx } from '@genoffice/xlsx-gateway/gateway/xlsx-gateway'
import { StylesheetEditor } from '@genoffice/xlsx-gateway/gateway/xlsx-styles'
import {
  fontEffectDeltaOfCommand,
  restoreProtectionSnapshot,
} from '../src/renderer/cell-protection-undo'
import { createEditJournal, fromNeutralStyle, toSaveEdits } from '../src/renderer/edit-journal'
import { draftFromSelection, formatCellsCommands } from '../src/renderer/format-cells'
import { selectionProtection } from '../src/renderer/selection-protection'
import { toUniverStyle } from '../src/renderer/univer-sync'
import type { WorkbookFile } from '../src/shared/desktop-api'
import { buildEditFixture } from './fixture-builder'

const BASE =
  '<styleSheet><fonts count="1"><font><name val="Aptos"/><family val="2"/></font></fonts><fills count="1"><fill/></fills><borders count="1"><border/></borders><cellXfs count="1"><xf fontId="0" custom="keep"><extLst><ext uri="vendor"/></extLst></xf></cellXfs></styleSheet>'

describe('native cell font effects', () => {
  it('writes all four underline values and both script positions while keeping unknown font/XF XML', () => {
    for (const mode of ['single', 'double', 'singleAccounting', 'doubleAccounting'] as const) {
      const editor = new StylesheetEditor(BASE)
      editor.resolveStyle(0, {
        underline: true,
        underlineStyle: mode,
        fontVerticalAlign: 'subscript',
      })
      const xml = editor.serialize()
      expect(xml).toContain(mode === 'single' ? '<u/>' : `<u val="${mode}"/>`)
      expect(xml).toContain('<vertAlign val="subscript"/>')
      expect(xml).toContain('<family val="2"/>')
      const xfs = /<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/.exec(xml)?.[1] ?? ''
      expect(xfs).toContain(
        '<xf fontId="1" custom="keep" applyFont="1"><extLst><ext uri="vendor"/></extLst></xf>',
      )
    }
    const editor = new StylesheetEditor(BASE)
    editor.resolveStyle(0, { underline: false, fontVerticalAlign: 'superscript' })
    expect(editor.serialize()).toContain('<vertAlign val="superscript"/>')
    expect(
      fromNeutralStyle({
        underline: true,
        underlineStyle: 'doubleAccounting',
        fontVerticalAlign: 'subscript',
      }),
    ).toMatchObject({
      ul: { s: 1, t: 10 },
      va: 2,
    })
    expect(
      toUniverStyle({
        bold: false,
        italic: false,
        underline: true,
        underlineStyle: 'singleAccounting',
        fontVerticalAlign: 'superscript',
        strikethrough: false,
        wrapText: false,
        diagonalUp: false,
        diagonalDown: false,
      }),
    ).toMatchObject({ ul: { s: 1 }, va: 3 })
  })

  it('backfills mixed native effects and emits only the changed effect', () => {
    const base = {
      bold: false,
      italic: false,
      underline: false,
      strikethrough: false,
      wrapText: false,
      diagonalUp: false,
      diagonalDown: false,
    }
    const styles = [
      base,
      {
        ...base,
        underline: true,
        underlineStyle: 'singleAccounting' as const,
        fontVerticalAlign: 'superscript' as const,
      },
      {
        ...base,
        underline: true,
        underlineStyle: 'doubleAccounting' as const,
        fontVerticalAlign: 'subscript' as const,
      },
    ]
    const read = {
      cells: [
        { row: 0, column: 0, value: 'A', styleIndex: 1 },
        { row: 0, column: 1, value: 3, formula: '=1+2', styleIndex: 2 },
      ],
      rows: [],
    }
    const range = { startRow: 0, endRow: 0, startColumn: 0, endColumn: 1 }
    const selection = selectionProtection(
      styles,
      { columnWidths: [] } as unknown as WorkbookFile['sheets'][number],
      read,
      createEditJournal(),
      'sheet-1',
      range,
    )
    expect(selection.font).toEqual({ underlineStyle: null, verticalAlign: null })
    const draft = draftFromSelection(null, selection)
    expect(draft.underlineMode).toBe('')
    expect(draft.fontVerticalAlign).toBe('')
    expect(formatCellsCommands(draft, { ...draft, underlineMode: 'doubleAccounting' })).toEqual([
      'font-effect:doubleAccounting:',
    ])
    expect(fontEffectDeltaOfCommand('font-effect::baseline')).toEqual({
      fontVerticalAlign: 'baseline',
    })
    expect(fontEffectDeltaOfCommand('font-effect:doubleAccounting:subscript')).toEqual({
      underline: true,
      underlineStyle: 'doubleAccounting',
      fontVerticalAlign: 'subscript',
    })
  })

  it('restores the exact original XF after Save → Undo without changing formula/cache', async () => {
    const original = await buildEditFixture()
    const edited = await applyCellEditsToXlsx(original, [
      {
        sheetName: 'Data',
        row: 0,
        column: 0,
        writeValue: false,
        cell: { value: null },
        style: {
          underline: true,
          underlineStyle: 'singleAccounting',
          fontVerticalAlign: 'superscript',
        },
      },
    ])
    const journal = createEditJournal()
    restoreProtectionSnapshot(
      { file: { sessionId: 'after-save' } as WorkbookFile, editJournal: journal },
      'sheet-1',
      {
        sessionId: 'before-save',
        cells: [{ row: 0, column: 0, styleReference: 1, locked: true, hidden: false, entry: null }],
      },
      {
        sessionId: 'before-save',
        cells: [{ row: 0, column: 0, styleReference: 1, locked: true, hidden: false, entry: null }],
      },
    )
    const undone = await applyCellEditsToXlsx(
      edited.buffer,
      toSaveEdits(journal).map((edit) => ({
        sheetName: 'Data',
        row: edit.row,
        column: edit.column,
        writeValue: edit.writeValue,
        cell: { value: edit.value },
        styleReference: edit.styleReference,
      })),
    )
    const before = await (
      await JSZip.loadAsync(original)
    )
      .file('xl/worksheets/sheet1.xml')!
      .async('string')
    const after = await (
      await JSZip.loadAsync(undone.buffer)
    )
      .file('xl/worksheets/sheet1.xml')!
      .async('string')
    expect(/<c r="A1"[^>]*>[\s\S]*?<\/c>/.exec(after)?.[0]).toBe(
      /<c r="A1"[^>]*>[\s\S]*?<\/c>/.exec(before)?.[0],
    )
    expect(after).toContain('<f>SUM(C1:C2)</f><v>5</v>')
  })
})
