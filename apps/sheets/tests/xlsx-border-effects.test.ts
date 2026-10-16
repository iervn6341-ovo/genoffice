import JSZip from 'jszip'
import { readFile } from 'node:fs/promises'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'

import { applyCellEditsToXlsx } from '@genoffice/xlsx-gateway/gateway/xlsx-gateway'
import { StylesheetEditor } from '@genoffice/xlsx-gateway/gateway/xlsx-styles'
import {
  borderActionCommand,
  borderActionOfCommand,
  borderDeltaForCell,
} from '../src/renderer/border-format'
import { restoreProtectionSnapshot } from '../src/renderer/cell-protection-undo'
import { createEditJournal, toSaveEdits } from '../src/renderer/edit-journal'
import { draftFromSelection } from '../src/renderer/format-cells'
import { selectionBorderState, selectionProtection } from '../src/renderer/selection-protection'
import type { CellProtectionFlags, ProtectionRange } from '../src/renderer/selection-protection'
import type { WorkbookFile } from '../src/shared/desktop-api'
import { buildEditFixture } from './fixture-builder'

const REPO_ROOT = join(fileURLToPath(new URL('.', import.meta.url)), '../../..')

const BASE =
  '<styleSheet><fonts count="1"><font/></fonts><fills count="1"><fill/></fills>' +
  '<borders count="1"><border custom="keep>a" diagonalDown="1"><left style="medium"><color rgb="FF102030"/></left><diagonal style="hair"/><horizontal style="dotted"/><vertical style="dashed"/><extLst><ext uri="vendor"/></extLst></border></borders>' +
  '<cellXfs count="1"><xf fontId="0" borderId="0" custom="xf"><extLst><ext uri="xf-vendor"/></extLst></xf></cellXfs></styleSheet>'

describe('native border effects', () => {
  it('retains Excel Mac XF extensions through Border + Alignment/Fill/Number/Protection and Undo/Redo', async () => {
    const sourceZip = await JSZip.loadAsync(
      await readFile(
        join(
          REPO_ROOT,
          'docs/qa/office-mac-field-evidence-2026-10-13/sheets-border-effects/excel-mac-border-source.xlsx',
        ),
      ),
    )
    const originalStyles = await sourceZip.file('xl/styles.xml')!.async('string')
    const originalXf =
      '<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"/>'
    expect(originalStyles).toContain(originalXf)
    sourceZip.file(
      'xl/styles.xml',
      originalStyles.replace(
        originalXf,
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" quotePrefix="1"><extLst><ext uri="urn:vendor:keep"/></extLst></xf>',
      ),
    )
    const source = await sourceZip.generateAsync({ type: 'nodebuffer' })
    const edits = [
      {
        sheetName: 'Sheet1',
        row: 0,
        column: 0,
        writeValue: false,
        cell: { value: null },
        style: {
          borderTop: { style: 'double' as const, color: '#102030' },
          verticalAlignment: 'distributed' as const,
        },
      },
      {
        sheetName: 'Sheet1',
        row: 0,
        column: 1,
        writeValue: false,
        cell: { value: null },
        style: {
          borderBottom: { style: 'dashed' as const, color: '#AABBCC' },
          fillColor: '#ABCDEF',
        },
      },
      {
        sheetName: 'Sheet1',
        row: 1,
        column: 0,
        writeValue: false,
        cell: { value: null },
        style: {
          borderLeft: { style: 'double' as const, color: '#667788' },
          numberFormat: '0.00',
        },
      },
      {
        sheetName: 'Sheet1',
        row: 1,
        column: 1,
        writeValue: false,
        cell: { value: null },
        style: {
          borderRight: { style: 'dashed' as const, color: '#112233' },
          protectionLocked: false,
          protectionHidden: true,
        },
      },
    ]
    const edited = await applyCellEditsToXlsx(source, edits)
    const assertEdited = async (buffer: Uint8Array) => {
      const zip = await JSZip.loadAsync(buffer)
      const styles = await zip.file('xl/styles.xml')!.async('string')
      const sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
      const xfs = [
        ...(/<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/.exec(styles)?.[1] ?? '').matchAll(
          /<xf\b[^>]*\/>|<xf\b[^>]*>[\s\S]*?<\/xf>/g,
        ),
      ].map((match) => match[0])
      const xfOf = (address: string) => {
        const cell = new RegExp(`<c\\b[^>]*r="${address}"[^>]*>`).exec(sheet)?.[0] ?? ''
        return xfs[Number(/\bs="(\d+)"/.exec(cell)?.[1] ?? 0)] ?? ''
      }
      for (const address of ['A1', 'B1', 'A2', 'B2']) {
        expect(xfOf(address)).toContain('quotePrefix="1"')
        expect(xfOf(address)).toContain('<extLst><ext uri="urn:vendor:keep"/></extLst>')
      }
      expect(xfOf('A1')).toContain('vertical="distributed"')
      expect(xfOf('A1')).toContain('applyBorder="1"')
      expect(xfOf('B1')).toContain('applyFill="1"')
      expect(xfOf('B1')).toContain('applyBorder="1"')
      expect(xfOf('A2')).toContain('numFmtId="2"')
      expect(xfOf('A2')).toContain('applyNumberFormat="1"')
      expect(xfOf('B2')).toContain('applyProtection="1"')
      expect(xfOf('B2')).toContain('locked="0"')
      expect(xfOf('B2')).toContain('hidden="1"')
      expect(styles).toContain('x14ac:knownFonts="1"')
      expect(sheet).toContain('<c r="A1" s="')
      expect(sheet).toContain('<v>1</v>')
      expect(sheet).toContain('<c r="B1" s="')
      expect(sheet).toContain('<v>2</v>')
      expect(sheet).toContain('<v>3</v>')
      expect(sheet).toContain('<v>4</v>')
    }
    await assertEdited(edited.buffer)
    const undone = await applyCellEditsToXlsx(edited.buffer, [
      {
        sheetName: 'Sheet1',
        row: 0,
        column: 0,
        writeValue: false,
        cell: { value: null },
        styleReference: 1,
      },
      {
        sheetName: 'Sheet1',
        row: 0,
        column: 1,
        writeValue: false,
        cell: { value: null },
        styleReference: 1,
      },
      {
        sheetName: 'Sheet1',
        row: 1,
        column: 0,
        writeValue: false,
        cell: { value: null },
        styleReference: 1,
      },
      {
        sheetName: 'Sheet1',
        row: 1,
        column: 1,
        writeValue: false,
        cell: { value: null },
        styleReference: 1,
      },
    ])
    const undoneSheet = await (
      await JSZip.loadAsync(undone.buffer)
    )
      .file('xl/worksheets/sheet1.xml')!
      .async('string')
    expect(undoneSheet).toContain('<c r="A1" s="1"><v>1</v></c>')
    expect(undoneSheet).toContain('<c r="B1" s="1"><v>2</v></c>')
    expect(undoneSheet).toContain('<c r="A2" s="1"><v>3</v></c>')
    expect(undoneSheet).toContain('<c r="B2" s="1"><v>4</v></c>')
    const redone = await applyCellEditsToXlsx(undone.buffer, edits)
    await assertEdited(redone.buffer)
  })

  it('round-trips the Excel Mac dotted border and auto color when adding the second diagonal', async () => {
    const source = await readFile(
      join(
        REPO_ROOT,
        'docs/qa/office-mac-field-evidence-2026-10-13/sheets-border-effects/excel-mac-border-source.xlsx',
      ),
    )
    const range: ProtectionRange = {
      startRow: 0,
      endRow: 0,
      startColumn: 0,
      endColumn: 0,
    }
    const delta = borderDeltaForCell(
      { modes: { diagonalDown: true }, style: 'dotted', color: '#000000' },
      range,
      {
        row: 0,
        column: 0,
        locked: true,
        hidden: false,
        horizontalAlignment: 'general',
        verticalAlignment: 'bottom',
        textRotation: 0,
        shrinkToFit: false,
        diagonalUp: true,
        diagonalDown: false,
      },
    )
    expect(delta).toEqual({ diagonalDown: true })
    const saved = await applyCellEditsToXlsx(source, [
      {
        sheetName: 'Sheet1',
        row: 0,
        column: 0,
        writeValue: false,
        cell: { value: null },
        style: delta!,
      },
    ])
    const zip = await JSZip.loadAsync(saved.buffer)
    const styles = await zip.file('xl/styles.xml')!.async('string')
    const sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
    expect(styles).toContain('diagonalUp="1" diagonalDown="1"')
    expect(styles).toContain('<diagonal style="dotted"><color auto="1"/></diagonal>')
    for (const edge of ['left', 'right', 'top', 'bottom'])
      expect(styles).toContain(`<${edge} style="dotted"><color auto="1"/></${edge}>`)
    for (const [address, value] of [
      ['B1', '2'],
      ['A2', '3'],
      ['B2', '4'],
    ])
      expect(sheet).toContain(`<c r="${address}" s="1"><v>${value}</v></c>`)
    expect(sheet).toMatch(/<c r="A1" s="\d+"><v>1<\/v><\/c>/)

    const restored = await applyCellEditsToXlsx(saved.buffer, [
      {
        sheetName: 'Sheet1',
        row: 0,
        column: 0,
        writeValue: false,
        cell: { value: null },
        styleReference: 1,
      },
    ])
    const restoredSheet = await (
      await JSZip.loadAsync(restored.buffer)
    )
      .file('xl/worksheets/sheet1.xml')!
      .async('string')
    expect(restoredSheet).toContain('<c r="A1" s="1"><v>1</v></c>')
  })

  it('edits only chosen edge and diagonal flags while retaining unknown border/XF XML', () => {
    const editor = new StylesheetEditor(BASE)
    editor.resolveStyle(0, {
      borderTop: { style: 'mediumDashDotDot', color: '#ABCDEF' },
      borderDiagonal: { style: 'dashDot', color: '#556677' },
      diagonalUp: true,
    })
    const xml = editor.serialize()
    expect(xml).toContain('<border custom="keep>a" diagonalDown="1" diagonalUp="1">')
    expect(xml).toContain('<left style="medium"><color rgb="FF102030"/></left>')
    expect(xml).toContain('<top style="mediumDashDotDot"><color rgb="FFABCDEF"/></top>')
    expect(xml).toContain('<diagonal style="dashDot"><color rgb="FF556677"/></diagonal>')
    expect(xml).toContain('<horizontal style="dotted"/><vertical style="dashed"/>')
    expect(xml).toContain('<extLst><ext uri="vendor"/></extLst>')
    expect(xml).toContain('custom="xf" applyBorder="1"><extLst><ext uri="xf-vendor"/></extLst>')
    expect(xml).toContain('<font/>')
  })

  it('keeps imported XF extensions when a font and border are changed together', () => {
    const editor = new StylesheetEditor(BASE)
    editor.resolveStyle(0, {
      bold: true,
      borderRight: { style: 'double', color: '#334455' },
    })
    const xml = editor.serialize()
    expect(xml).toContain('custom="xf"')
    expect(xml).toContain('<extLst><ext uri="xf-vendor"/></extLst>')
    expect(xml).toContain('<b/>')
    expect(xml).toContain('<right style="double"><color rgb="FF334455"/></right>')
  })

  it('updates a direct alignment child without disturbing its other fields or XF extension', () => {
    const base = BASE.replace(
      '<xf fontId="0" borderId="0" custom="xf"><extLst>',
      '<xf fontId="0" borderId="0" custom="xf" quotePrefix="1"><alignment horizontal="left" readingOrder="2" x14:feature="keep"/><extLst>',
    )
    const editor = new StylesheetEditor(base)
    const index = editor.resolveStyle(0, {
      borderRight: { style: 'dotted', color: '#445566' },
      verticalAlignment: 'distributed',
    })
    const xml = editor.serialize()
    const xfs = [
      ...(/<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/.exec(xml)?.[1] ?? '').matchAll(
        /<xf\b[^>]*\/>|<xf\b[^>]*>[\s\S]*?<\/xf>/g,
      ),
    ].map((match) => match[0])
    expect(xfs[index]).toContain('quotePrefix="1"')
    expect(xfs[index]).toContain(
      '<alignment horizontal="left" readingOrder="2" x14:feature="keep" vertical="distributed"/>',
    )
    expect(xfs[index]).toContain('<extLst><ext uri="xf-vendor"/></extLst>')
  })

  it('retains imported XF metadata for Font + Number and Fill + Protection without border edits', () => {
    const editor = new StylesheetEditor(BASE)
    const fontNumber = editor.resolveStyle(0, { bold: true, numberFormat: '0.00' })
    const fillProtection = editor.resolveStyle(0, {
      fillColor: '#AABBCC',
      protectionLocked: false,
      protectionHidden: true,
    })
    const xml = editor.serialize()
    const xfs = [
      ...(/<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/.exec(xml)?.[1] ?? '').matchAll(
        /<xf\b[^>]*\/>|<xf\b[^>]*>[\s\S]*?<\/xf>/g,
      ),
    ].map((match) => match[0])
    for (const index of [fontNumber, fillProtection]) {
      expect(xfs[index]).toContain('custom="xf"')
      expect(xfs[index]).toContain('<extLst><ext uri="xf-vendor"/></extLst>')
    }
    expect(xfs[fontNumber]).toContain('applyFont="1"')
    expect(xfs[fontNumber]).toContain('numFmtId="2"')
    expect(xfs[fillProtection]).toContain('applyFill="1"')
    expect(xfs[fillProtection]).toContain('applyProtection="1"')
    expect(xfs[fillProtection]).toContain('locked="0"')
    expect(xfs[fillProtection]).toContain('hidden="1"')
  })

  it('turns one edge and diagonal direction off without removing the other direction or extensions', () => {
    const editor = new StylesheetEditor(BASE)
    editor.resolveStyle(0, { borderLeft: null, diagonalUp: false })
    const xml = editor.serialize()
    expect(xml).toContain('diagonalDown="1" diagonalUp="0"')
    expect(xml).toContain('<left/>')
    expect(xml).toContain('<diagonal style="hair"/>')
    expect(xml).toContain('<horizontal style="dotted"/><vertical style="dashed"/>')
    expect(xml).toContain('<extLst><ext uri="vendor"/></extLst>')
  })

  it('maps a 2×2 inside-border action to both sides of each shared seam', () => {
    const range: ProtectionRange = {
      startRow: 0,
      endRow: 1,
      startColumn: 0,
      endColumn: 1,
    }
    const base = {
      locked: true,
      hidden: false,
      horizontalAlignment: 'general',
      verticalAlignment: 'bottom',
      textRotation: 0,
      shrinkToFit: false,
    }
    const cells: CellProtectionFlags[] = [
      { ...base, row: 0, column: 0 },
      { ...base, row: 0, column: 1 },
      { ...base, row: 1, column: 0 },
      { ...base, row: 1, column: 1 },
    ]
    const request = {
      modes: { insideHorizontal: true, insideVertical: true, diagonalUp: true },
      style: 'double' as const,
      color: '#112233',
    }
    expect(borderActionOfCommand(borderActionCommand(request))).toEqual(request)
    expect(cells.map((cell) => borderDeltaForCell(request, range, cell))).toEqual([
      {
        borderBottom: { style: 'double', color: '#112233' },
        borderRight: { style: 'double', color: '#112233' },
        diagonalUp: true,
        borderDiagonal: { style: 'double', color: '#112233' },
      },
      {
        borderBottom: { style: 'double', color: '#112233' },
        borderLeft: { style: 'double', color: '#112233' },
        diagonalUp: true,
        borderDiagonal: { style: 'double', color: '#112233' },
      },
      {
        borderTop: { style: 'double', color: '#112233' },
        borderRight: { style: 'double', color: '#112233' },
        diagonalUp: true,
        borderDiagonal: { style: 'double', color: '#112233' },
      },
      {
        borderTop: { style: 'double', color: '#112233' },
        borderLeft: { style: 'double', color: '#112233' },
        diagonalUp: true,
        borderDiagonal: { style: 'double', color: '#112233' },
      },
    ])
    expect(
      selectionBorderState(
        cells.map((cell, index) => ({
          ...cell,
          borderRight: index % 2 === 0,
          borderLeft: index % 2 === 1,
        })),
        range,
      ),
    ).toMatchObject({ insideVertical: true, insideHorizontal: false })
  })

  it('saves native 2×2 inside and diagonal borders without touching cell values or formula cache', async () => {
    const range: ProtectionRange = {
      startRow: 0,
      endRow: 1,
      startColumn: 0,
      endColumn: 1,
    }
    const action = {
      modes: { insideHorizontal: true, insideVertical: true, diagonalDown: true },
      style: 'dashDot' as const,
      color: '#234567',
    }
    const edits = [
      { row: 0, column: 0 },
      { row: 0, column: 1 },
      { row: 1, column: 0 },
      { row: 1, column: 1 },
    ].map((cell) => ({
      sheetName: 'Data',
      ...cell,
      writeValue: false,
      cell: { value: null },
      style: borderDeltaForCell(action, range, {
        ...cell,
        locked: true,
        hidden: false,
        horizontalAlignment: 'general',
        verticalAlignment: 'bottom',
        textRotation: 0,
        shrinkToFit: false,
      })!,
    }))
    const saved = await applyCellEditsToXlsx(await buildEditFixture(), edits)
    const zip = await JSZip.loadAsync(saved.buffer)
    const sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
    const styles = await zip.file('xl/styles.xml')!.async('string')
    const xfs = [
      ...(/<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/.exec(styles)?.[1] ?? '').matchAll(
        /<xf\b[^>]*\/>|<xf\b[^>]*>[\s\S]*?<\/xf>/g,
      ),
    ].map((match) => match[0])
    const borders = [
      ...(/<borders\b[^>]*>([\s\S]*?)<\/borders>/.exec(styles)?.[1] ?? '').matchAll(
        /<border\b[^>]*\/>|<border\b[^>]*>[\s\S]*?<\/border>/g,
      ),
    ].map((match) => match[0])
    const borderOf = (address: string): string => {
      const cell = new RegExp(`<c\\b[^>]*r="${address}"[^>]*`).exec(sheet)?.[0] ?? ''
      const xf = xfs[Number(/\bs="(\d+)"/.exec(cell)?.[1] ?? 0)] ?? ''
      return borders[Number(/\bborderId="(\d+)"/.exec(xf)?.[1] ?? 0)] ?? ''
    }
    expect(borderOf('A1')).toContain('<right style="dashDot">')
    expect(borderOf('A1')).toContain('<bottom style="dashDot">')
    expect(borderOf('B1')).toContain('<left style="dashDot">')
    expect(borderOf('A2')).toContain('<top style="dashDot">')
    expect(borderOf('B2')).toContain('<diagonal style="dashDot">')
    for (const address of ['A1', 'B1', 'A2', 'B2'])
      expect(borderOf(address)).toContain('diagonalDown="1"')
    expect(sheet).toContain('<c r="A1" t="s" s="')
    expect(sheet).toContain('<v>0</v>')
    expect(sheet).toContain('<f>SUM(C1:C2)</f><v>5</v>')
  })

  it('backfills a mixed imported edge and restores the original XF after Save → Undo', async () => {
    const base = {
      bold: false,
      italic: false,
      underline: false,
      strikethrough: false,
      wrapText: false,
      diagonalUp: false,
      diagonalDown: false,
    }
    const selection = selectionProtection(
      [base, { ...base, borderTop: { style: 'hair', color: '#102030' } }],
      { columnWidths: [] } as unknown as WorkbookFile['sheets'][number],
      {
        cells: [
          { row: 0, column: 0, value: 'A', styleIndex: 1 },
          { row: 0, column: 1, value: 'B', styleIndex: 0 },
        ],
        rows: [],
      },
      createEditJournal(),
      'sheet-1',
      { startRow: 0, endRow: 0, startColumn: 0, endColumn: 1 },
    )
    expect(selection.border).toMatchObject({ top: null, insideHorizontal: null })
    const single = selectionProtection(
      [base, { ...base, borderTop: { style: 'hair', color: '#102030' } }],
      { columnWidths: [] } as unknown as WorkbookFile['sheets'][number],
      { cells: [{ row: 0, column: 0, value: 'A', styleIndex: 1 }], rows: [] },
      createEditJournal(),
      'sheet-1',
      { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 },
    )
    expect(draftFromSelection(null, single)).toMatchObject({
      borderStyle: 'hair',
      borderColor: '#102030',
      borderEdges: { top: 'on', bottom: 'off' },
    })

    const original = await buildEditFixture()
    const edited = await applyCellEditsToXlsx(original, [
      {
        sheetName: 'Data',
        row: 0,
        column: 0,
        writeValue: false,
        cell: { value: null },
        style: { borderTop: { style: 'double', color: '#112233' } },
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
      toSaveEdits(journal).map((entry) => ({
        sheetName: 'Data',
        row: entry.row,
        column: entry.column,
        writeValue: false,
        cell: { value: null },
        styleReference: entry.styleReference,
      })),
    )
    const sheet = await (
      await JSZip.loadAsync(undone.buffer)
    )
      .file('xl/worksheets/sheet1.xml')!
      .async('string')
    expect(sheet).toContain('<c r="A1" t="s" s="1"><v>0</v></c>')
    expect(sheet).toContain('<f>SUM(C1:C2)</f><v>5</v>')
  })
})
