import JSZip from 'jszip'
import { readFile } from 'node:fs/promises'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'
import { XMLValidator } from 'fast-xml-parser'

import { applyCellEditsToXlsx } from '@genoffice/xlsx-gateway/gateway/xlsx-gateway'
import { borderDeltaForCell } from '../src/renderer/border-format'
import { createEditJournal } from '../src/renderer/edit-journal'
import { draftFromSelection } from '../src/renderer/format-cells'
import {
  borderControlHasVisibleSeam,
  borderSelectionIsSafe,
  canEditSelectionBorder,
} from '../src/renderer/merged-border-geometry'
import { selectionBorderState, selectionProtection } from '../src/renderer/selection-protection'
import type { CellProtectionFlags, ProtectionRange } from '../src/renderer/selection-protection'
import type { WorkbookCellStyle, WorkbookFile } from '../src/shared/desktop-api'

const REPO_ROOT = join(fileURLToPath(new URL('.', import.meta.url)), '../../..')

const MERGE: ProtectionRange = {
  startRow: 0,
  endRow: 1,
  startColumn: 0,
  endColumn: 1,
}
const ADJACENT: ProtectionRange = { ...MERGE, endColumn: 2 }
const flags = (row: number, column: number): CellProtectionFlags => ({
  row,
  column,
  locked: true,
  hidden: false,
  horizontalAlignment: 'general',
  verticalAlignment: 'bottom',
  textRotation: 0,
  shrinkToFit: false,
})
const cells = [flags(0, 0), flags(0, 1), flags(1, 0), flags(1, 1)]

function xmlStyles(styles: string): { xfs: string[]; borders: string[] } {
  const section = (name: string) =>
    new RegExp(`<${name}\\b[^>]*>([\\s\\S]*?)<\\/${name}>`).exec(styles)?.[1] ?? ''
  return {
    xfs: [...section('cellXfs').matchAll(/<xf\b[^>]*\/>|<xf\b[^>]*>[\s\S]*?<\/xf>/g)].map(
      (match) => match[0],
    ),
    borders: [
      ...section('borders').matchAll(/<border\b[^>]*\/>|<border\b[^>]*>[\s\S]*?<\/border>/g),
    ].map((match) => match[0]),
  }
}

function styleReference(sheet: string, address: string): number | null {
  const cell = new RegExp(`<c\\b[^>]*\\br="${address}"[^>]*>`).exec(sheet)?.[0]
  return cell ? Number(/\bs="(\d+)"/.exec(cell)?.[1] ?? 0) : null
}

function borderOf(sheet: string, styles: string, address: string): string {
  const style = xmlStyles(styles)
  const reference = styleReference(sheet, address)
  const xf = style.xfs[reference ?? 0] ?? ''
  return style.borders[Number(/\bborderId="(\d+)"/.exec(xf)?.[1] ?? 0)] ?? ''
}

describe('merged-cell borders', () => {
  it('backfills an imported Excel corner-style merge as one outline and hides the internal seams', () => {
    const base = {
      bold: false,
      italic: false,
      underline: false,
      strikethrough: false,
      wrapText: false,
      diagonalUp: false,
      diagonalDown: false,
    }
    const edge = { style: 'dotted' as const, color: '#000000' }
    const styles: WorkbookCellStyle[] = [
      base,
      { ...base, borderLeft: edge, borderTop: edge, borderDiagonal: edge, diagonalUp: true },
      { ...base, borderRight: edge, borderTop: edge, borderDiagonal: edge, diagonalUp: true },
      { ...base, borderLeft: edge, borderBottom: edge, borderDiagonal: edge, diagonalUp: true },
      { ...base, borderRight: edge, borderBottom: edge, borderDiagonal: edge, diagonalUp: true },
    ]
    const selected = selectionProtection(
      styles,
      { columnWidths: [] } as unknown as WorkbookFile['sheets'][number],
      {
        cells: cells.map((cell, index) => ({
          row: cell.row,
          column: cell.column,
          value: index === 0 ? 'Merged sample' : null,
          styleIndex: index + 1,
        })),
        rows: [],
        merges: [MERGE],
      },
      createEditJournal(),
      'sheet-1',
      MERGE,
    )
    expect(selected.merges).toEqual([MERGE])
    expect(selected.border).toMatchObject({
      left: true,
      right: true,
      top: true,
      bottom: true,
      insideHorizontal: null,
      insideVertical: null,
      diagonalUp: true,
      diagonalDown: false,
      lineStyle: 'dotted',
    })
    expect(draftFromSelection(null, selected).borderEdges).toMatchObject({
      left: 'on',
      right: 'on',
      top: 'on',
      bottom: 'on',
      insideHorizontal: '',
      insideVertical: '',
      diagonalUp: 'on',
    })
  })

  it('uses visible regions: no internal merge seam, but the adjacent cell is an inside seam', () => {
    expect(borderSelectionIsSafe(MERGE, [MERGE])).toBe(true)
    expect(borderSelectionIsSafe(ADJACENT, [MERGE])).toBe(true)
    expect(borderSelectionIsSafe({ ...MERGE, endColumn: 0 }, [MERGE])).toBe(false)
    expect(canEditSelectionBorder({ range: MERGE, merges: [MERGE], mergesComplete: true })).toBe(
      true,
    )
    expect(canEditSelectionBorder({ range: MERGE, merges: [], mergesComplete: false })).toBe(false)
    expect(borderControlHasVisibleSeam(MERGE, [MERGE], 'down')).toBe(false)
    expect(borderControlHasVisibleSeam(MERGE, [MERGE], 'right')).toBe(false)
    expect(borderControlHasVisibleSeam(ADJACENT, [MERGE], 'down')).toBe(true)
    expect(borderControlHasVisibleSeam(ADJACENT, [MERGE], 'right')).toBe(true)
    const all = {
      modes: {
        left: true,
        right: true,
        top: true,
        bottom: true,
        insideHorizontal: true,
        insideVertical: true,
        diagonalUp: true,
      },
      style: 'dotted' as const,
      color: '#102030',
    }
    expect(cells.map((cell) => borderDeltaForCell(all, MERGE, cell, [MERGE]))).toEqual([
      {
        borderLeft: { style: 'dotted', color: '#102030' },
        borderTop: { style: 'dotted', color: '#102030' },
        diagonalUp: true,
        borderDiagonal: { style: 'dotted', color: '#102030' },
      },
      {
        borderTop: { style: 'dotted', color: '#102030' },
        borderRight: { style: 'dotted', color: '#102030' },
        diagonalUp: true,
        borderDiagonal: { style: 'dotted', color: '#102030' },
      },
      {
        borderLeft: { style: 'dotted', color: '#102030' },
        borderBottom: { style: 'dotted', color: '#102030' },
        diagonalUp: true,
        borderDiagonal: { style: 'dotted', color: '#102030' },
      },
      {
        borderBottom: { style: 'dotted', color: '#102030' },
        borderRight: { style: 'dotted', color: '#102030' },
        diagonalUp: true,
        borderDiagonal: { style: 'dotted', color: '#102030' },
      },
    ])
    expect(selectionBorderState(cells, MERGE, [MERGE])).toMatchObject({
      insideHorizontal: null,
      insideVertical: null,
    })
    const seam = [...cells, flags(0, 2), flags(1, 2)].map((cell) =>
      borderDeltaForCell(all, ADJACENT, cell, [MERGE]),
    )
    expect(seam[1]).toMatchObject({ borderRight: { style: 'dotted', color: '#102030' } })
    expect(seam[3]).toMatchObject({ borderRight: { style: 'dotted', color: '#102030' } })
    expect(seam[4]).toMatchObject({ borderLeft: { style: 'dotted', color: '#102030' } })
    expect(seam[5]).toMatchObject({ borderLeft: { style: 'dotted', color: '#102030' } })
    expect(seam[0]).not.toHaveProperty('borderRight')
    expect(seam[2]).not.toHaveProperty('borderRight')
    const cleared = [...cells, flags(0, 2), flags(1, 2)].map((cell) =>
      borderDeltaForCell(
        { modes: { insideVertical: false }, style: 'dotted', color: '#102030' },
        ADJACENT,
        cell,
        [MERGE],
      ),
    )
    expect(cleared[0]).toBeNull()
    expect(cleared[1]).toEqual({ borderRight: null })
    expect(cleared[2]).toBeNull()
    expect(cleared[3]).toEqual({ borderRight: null })
    expect(cleared[4]).toEqual({ borderLeft: null })
    expect(cleared[5]).toEqual({ borderLeft: null })
  })

  it('edits the real Excel Mac merged fixture without changing the neighbor or merged text, and unmerges with the corner styles intact', async () => {
    let source: Buffer = await readFile(
      join(
        REPO_ROOT,
        'docs/qa/office-mac-field-evidence-2026-10-13/sheets-merged-border/excel-mac-merged-border-source.xlsx',
      ),
    )
    const input = await JSZip.loadAsync(source)
    const originalSheet = await input.file('xl/worksheets/sheet1.xml')!.async('string')
    const originalStyles = await input.file('xl/styles.xml')!.async('string')
    expect(originalSheet).toContain('<mergeCell ref="A1:B2"/>')
    expect(
      ['A1', 'B1', 'A2', 'B2'].map((address) => styleReference(originalSheet, address)),
    ).toEqual([1, 2, 3, 4])
    const macUnmerged = await JSZip.loadAsync(
      await readFile(
        join(
          REPO_ROOT,
          'docs/qa/office-mac-field-evidence-2026-10-13/sheets-merged-border/excel-mac-unmerged-after-border.xlsx',
        ),
      ),
    )
    const macUnmergedSheet = await macUnmerged.file('xl/worksheets/sheet1.xml')!.async('string')
    expect(macUnmergedSheet).not.toContain('<mergeCell ref="A1:B2"/>')
    expect(
      ['A1', 'B1', 'A2', 'B2'].map((address) => styleReference(macUnmergedSheet, address)),
    ).toEqual([1, 2, 3, 4])
    expect(styleReference(originalSheet, 'C1')).toBe(0)
    for (const address of ['A1', 'B1', 'A2', 'B2']) {
      expect(borderOf(originalSheet, originalStyles, address)).toContain('diagonalUp="1"')
      expect(borderOf(originalSheet, originalStyles, address)).toContain('<color auto="1"/>')
    }
    const firstXf = xmlStyles(originalStyles).xfs[1]!
    const extendedXf = firstXf
      .replace(/^<xf\b/, '<xf quotePrefix="1"')
      .replace('</xf>', '<extLst><ext uri="urn:vendor:keep"/></extLst></xf>')
    const extendedStyles = originalStyles.replace(firstXf, extendedXf)
    expect(XMLValidator.validate(extendedStyles)).toBe(true)
    input.file('xl/styles.xml', extendedStyles)
    input.file(
      'xl/worksheets/sheet1.xml',
      originalSheet.replace(
        '<c r="B2" s="4"/></row>',
        '<c r="B2" s="4"/><c r="C2"><f>1+2</f><v>3</v></c></row>',
      ),
    )
    source = await input.generateAsync({ type: 'nodebuffer' })
    const action = {
      modes: { left: true, right: true, top: true, bottom: true, diagonalDown: true },
      style: 'double' as const,
      color: '#445566',
    }
    const edits = cells.map((cell) => ({
      sheetName: 'Sheet1',
      row: cell.row,
      column: cell.column,
      writeValue: false,
      cell: { value: null },
      style: borderDeltaForCell(action, MERGE, { ...cell, diagonalUp: true }, [MERGE])!,
    }))
    const changed = await applyCellEditsToXlsx(source, edits)
    const changedZip = await JSZip.loadAsync(changed.buffer)
    const changedSheet = await changedZip.file('xl/worksheets/sheet1.xml')!.async('string')
    const changedStyles = await changedZip.file('xl/styles.xml')!.async('string')
    expect(changedSheet).toContain('<mergeCell ref="A1:B2"/>')
    expect(changedSheet).toContain('<c r="C1" t="s"><v>1</v></c>')
    expect(changedSheet).toContain('<c r="C2"><f>1+2</f><v>3</v></c>')
    expect(changedSheet).toContain('<c r="A1"')
    expect(changedSheet).toContain('<v>0</v>')
    const editedA1Xf = xmlStyles(changedStyles).xfs[styleReference(changedSheet, 'A1')!]!
    expect(editedA1Xf).toContain('quotePrefix="1"')
    expect(editedA1Xf).toContain('<extLst><ext uri="urn:vendor:keep"/></extLst>')
    for (const [address, edges] of [
      ['A1', ['left', 'top']],
      ['B1', ['right', 'top']],
      ['A2', ['left', 'bottom']],
      ['B2', ['right', 'bottom']],
    ] as const) {
      const border = borderOf(changedSheet, changedStyles, address)
      for (const edge of edges) expect(border).toContain(`<${edge} style="double">`)
      expect(border).toContain('diagonalUp="1" diagonalDown="1"')
      expect(border).toContain('<diagonal style="dotted"><color auto="1"/></diagonal>')
      if (address === 'A1' || address === 'A2') expect(border).toContain('<right/>')
      if (address === 'A1' || address === 'B1') expect(border).toContain('<bottom/>')
    }
    const restored = await applyCellEditsToXlsx(
      changed.buffer,
      cells.map((cell, index) => ({
        sheetName: 'Sheet1',
        row: cell.row,
        column: cell.column,
        writeValue: false,
        cell: { value: null },
        styleReference: index + 1,
      })),
    )
    const restoredZip = await JSZip.loadAsync(restored.buffer)
    const restoredSheet = await restoredZip.file('xl/worksheets/sheet1.xml')!.async('string')
    expect(
      ['A1', 'B1', 'A2', 'B2'].map((address) => styleReference(restoredSheet, address)),
    ).toEqual([1, 2, 3, 4])
    expect(restoredSheet).toContain('<c r="C2"><f>1+2</f><v>3</v></c>')
    const unmerged = await applyCellEditsToXlsx(
      changed.buffer,
      [],
      [{ sheetName: 'Sheet1', ops: [{ kind: 'unmerge-cells', range: MERGE }] }],
    )
    const unmergedZip = await JSZip.loadAsync(unmerged.buffer)
    const unmergedSheet = await unmergedZip.file('xl/worksheets/sheet1.xml')!.async('string')
    expect(unmergedSheet).not.toContain('<mergeCell ref="A1:B2"/>')
    expect(unmergedSheet).toContain('<c r="C2"><f>1+2</f><v>3</v></c>')
    expect(
      ['A1', 'B1', 'A2', 'B2'].map((address) => styleReference(unmergedSheet, address)),
    ).toEqual(['A1', 'B1', 'A2', 'B2'].map((address) => styleReference(changedSheet, address)))
    const unmergedStyles = await unmergedZip.file('xl/styles.xml')!.async('string')
    for (const address of ['A1', 'B1', 'A2', 'B2'])
      expect(borderOf(unmergedSheet, unmergedStyles, address)).toBe(
        borderOf(changedSheet, changedStyles, address),
      )
  })

  it('writes only visible B/C and C1/C2 inside seams beside an imported merge', async () => {
    const source = await readFile(
      join(
        REPO_ROOT,
        'docs/qa/office-mac-field-evidence-2026-10-13/sheets-merged-border/excel-mac-merged-border-source.xlsx',
      ),
    )
    const action = {
      modes: { insideHorizontal: true, insideVertical: false },
      style: 'dotted' as const,
      color: '#000000',
    }
    const edits = [...cells, flags(0, 2), flags(1, 2)].flatMap((cell) => {
      const style = borderDeltaForCell(action, ADJACENT, cell, [MERGE])
      return style
        ? [
            {
              sheetName: 'Sheet1',
              row: cell.row,
              column: cell.column,
              writeValue: false,
              cell: { value: null },
              style,
            },
          ]
        : []
    })
    const changed = await applyCellEditsToXlsx(source, edits)
    const zip = await JSZip.loadAsync(changed.buffer)
    const sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
    const styles = await zip.file('xl/styles.xml')!.async('string')
    expect(sheet).toContain('<mergeCell ref="A1:B2"/>')
    expect(sheet).toContain('<c r="C1"')
    expect(sheet).toContain('<v>1</v>')
    expect(borderOf(sheet, styles, 'A1')).not.toContain('<right style=')
    expect(borderOf(sheet, styles, 'A1')).not.toContain('<bottom style=')
    expect(borderOf(sheet, styles, 'B1')).not.toContain('<right style=')
    expect(borderOf(sheet, styles, 'B2')).not.toContain('<right style=')
    expect(borderOf(sheet, styles, 'C1')).not.toContain('<left style=')
    expect(borderOf(sheet, styles, 'C2')).not.toContain('<left style=')
    expect(borderOf(sheet, styles, 'C1')).toContain('<bottom style="dotted">')
    expect(borderOf(sheet, styles, 'C2')).toContain('<top style="dotted">')
  })
})
