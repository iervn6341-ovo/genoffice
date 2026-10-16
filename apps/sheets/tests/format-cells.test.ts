import { describe, expect, it } from 'vitest'

import {
  accountingUnderlineChangeBlocked,
  accountingUnderlineUnavailable,
  degreesToOoxmlRotation,
  draftFromSelection,
  formatCellsCommands,
  ooxmlRotationDegrees,
} from '../src/renderer/format-cells'
import { createEditJournal, recordNeutralStyleEdit } from '../src/renderer/edit-journal'
import { protectionReadRange, selectionProtection } from '../src/renderer/selection-protection'

import type { WorkbookCellStyle, WorkbookFile } from '../src/shared/desktop-api'
import type { SelectionFormat } from '../src/renderer/selection-format'

const SELECTION: SelectionFormat = {
  fontFamily: 'Calibri',
  fontSize: 14,
  bold: true,
  italic: false,
  underline: false,
  strike: false,
  wrap: false,
  horizontalAlignment: 'left',
  verticalAlignment: 'center',
  textRotation: 45,
  indent: 0,
  fontColor: '#C00000',
  fillColor: '#FFF2CC',
  numberFormat: '0.00%',
  link: null,
}

describe('draftFromSelection', () => {
  it('prefills every echoed setting', () => {
    expect(draftFromSelection(SELECTION)).toEqual({
      pattern: '0.00%',
      family: 'Calibri',
      size: '14',
      bold: 'on',
      italic: 'off',
      underline: 'off',
      underlineMode: 'none',
      fontVerticalAlign: '',
      strike: 'off',
      fontColor: '#c00000',
      border: '',
      borderColor: '#000000',
      borderStyle: 'thin',
      borderEdges: {
        left: '',
        top: '',
        bottom: '',
        right: '',
        insideHorizontal: '',
        insideVertical: '',
        diagonalUp: '',
        diagonalDown: '',
      },
      fill: '#fff2cc',
      fillPattern: '',
      patternColor: '#000000',
      noFill: false,
      hAlign: 'left',
      vAlign: 'middle',
      wrapText: 'off',
      shrinkToFit: '',
      mergeCells: '',
      rotation: '45',
      indent: '',
      locked: '',
      hidden: '',
    })
  })

  it('falls back to unchanged tri-states without a selection', () => {
    const draft = draftFromSelection(null)
    expect(draft.pattern).toBe('')
    expect(draft.bold).toBe('')
    expect(draft.fontColor).toBe('')
    expect(draft.fill).toBe('')
  })
})

describe('formatCellsCommands', () => {
  const initial = draftFromSelection(SELECTION)

  it('emits nothing when the draft is untouched', () => {
    expect(formatCellsCommands(initial, initial)).toEqual([])
  })

  it('emits only the changed settings', () => {
    const edited = { ...initial, pattern: '$#,##0.00', bold: 'off' as const, fill: '#00aa55' }
    expect(formatCellsCommands(initial, edited)).toEqual([
      'format:$#,##0.00',
      'bold:off',
      'fill:#00aa55',
    ])
  })

  it('re-picking the current values emits nothing', () => {
    const edited = { ...initial, family: 'Calibri', fontColor: '#c00000' }
    expect(formatCellsCommands(initial, edited)).toEqual([])
  })

  it('any border preset counts as a change and carries its color and line style', () => {
    const edited = { ...initial, border: 'all', borderColor: '#112233' }
    expect(formatCellsCommands(initial, edited)).toEqual(['border-edges:111111--:thin:#112233'])
    const doubled = { ...initial, border: 'bottom', borderStyle: 'double' }
    expect(formatCellsCommands(initial, doubled)).toEqual(['border-edges:--1-----:double:#000000'])
  })

  it('no-fill clears only when the selection had a fill', () => {
    expect(formatCellsCommands(initial, { ...initial, noFill: true })).toEqual(['fill:none'])
    const unfilled = draftFromSelection({ ...SELECTION, fillColor: null })
    expect(formatCellsCommands(unfilled, { ...unfilled, noFill: true })).toEqual([])
  })

  it('explicit white on an unfilled selection emits a fill command', () => {
    const unfilled = draftFromSelection({ ...SELECTION, fillColor: null })
    expect(unfilled.fill).toBe('')
    expect(formatCellsCommands(unfilled, { ...unfilled, fill: '#ffffff' })).toEqual([
      'fill:#ffffff',
    ])
  })

  it('emits alignment, wrap, rotation, and indent changes', () => {
    const edited = {
      ...initial,
      hAlign: 'distributed',
      vAlign: 'top',
      wrapText: 'on' as const,
      rotation: 'vertical',
      indent: '2',
    }
    expect(formatCellsCommands(initial, edited)).toEqual([
      'align:distributed',
      'valign:top',
      'wrap:on',
      'rotate-ooxml:255',
      'indent:2',
    ])
  })

  it('re-picking the echoed alignment emits nothing', () => {
    const edited = { ...initial, hAlign: 'left', vAlign: 'middle', wrapText: 'off' as const }
    expect(formatCellsCommands(initial, edited)).toEqual([])
  })

  it('echoes an existing indent and clears it with an explicit 0', () => {
    const indented = draftFromSelection({ ...SELECTION, indent: 3 })
    expect(indented.indent).toBe('3')
    expect(formatCellsCommands(indented, indented)).toEqual([])
    expect(formatCellsCommands(indented, { ...indented, indent: '0' })).toEqual(['indent:0'])
  })

  it('combines protection flags into one command', () => {
    expect(formatCellsCommands(initial, { ...initial, locked: 'off' as const })).toEqual([
      'cellprot:locked-off:',
    ])
    expect(
      formatCellsCommands(initial, { ...initial, locked: 'off' as const, hidden: 'on' as const }),
    ).toEqual(['cellprot:locked-off:hidden-on'])
  })

  it('keeps a mixed native alignment blank and changes only the chosen attribute', () => {
    const journal = createEditJournal()
    const styles: WorkbookCellStyle[] = [
      {
        bold: false,
        italic: false,
        underline: false,
        strikethrough: false,
        wrapText: false,
        diagonalUp: false,
        diagonalDown: false,
      },
      {
        bold: false,
        italic: false,
        underline: false,
        strikethrough: false,
        wrapText: false,
        diagonalUp: false,
        diagonalDown: false,
        horizontalAlignment: 'fill',
        verticalAlignment: 'distributed',
        textRotation: 37,
        shrinkToFit: true,
      },
      {
        bold: false,
        italic: false,
        underline: false,
        strikethrough: false,
        wrapText: false,
        diagonalUp: false,
        diagonalDown: false,
        horizontalAlignment: 'centerContinuous',
        verticalAlignment: 'justify',
        textRotation: 135,
      },
    ]
    const selection = selectionProtection(
      styles,
      { columnWidths: [] } as unknown as WorkbookFile['sheets'][number],
      {
        cells: [
          { row: 0, column: 0, value: 'one', styleIndex: 1 },
          { row: 0, column: 1, value: 'two', styleIndex: 2 },
        ],
        rows: [],
      },
      journal,
      'sheet-1',
      { startRow: 0, endRow: 0, startColumn: 0, endColumn: 1 },
    )
    expect(selection.alignment).toEqual({
      horizontal: null,
      vertical: null,
      rotation: null,
      shrinkToFit: null,
      merged: false,
    })
    const draft = draftFromSelection(SELECTION, selection)
    expect(draft).toMatchObject({ hAlign: '', vAlign: '', rotation: '', shrinkToFit: '' })
    expect(formatCellsCommands(draft, { ...draft, hAlign: 'centerContinuous' })).toEqual([
      'align:centerContinuous',
    ])
    expect(formatCellsCommands(draft, { ...draft, shrinkToFit: 'on' })).toEqual(['shrink:on'])
  })

  it('converts arbitrary signed degrees to reversible OOXML rotation values', () => {
    expect(degreesToOoxmlRotation('37')).toBe('37')
    expect(degreesToOoxmlRotation('-37')).toBe('127')
    expect(ooxmlRotationDegrees('127')).toBe('-37')
    expect(degreesToOoxmlRotation('91')).toBe('')
    expect(degreesToOoxmlRotation('1.5')).toBe('')
  })

  it('keeps Center Across Selection separate from a true merged range', () => {
    const full = { startRow: 0, endRow: 0, startColumn: 0, endColumn: 1 }
    const selected = selectionProtection(
      [
        {
          bold: false,
          italic: false,
          underline: false,
          strikethrough: false,
          wrapText: false,
          diagonalUp: false,
          diagonalDown: false,
          horizontalAlignment: 'centerContinuous',
        },
      ],
      { columnWidths: [] } as unknown as WorkbookFile['sheets'][number],
      { cells: [], rows: [], merges: [full] },
      createEditJournal(),
      'sheet-1',
      full,
    )
    expect(selected.alignment.horizontal).toBe('centerContinuous')
    expect(selected.alignment.merged).toBe(true)
    const draft = draftFromSelection(SELECTION, selected)
    expect(formatCellsCommands(draft, { ...draft, mergeCells: 'off' })).toEqual(['merge:unmerge'])
    expect(formatCellsCommands(draft, { ...draft, hAlign: 'centerContinuous' })).toEqual([])
  })

  it('reports a partly merged selection as mixed instead of fully merged', () => {
    const selected = selectionProtection(
      [
        {
          bold: false,
          italic: false,
          underline: false,
          strikethrough: false,
          wrapText: false,
          diagonalUp: false,
          diagonalDown: false,
        },
      ],
      { columnWidths: [] } as unknown as WorkbookFile['sheets'][number],
      {
        cells: [],
        rows: [],
        merges: [{ startRow: 0, endRow: 0, startColumn: 0, endColumn: 1 }],
      },
      createEditJournal(),
      'sheet-1',
      { startRow: 0, endRow: 0, startColumn: 0, endColumn: 2 },
    )
    expect(selected.alignment.merged).toBeNull()
    expect(draftFromSelection(SELECTION, selected).mergeCells).toBe('')
  })
})

describe('selection protection backfill', () => {
  const style = (flags: Partial<WorkbookCellStyle> = {}): WorkbookCellStyle => ({
    bold: false,
    italic: false,
    underline: false,
    strikethrough: false,
    wrapText: false,
    diagonalUp: false,
    diagonalDown: false,
    ...flags,
  })
  const styles = [
    style(),
    style({ protectionLocked: false }),
    style({ protectionHidden: true }),
    style({ protectionLocked: false, protectionHidden: true }),
  ]
  const sheet = {
    columnWidths: [{ startColumn: 4, endColumn: 4, hidden: false, styleIndex: 2 }],
  } as WorkbookFile['sheets'][number]
  const range = { startRow: 0, endRow: 0, startColumn: 0, endColumn: 4 }
  const read = {
    cells: [
      { row: 0, column: 0, value: 7 },
      { row: 0, column: 1, value: 8, styleIndex: 1 },
      { row: 0, column: 2, value: 3, formula: '=1+2', styleIndex: 2 },
      { row: 0, column: 3, value: null, styleIndex: 3 },
    ],
    rows: [],
  }

  it('resolves unset defaults, opposite imported flags, mixed range, and column defaults', () => {
    const journal = createEditJournal()
    const result = selectionProtection(styles, sheet, read, journal, 'sheet-1', range)
    expect(result.cells.map(({ locked, hidden }) => [locked, hidden])).toEqual([
      [true, false],
      [false, false],
      [true, true],
      [false, true],
      [true, true],
    ])
    expect(result.locked).toBeNull()
    expect(result.hidden).toBeNull()
    const initial = draftFromSelection(SELECTION, result)
    expect(initial.locked).toBe('')
    expect(initial.hidden).toBe('')
    expect(formatCellsCommands(initial, initial)).toEqual([])
    expect(formatCellsCommands(initial, { ...initial, locked: 'off' })).toEqual([
      'cellprot:locked-off:',
    ])
  })

  it('backfills uniform states and overlays a partial journal edit without changing its opposite flag', () => {
    const journal = createEditJournal()
    const single = { startRow: 0, endRow: 0, startColumn: 2, endColumn: 2 }
    const imported = selectionProtection(styles, sheet, read, journal, 'sheet-1', single)
    expect(draftFromSelection(SELECTION, imported)).toMatchObject({ locked: 'on', hidden: 'on' })
    recordNeutralStyleEdit(journal, 'sheet-1', 0, 2, { protectionLocked: false })
    const edited = selectionProtection(styles, sheet, read, journal, 'sheet-1', single)
    expect(draftFromSelection(SELECTION, edited)).toMatchObject({ locked: 'off', hidden: 'on' })
    expect(read.cells[2]).toMatchObject({ formula: '=1+2', value: 3 })
  })

  it('clips the native read while resolving valid blank cells beyond the used extent', () => {
    const selected = { startRow: 4, endRow: 4, startColumn: 2, endColumn: 3 }
    expect(protectionReadRange(selected, 5, 3, false)).toEqual({
      startRow: 4,
      endRow: 4,
      startColumn: 2,
      endColumn: 2,
    })
    expect(protectionReadRange({ ...selected, startRow: 6, endRow: 7 }, 5, 3, false)).toEqual({
      startRow: 4,
      endRow: 4,
      startColumn: 2,
      endColumn: 2,
    })
    expect(protectionReadRange(selected, 5, 3, true)).toEqual(selected)
    const result = selectionProtection(
      styles,
      sheet,
      {
        cells: [{ row: 4, column: 2, value: null, styleIndex: 2 }],
        rows: [{ row: 4, hidden: false }],
      },
      createEditJournal(),
      'sheet-1',
      selected,
    )
    expect(result.cells.map(({ locked, hidden }) => [locked, hidden])).toEqual([
      [true, true],
      [true, false],
    ])
  })
})

describe('accounting underline document-renderer guard', () => {
  const normal: WorkbookCellStyle = {
    bold: false,
    italic: false,
    underline: false,
    strikethrough: false,
    wrapText: false,
    diagonalUp: false,
    diagonalDown: false,
  }
  const sheet = { columnWidths: [] } as unknown as WorkbookFile['sheets'][number]
  const single = { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }
  const selected = (value: string, rich = false, rotation = 0) =>
    selectionProtection(
      [{ ...normal, textRotation: rotation }],
      sheet,
      {
        cells: [
          {
            row: 0,
            column: 0,
            value,
            ...(rich
              ? {
                  rich: [
                    {
                      text: value,
                      bold: false,
                      italic: false,
                      underline: false,
                      strikethrough: false,
                    },
                  ],
                }
              : {}),
          },
        ],
        rows: [],
      },
      createEditJournal(),
      'sheet-1',
      single,
    )

  it('detects rich, multiline, and rotated cells while allowing plain unrotated text', () => {
    expect(accountingUnderlineUnavailable(selected('Plain'), '0')).toBe(false)
    expect(selected('Rich', true).cells[0]?.richDocument).toBe(true)
    expect(accountingUnderlineUnavailable(selected('Rich', true), '0')).toBe(true)
    expect(accountingUnderlineUnavailable(selected('One\nTwo'), '0')).toBe(true)
    expect(accountingUnderlineUnavailable(selected('Plain', false, 45), '45')).toBe(true)
    expect(accountingUnderlineUnavailable(selected('Plain', false, 45), '0')).toBe(false)
  })

  it('blocks a newly unsupported combination but permits clearing an imported mode', () => {
    const selection = selected('Plain')
    const initial = draftFromSelection(null, selection)
    const accounting = { ...initial, underlineMode: 'singleAccounting' as const }
    expect(accountingUnderlineChangeBlocked(initial, accounting, selection)).toBe(false)
    expect(
      accountingUnderlineChangeBlocked(initial, { ...accounting, rotation: '45' }, selection),
    ).toBe(true)
    const imported = selected('Rich', true)
    const importedDraft = {
      ...draftFromSelection(null, imported),
      underlineMode: 'singleAccounting' as const,
    }
    expect(accountingUnderlineChangeBlocked(importedDraft, importedDraft, imported)).toBe(false)
    expect(
      accountingUnderlineChangeBlocked(
        importedDraft,
        { ...importedDraft, underlineMode: 'single' },
        imported,
      ),
    ).toBe(false)
  })
})
