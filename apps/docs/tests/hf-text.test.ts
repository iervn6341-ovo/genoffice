import { describe, expect, it } from 'vitest'
import { PAGE_MARK, TOTAL_PAGES_MARK, type HeaderFooter } from '@genoffice/docx-engine'
import {
  applyHfText,
  appendHfLayout,
  hfEditText,
  insertHfNativeField,
  insertHfPictureAtSelection,
} from '../src/renderer/editor/hf-text'
import { hfDisplayedRuns } from '../src/renderer/editor/hf-dom'

describe('hfEditText', () => {
  it('renders paragraphs as lines with field sentinels as tokens', () => {
    const value: HeaderFooter = {
      text: '',
      paras: [
        { align: 'left', runs: [{ text: 'Confidential' }] },
        { align: 'center', runs: [{ text: `Page ${PAGE_MARK} of ${TOTAL_PAGES_MARK}` }] },
      ],
    }
    expect(hfEditText(value)).toBe('Confidential\nPage {PAGE} of {NUMPAGES}')
  })

  it('falls back to the legacy single line and appends the page mark', () => {
    expect(hfEditText({ text: 'Draft', pageNumber: true })).toBe(`Draft {PAGE}`)
    expect(hfEditText({ text: '' })).toBe('')
  })

  it('skips layout-table rows', () => {
    const value: HeaderFooter = {
      text: '',
      paras: [
        { runs: [], cells: [{ paras: [[{ text: 'cell' }]] }] },
        { align: 'right', runs: [{ text: 'visible' }] },
      ],
    }
    expect(hfEditText(value)).toBe('visible')
  })
})

describe('applyHfText', () => {
  it('refreshes an authored dirty DATE on the next day with the selected native format', () => {
    const created = insertHfNativeField(
      null,
      '',
      0,
      0,
      'DATE',
      '',
      new Date(2026, 9, 13),
      'MM/dd/yyyy',
    )
    const run = created.value.paras?.[0]?.runs[0]
    expect(run?.hfField?.instruction).toBe(' DATE \\@ "MM/dd/yyyy" ')
    expect(run?.hfField?.xml).toContain('w:dirty="true"')
    expect(hfDisplayedRuns(run!, new Date(2026, 9, 14))[0].text).toBe('10/14/2026')
    const importedComplex = {
      ...run!,
      hfField: {
        ...run!.hfField!,
        xml: '<w:r><w:fldChar w:fldCharType="begin" w:dirty="true"/></w:r>',
      },
    }
    expect(hfDisplayedRuns(importedComplex, new Date(2026, 9, 14))[0].text).toBe('10/13/2026')
  })

  it('appends a real tabbed layout without replacing native fields', () => {
    const dated = insertHfNativeField(null, 'Before', 6, 6, 'DATE')
    const layout = appendHfLayout(dated.value, dated.text, 'threeColumn')
    expect(
      layout.value.paras?.[0]?.runs.some((run) => run.hfField?.instruction.includes('DATE')),
    ).toBe(true)
    expect(
      layout.value.paras
        ?.at(-1)
        ?.runs.map((run) => run.text)
        .join(''),
    ).toBe('\t\t')
    expect(layout.value.paras?.at(-1)?.tabStops).toEqual([
      { pos: 4680, val: 'center' },
      { pos: 9360, val: 'right' },
    ])
  })

  it('appends DATE and PAGE as native atoms in the date-page layout', () => {
    const layout = appendHfLayout(null, '', 'datePage', '', new Date('2026-10-13T00:00:00'))
    const fields = layout.value.paras?.flatMap((para) =>
      para.runs.flatMap((run) => run.hfField?.instruction ?? []),
    )
    expect(fields).toEqual([' DATE \\@ "yyyy-MM-dd" ', ' PAGE '])
    expect(layout.text).toContain('\t')
  })

  it('inserts a picture without dropping adjacent protected fields or text', () => {
    const dated = insertHfNativeField(null, 'Before After', 7, 7, 'DATE')
    const withPicture = insertHfPictureAtSelection(
      dated.value,
      dated.text,
      dated.text.length,
      dated.text.length,
      {
        base64: 'AQID',
        mime: 'image/png',
        widthPx: 120,
        heightPx: 40,
      },
    )
    expect(
      withPicture.value.paras
        ?.flatMap((para) => para.runs)
        .some((run) => run.hfField?.instruction.includes('DATE')),
    ).toBe(true)
    expect(
      withPicture.value.paras
        ?.flatMap((para) => para.runs)
        .some((run) => run.image?.pendingMedia?.base64 === 'AQID'),
    ).toBe(true)
    expect(withPicture.text).toContain('Before')
    expect(withPicture.text).toContain('After')
  })

  it('keeps per-line paragraph format and first-run styling', () => {
    const value: HeaderFooter = {
      text: '',
      paras: [
        { align: 'left', runs: [{ text: 'old', bold: true }] },
        { align: 'right', runs: [{ text: 'other' }] },
      ],
    }
    const next = applyHfText(value, 'new first\nnew second')
    expect(next.paras).toEqual([
      { align: 'left', runs: [{ text: 'new first', bold: true }] },
      { align: 'right', runs: [{ text: 'new second' }] },
    ])
    expect(next.text).toBe('new firstnew second')
  })

  it('an edit keeps the formatting of the runs it does not touch', () => {
    const value: HeaderFooter = {
      text: '',
      paras: [
        {
          align: 'left',
          runs: [{ text: 'Quarterly ' }, { text: 'Report', bold: true, color: 'FF0000' }],
        },
      ],
    }
    // typed at the end: continues the bold red run
    expect(applyHfText(value, 'Quarterly Report Q3').paras?.[0]?.runs).toEqual([
      { text: 'Quarterly ' },
      { text: 'Report Q3', bold: true, color: 'FF0000' },
    ])
    // one character changed in the plain part: the red word is untouched
    expect(applyHfText(value, 'Quartely Report').paras?.[0]?.runs).toEqual([
      { text: 'Quartely ' },
      { text: 'Report', bold: true, color: 'FF0000' },
    ])
    // typed at the start: takes the first character's formatting
    expect(applyHfText(value, 'Q1 Quarterly Report').paras?.[0]?.runs).toEqual([
      { text: 'Q1 Quarterly ' },
      { text: 'Report', bold: true, color: 'FF0000' },
    ])
    // a deletion across the boundary keeps both sides' formatting
    expect(applyHfText(value, 'Quarterport').paras?.[0]?.runs).toEqual([
      { text: 'Quarter' },
      { text: 'port', bold: true, color: 'FF0000' },
    ])
  })

  it('converts tokens back to field sentinels', () => {
    const next = applyHfText(null, '{PAGE} / {NUMPAGES}')
    expect(next.paras?.[0]?.runs[0]?.text).toBe(`${PAGE_MARK} / ${TOTAL_PAGES_MARK}`)
    expect(next.pageNumber).toBe(true)
    expect(applyHfText(next, '{NUMPAGES}').pageNumber).toBe(false)
  })

  it('splices cells rows back at their original positions', () => {
    const value: HeaderFooter = {
      text: '',
      paras: [
        { runs: [], cells: [{ paras: [[{ text: 'logo row' }]] }] },
        { align: 'center', runs: [{ text: 'title' }] },
      ],
    }
    const next = applyHfText(value, 'edited title')
    expect(next.paras?.[0]?.cells).toBeTruthy()
    expect(next.paras?.[1]?.runs[0]?.text).toBe('edited title')
  })

  it('extra lines reuse the last template; empty text clears', () => {
    const value: HeaderFooter = {
      text: '',
      paras: [{ align: 'right', runs: [{ text: 'one', italic: true }] }],
    }
    const next = applyHfText(value, 'a\nb')
    expect(next.paras).toEqual([
      { align: 'right', runs: [{ text: 'a', italic: true }] },
      { align: 'right', runs: [{ text: 'b', italic: true }] },
    ])
    const cleared = applyHfText(value, '')
    expect(cleared.text).toBe('')
    expect(cleared.paras).toEqual([{ align: 'right', runs: [] }])
  })
})

describe('native header/footer fields', () => {
  it('inserts DATE at the selected cursor and keeps it atomic while surrounding text changes', () => {
    const base: HeaderFooter = {
      text: 'Before  after',
      paras: [{ runs: [{ text: 'Before  after' }] }],
    }
    const inserted = insertHfNativeField(
      base,
      'Before  after',
      7,
      7,
      'DATE',
      '',
      new Date(2026, 9, 13),
    )
    expect(inserted.text).toBe('Before ⟦DATE·1⟧ after')
    expect(inserted.caret).toBe('Before ⟦DATE·1⟧'.length)
    const field = inserted.value.paras?.[0].runs.find((run) => run.hfField)?.hfField
    expect(field?.instruction).toBe(' DATE \\@ "yyyy-MM-dd" ')
    expect(field?.xml).toContain('2026-10-13')
    const next = applyHfText(inserted.value, inserted.text.replace('Before', 'Edited'))
    expect(next.paras?.[0].runs.find((run) => run.hfField)?.hfField).toEqual(field)
    expect(() => applyHfText(next, 'Edited 2026-10-13 after')).toThrow()
  })

  it('creates whitelisted PAGE, NUMPAGES, FILENAME and AUTHOR fields with native XML', () => {
    let value: HeaderFooter | null = null
    let text = ''
    for (const kind of ['PAGE', 'NUMPAGES', 'FILENAME', 'AUTHOR'] as const) {
      const inserted = insertHfNativeField(
        value,
        text,
        text.length,
        text.length,
        kind,
        'report.docx',
      )
      value = inserted.value
      text = inserted.text
    }
    expect(value?.paras?.[0].runs.filter((run) => run.hfField)).toHaveLength(4)
    expect(value?.pageNumber).toBe(true)
    expect(text).toContain('⟦FILENAME·1⟧')
    expect(
      value?.paras?.[0].runs.find((run) => run.hfField?.instruction.trim() === 'FILENAME')?.hfField
        ?.xml,
    ).toContain('report.docx')
    expect(
      value?.paras?.[0].runs.find((run) => run.hfField?.instruction.trim() === 'AUTHOR')?.hfField
        ?.xml,
    ).toContain('w:instr=" AUTHOR "')
  })
})
