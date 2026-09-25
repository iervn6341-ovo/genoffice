import { describe, expect, it } from 'vitest'
import { PAGE_MARK, TOTAL_PAGES_MARK, type HeaderFooter } from '@genoffice/docx-engine'
import { applyHfText, hfEditText } from '../src/renderer/editor/hf-text'

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
