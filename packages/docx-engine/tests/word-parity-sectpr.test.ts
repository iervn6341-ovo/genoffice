import { describe, expect, it } from 'vitest'

import { applySectionSettings, lineNumberingOf, sectionSettingsFromXml } from '../src/section'

/**
 * Section settings against Word 16 itself: each fixture is the w:sectPr Word wrote (Word.Application
 * through COM, SaveAs2 to .docx) after the same ribbon operation — Line Numbers, Page Borders ▸ Box,
 * Columns ▸ Two with a line between. The engine reads what Word wrote and writes the same thing.
 */
// the blank A4 document of a zh-TW Word: margins 1" / 1.25", header 851, footer 992, a 425 column gap
const BLANK =
  '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>' +
  '<w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="851" w:footer="992" w:gutter="0"/>' +
  '<w:cols w:space="425"/><w:docGrid w:type="lines" w:linePitch="360"/></w:sectPr>'

// Word: Line Numbers ▸ Continuous / Restart Each Section / Restart Each Page (COM RestartMode 0 / 1 / 2)
const WORD_LINE_NUMBERS = {
  continuous: '<w:lnNumType w:countBy="1" w:restart="continuous"/>',
  newSection: '<w:lnNumType w:countBy="1" w:restart="newSection"/>',
  // Restart Each Page is the schema default: Word writes no restart attribute
  newPage: '<w:lnNumType w:countBy="1"/>',
}

// Word: Page Borders ▸ Box (defaults: single, 1/2 pt, auto color, 24 pt from the page edge)
const WORD_PAGE_BORDERS =
  '<w:pgBorders w:offsetFrom="page">' +
  '<w:top w:val="single" w:sz="4" w:space="24" w:color="auto"/>' +
  '<w:left w:val="single" w:sz="4" w:space="24" w:color="auto"/>' +
  '<w:bottom w:val="single" w:sz="4" w:space="24" w:color="auto"/>' +
  '<w:right w:val="single" w:sz="4" w:space="24" w:color="auto"/></w:pgBorders>'

// Word: Columns ▸ Two, Line between
const WORD_TWO_COLUMNS = '<w:cols w:num="2" w:sep="1" w:space="425"/>'

const settingsOf = (xml: string) => sectionSettingsFromXml(xml, {})

describe('line numbering reads what Word writes', () => {
  it('Word writes no restart attribute for "Restart Each Page", and the engine reads it as newPage', () => {
    for (const [restart, tag] of Object.entries(WORD_LINE_NUMBERS)) {
      expect(lineNumberingOf(tag), restart).toMatchObject({ countBy: 1, start: 1, restart })
    }
  })

  it('writes each restart mode the way Word reads it (the explicit newPage is the same as no attribute)', () => {
    for (const [restart, wordTag] of Object.entries(WORD_LINE_NUMBERS)) {
      const out = applySectionSettings(BLANK, {
        ...settingsOf(BLANK),
        lineNumbers: { countBy: 1, start: 1, restart: restart as 'newPage' },
      })
      const written = /<w:lnNumType[^>]*\/>/.exec(out)![0]
      // read back, it is what Word's own element says
      expect(lineNumberingOf(written)).toEqual(lineNumberingOf(wordTag))
      // and it sits where Word puts it: after pgMar, before cols
      expect(out.indexOf('<w:pgMar')).toBeLessThan(out.indexOf('<w:lnNumType'))
      expect(out.indexOf('<w:lnNumType')).toBeLessThan(out.indexOf('<w:cols'))
    }
  })

  it('a start number is Word\'s "start at" minus one (w:start is the count of lines skipped)', () => {
    const tag = '<w:lnNumType w:countBy="2" w:start="4" w:distance="360" w:restart="continuous"/>'
    expect(lineNumberingOf(tag)).toEqual({
      countBy: 2,
      start: 5,
      distance: 360,
      restart: 'continuous',
    })
    const out = applySectionSettings(BLANK, {
      ...settingsOf(BLANK),
      lineNumbers: { countBy: 2, start: 5, distance: 360, restart: 'continuous' },
    })
    expect(out).toContain(
      '<w:lnNumType w:countBy="2" w:start="4" w:distance="360" w:restart="continuous"/>',
    )
  })
})

describe('page borders and columns write what Word writes', () => {
  it('Page Borders ▸ Box: Word’s element, byte for byte, between pgMar and cols', () => {
    const out = applySectionSettings(BLANK, { ...settingsOf(BLANK), pageBorder: true })
    expect(out).toContain(WORD_PAGE_BORDERS)
    expect(out.indexOf('<w:pgBorders')).toBeGreaterThan(out.indexOf('<w:pgMar'))
    expect(out.indexOf('<w:pgBorders')).toBeLessThan(out.indexOf('<w:cols'))
  })

  it('reads Word’s Box back as a visible page border measured from the page edge', () => {
    const xml = BLANK.replace('<w:cols', `${WORD_PAGE_BORDERS}<w:cols`)
    expect(settingsOf(xml).pageBorder).toBe(true)
    expect(settingsOf(xml).pageBorderProps).toMatchObject({ offsetFrom: 'page', spacePt: 24 })
  })

  it('Columns ▸ Two with a line between: two columns, the gap Word kept, a separator', () => {
    const out = applySectionSettings(BLANK, { ...settingsOf(BLANK), columns: 2, colSep: true })
    const cols = /<w:cols[^>]*\/>/.exec(out)![0]
    expect(cols).toContain('w:num="2"')
    expect(cols).toContain('w:sep="1"')
    expect(cols).toContain('w:space="425"')
    // Word's own element says the same
    const word = settingsOf(BLANK.replace('<w:cols w:space="425"/>', WORD_TWO_COLUMNS))
    expect(word.columns).toBe(2)
    expect(word.colSep).toBe(true)
    expect(word.colSpace).toBe(425)
  })
})
