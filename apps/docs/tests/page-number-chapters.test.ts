import { describe, expect, it } from 'vitest'
import { parseDocx, sectionFromSectPr } from '@genoffice/docx-engine'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { pageNumberEditsAgainstBaseline } from '../src/renderer/editor/layout-history'
import { chapterNumbersForPages, formatPageNumber } from '../src/renderer/pagination-hf'

const sectPr =
  '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/><w:cols w:space="425"/></w:sectPr>'
const section = () => sectionFromSectPr(sectPr, 0, 0)

describe('chapter-aware page numbering', () => {
  it('formats the native Mac formats and chapter separators', () => {
    expect(formatPageNumber(2, 'decimalFullWidth')).toBe('２')
    expect(formatPageNumber(2, 'decimalEnclosedCircle')).toBe('②')
    expect(formatPageNumber(2, 'japaneseLegal')).toBe('弐')
    expect(formatPageNumber(2, 'ideographTraditional')).toBe('乙')
    expect(formatPageNumber(2, 'upperRoman', '3', 'period')).toBe('3.II')
    expect(formatPageNumber(2, 'decimal', '3', 'emDash')).toBe('3—2')
  })

  it('finds the latest numbered heading for each physical page', () => {
    const fakeEl = (style: string, marker: string) =>
      ({
        getAttribute: (name: string) =>
          name === 'data-style' ? style : name === 'data-marker' ? marker : null,
      }) as HTMLElement
    const blocks = [
      { top: 10, height: 10, el: fakeEl('Heading1', '1.') },
      { top: 210, height: 10, el: fakeEl('Heading1', '2.') },
    ]
    const slices = [
      { start: 0, end: 100, section: 0 },
      { start: 100, end: 200, section: 0 },
      { start: 200, end: 300, section: 0 },
    ]
    expect(
      chapterNumbersForPages(
        slices,
        blocks,
        [{ ...section(), pageNumberChapterStyle: 1 }],
        new Map([[1, 'Heading1']]),
      ),
    ).toEqual(['1', '1', '2'])
  })

  it('recomputes persistence edits against the saved baseline for Undo and Redo', () => {
    const original = section()
    const edited = {
      ...original,
      pageNumberFmt: 'upperRoman',
      pageNumberStart: 4,
      pageNumberChapterStyle: 1,
      pageNumberChapterSeparator: 'period',
    }
    expect(pageNumberEditsAgainstBaseline([edited], [original]).final).toEqual({
      fmt: 'upperRoman',
      start: 4,
      chapStyle: 1,
      chapSep: 'period',
    })
    // After Save, the original snapshot becomes a pending reverse patch.
    expect(pageNumberEditsAgainstBaseline([original], [edited]).final).toEqual({
      fmt: undefined,
      start: undefined,
      chapStyle: null,
      chapSep: null,
    })
    expect(pageNumberEditsAgainstBaseline([edited], [edited]).final).toBeNull()
    expect(pageNumberEditsAgainstBaseline([edited, original], [original, original])).toEqual({
      final: null,
      sections: [0],
    })
  })

  it('recognizes a numbered Heading style in a native DOCX', async () => {
    const parsed = await parseDocx(
      await buildDocx({
        withNumbering: true,
        extraStylesXml:
          '<w:style w:type="paragraph" w:styleId="Heading3"><w:name w:val="heading 3"/>' +
          '<w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr>' +
          '<w:outlineLvl w:val="2"/></w:pPr></w:style>',
        bodyXml:
          '<w:p><w:pPr><w:pStyle w:val="Heading3"/></w:pPr><w:r><w:t>Chapter</w:t></w:r></w:p>',
      }),
    )
    const styleId = parsed.headingStyleIds.get(3)
    const numbering = styleId ? parsed.styles.get(styleId)?.numPr : undefined
    expect(numbering && numbering !== 'none' && parsed.numbering.has(numbering.numId)).toBe(true)
  })
})
