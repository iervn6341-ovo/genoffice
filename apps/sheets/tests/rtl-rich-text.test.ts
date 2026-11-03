import { HorizontalAlign } from '@univerjs/core'
import { DocumentSkeleton, FontCache } from '@univerjs/engine-render'
import { describe, expect, it } from 'vitest'

import {
  bidiVisualOrder,
  classifyBidiGlyph,
  divideSelectionExtent,
  glyphBoundaryX,
  installRichTextBidiFix,
  logicalGlyphContent,
  paragraphIsRtl,
  REORDERED_EDITOR_IDS,
  reorderRichCellSkeleton,
  removeCjkLatinSpacingFromSkeleton,
  resolveBidiLevels,
} from '../src/renderer/rich-text-bidi-fix'

const AR_PRICE = 'السعر' // Arabic "price"
const AR_HEH = 'هـ' // heh + tatweel
const HE_SHALOM = 'שלום'
const AR_INDIC = '١٤٤٦' // Arabic-Indic 1446

const reverse = (text: string): string => [...text].reverse().join('')

// Univer's ArabicHandler emits one glyph per contiguous Arabic chunk with the
// characters already reversed; everything else is one glyph per character.
function univerGlyphs(words: string[]): { content: string; width: number; left: number }[] {
  let left = 0
  return words.map((content) => {
    const glyph = { content, width: content.length * 10, left }
    left += glyph.width
    return glyph
  })
}

function visualContents(glyphs: { content: string; left: number }[]): string[] {
  return [...glyphs].sort((a, b) => a.left - b.left).map((g) => g.content)
}

function makeSkeleton(
  id: string,
  dataStream: string,
  lines: { paragraphIndex?: number; glyphs: { content: string; width: number; left: number }[] }[],
) {
  return {
    getViewModel: () => ({
      getDataModel: () => ({ getSnapshot: () => ({ id, body: { dataStream } }) }),
    }),
    getSkeletonData: () => ({
      pages: [
        {
          sections: [
            {
              columns: [
                {
                  lines: lines.map((line) => ({
                    paragraphIndex: line.paragraphIndex,
                    divides: [{ glyphGroup: line.glyphs }],
                  })),
                },
              ],
            },
          ],
        },
      ],
    }),
  }
}

describe('logicalGlyphContent', () => {
  it('un-reverses multi-char Arabic chunks (ArabicHandler order)', () => {
    expect(logicalGlyphContent(reverse(AR_PRICE))).toBe(AR_PRICE)
  })

  it('leaves single chars, Hebrew and Latin untouched', () => {
    expect(logicalGlyphContent('ع')).toBe('ع')
    expect(logicalGlyphContent(HE_SHALOM)).toBe(HE_SHALOM)
    expect(logicalGlyphContent('abc')).toBe('abc')
    expect(logicalGlyphContent(`${AR_PRICE} `)).toBe(`${AR_PRICE} `)
  })
})

describe('classifyBidiGlyph', () => {
  it('classifies scripts and digits', () => {
    expect(classifyBidiGlyph(AR_PRICE)).toBe('AL')
    expect(classifyBidiGlyph(HE_SHALOM.slice(0, 1))).toBe('R')
    expect(classifyBidiGlyph('a')).toBe('L')
    expect(classifyBidiGlyph('7')).toBe('EN')
    expect(classifyBidiGlyph(AR_INDIC)).toBe('AN')
    expect(classifyBidiGlyph(' ')).toBe('WS')
    expect(classifyBidiGlyph('(')).toBe('ON')
  })

  it('lets a strong letter win over leading digits in a mixed chunk', () => {
    expect(classifyBidiGlyph(`${AR_INDIC}${AR_PRICE}`)).toBe('AL')
  })

  it('honors directional marks', () => {
    expect(classifyBidiGlyph('‏')).toBe('R')
    expect(classifyBidiGlyph('‎')).toBe('L')
  })

  it('treats Script=Common Arabic punctuation and tatweel as AL', () => {
    // U+0640 is a modifier LETTER of Script=Common: without the explicit
    // class it would read as strong LTR and detach from its Hijri "هـ"
    expect(classifyBidiGlyph('\u0640')).toBe('AL')
    expect(classifyBidiGlyph('\u061F')).toBe('AL')
    expect(classifyBidiGlyph('\u061B')).toBe('AL')
  })

  it('keeps the Hijri suffix together after a year in an RTL paragraph', () => {
    // "1446هـ" glyph-per-char: EN EN EN EN AL AL — the suffix is one RTL run
    // seated left of the digits, tatweel closest to the left edge
    const levels = resolveBidiLevels(
      ['1', '4', '4', '6', 'ه', '\u0640'].map(classifyBidiGlyph),
      true,
    )
    expect(levels.slice(4)).toEqual([1, 1])
    expect(bidiVisualOrder(levels)).toEqual([5, 4, 0, 1, 2, 3])
  })
})

describe('resolveBidiLevels + bidiVisualOrder', () => {
  it('keeps pure LTR text in order', () => {
    const levels = resolveBidiLevels(['L', 'WS', 'EN', 'WS', 'L'], false)
    expect(bidiVisualOrder(levels)).toEqual([0, 1, 2, 3, 4])
  })

  it('reorders an RTL paragraph with embedded European digits', () => {
    // AL WS EN EN WS AL — digits stay LTR inside the reversed line
    const levels = resolveBidiLevels(['AL', 'WS', 'EN', 'EN', 'WS', 'AL'], true)
    expect(bidiVisualOrder(levels)).toEqual([5, 4, 2, 3, 1, 0])
  })

  it('swaps adjacent Arabic words inside an LTR paragraph', () => {
    const levels = resolveBidiLevels(['L', 'WS', 'AL', 'WS', 'AL', 'WS', 'L'], false)
    expect(bidiVisualOrder(levels)).toEqual([0, 1, 4, 3, 2, 5, 6])
  })

  it('resolves neutrals between RTL letters to RTL', () => {
    const levels = resolveBidiLevels(['R', 'ON', 'R'], true)
    expect(levels).toEqual([1, 1, 1])
  })

  // W2 seeds from sos (R in an RTL paragraph, never AL): matches python-bidi
  // for both scripts.
  it('keeps EN digits and their percent sign together in a Hebrew paragraph', () => {
    // "shalom 10%" — % joins the digits (W5 on EN), all level 2
    expect(resolveBidiLevels(['R', 'WS', 'EN', 'EN', 'ET'], true)).toEqual([1, 1, 2, 2, 2])
    // leading digits with no strong letter before them stay EN too
    expect(resolveBidiLevels(['EN', 'EN', 'ET', 'WS', 'R'], true)).toEqual([2, 2, 2, 1, 1])
  })

  it('turns EN after an Arabic letter into AN, detaching the percent sign', () => {
    // "<AR> 10%" — W2 makes the digits AN, so W5 no longer captures the ET
    expect(resolveBidiLevels(['AL', 'WS', 'EN', 'EN', 'ET'], true)).toEqual([1, 1, 2, 2, 1])
  })

  it('keeps an NBSP-separated number as one run in an RTL paragraph', () => {
    // NBSP is CS, not WS — "1<NBSP>234" must not split (W4), per python-bidi
    expect(classifyBidiGlyph('\u00A0')).toBe('CS')
    expect(resolveBidiLevels(['R', 'WS', 'EN', 'CS', 'EN', 'EN', 'EN'], true)).toEqual([
      1, 1, 2, 2, 2, 2, 2,
    ])
  })
})

describe('reorderRichCellSkeleton', () => {
  it('restores Arabic glyph content and seats an RTL line in visual order', () => {
    const glyphs = univerGlyphs([reverse(AR_PRICE), ' ', '1', '4', '4', '6', ' ', reverse(AR_HEH)])
    const dataStream = `${AR_PRICE} 1446 ${AR_HEH}\r\n`
    const skeleton = makeSkeleton('rich-cell', dataStream, [
      { paragraphIndex: dataStream.length - 2, glyphs },
    ])
    reorderRichCellSkeleton(skeleton)
    expect(visualContents(glyphs)).toEqual([AR_HEH, ' ', '1', '4', '4', '6', ' ', AR_PRICE])
    expect(glyphs[0]?.content).toBe(AR_PRICE)
    // widths travel with their glyphs, so run styling stays attached
    expect(glyphs.reduce((sum, g) => Math.max(sum, g.left + g.width), 0)).toBe(130)
  })

  it('mirrors paired brackets on RTL runs', () => {
    const he = [...HE_SHALOM]
    const glyphs = univerGlyphs([...he, ' ', '(', ...he, ')', ' ', '1', '2', '3'])
    const dataStream = `${HE_SHALOM} (${HE_SHALOM}) 123\r\n`
    const skeleton = makeSkeleton('rich-cell', dataStream, [
      { paragraphIndex: dataStream.length - 2, glyphs },
    ])
    reorderRichCellSkeleton(skeleton)
    const visual = visualContents(glyphs)
    expect(visual.slice(0, 4)).toEqual(['1', '2', '3', ' '])
    expect(visual[4]).toBe('(')
    expect(visual[9]).toBe(')')
    expect(visual.slice(5, 9)).toEqual([...HE_SHALOM].reverse())
    expect(visual.slice(11)).toEqual([...HE_SHALOM].reverse())
  })

  it('is idempotent across repeated calculate passes', () => {
    const glyphs = univerGlyphs([reverse(AR_PRICE), ' ', reverse(AR_HEH)])
    const dataStream = `${AR_PRICE} ${AR_HEH}\r\n`
    const skeleton = makeSkeleton('rich-cell', dataStream, [
      { paragraphIndex: dataStream.length - 2, glyphs },
    ])
    reorderRichCellSkeleton(skeleton)
    const after = glyphs.map((g) => ({ ...g }))
    reorderRichCellSkeleton(skeleton)
    expect(glyphs).toEqual(after)
  })

  it('reprocesses a line whose glyphs were rebuilt by a relayout', () => {
    const line: {
      paragraphIndex?: number
      glyphs: { content: string; width: number; left: number }[]
    } = { glyphs: univerGlyphs([reverse(AR_PRICE), ' ', reverse(AR_HEH)]) }
    const dataStream = `${AR_PRICE} ${AR_HEH}\r\n`
    line.paragraphIndex = dataStream.length - 2
    const skeleton = makeSkeleton('rich-cell', dataStream, [line])
    reorderRichCellSkeleton(skeleton)
    // simulate an incremental relayout handing out fresh reversed glyphs
    line.glyphs = univerGlyphs([reverse(AR_PRICE), ' ', reverse(AR_HEH)])
    reorderRichCellSkeleton(skeleton)
    expect(visualContents(line.glyphs)).toEqual([AR_HEH, ' ', AR_PRICE])
  })

  it('reorders the in-cell editor and the formula bar like the cell', () => {
    for (const id of REORDERED_EDITOR_IDS) {
      const glyphs = univerGlyphs([reverse(AR_PRICE), ' ', reverse(AR_HEH)])
      const dataStream = `${AR_PRICE} ${AR_HEH}\r\n`
      const skeleton = makeSkeleton(id, dataStream, [
        { paragraphIndex: dataStream.length - 2, glyphs },
      ])
      reorderRichCellSkeleton(skeleton)
      expect(visualContents(glyphs)).toEqual([AR_HEH, ' ', AR_PRICE])
      expect(glyphs[0]?.content).toBe(AR_PRICE)
    }
  })

  it('ignores other internal editors', () => {
    const glyphs = univerGlyphs([reverse(AR_PRICE), ' ', reverse(AR_HEH)])
    const before = glyphs.map((g) => ({ ...g }))
    const skeleton = makeSkeleton('__INTERNAL_EDITOR__DOCS_ZEN', `${AR_PRICE} ${AR_HEH}\r\n`, [
      { glyphs },
    ])
    reorderRichCellSkeleton(skeleton)
    expect(glyphs).toEqual(before)
  })

  it('keeps an edited formula LTR-based while flipping its Arabic literal', () => {
    // =A1&"<AR> <HEH>" — the reference and operators stay in reading order,
    // only the Arabic literal flips (an RTL base would reverse the whole line)
    const glyphs = univerGlyphs([
      '=',
      'A',
      '1',
      '&',
      '"',
      reverse(AR_PRICE),
      ' ',
      reverse(AR_HEH),
      '"',
    ])
    const dataStream = `=A1&"${AR_PRICE} ${AR_HEH}"\r\n`
    const editor = makeSkeleton('__INTERNAL_EDITOR__DOCS_NORMAL', dataStream, [
      { paragraphIndex: dataStream.length - 2, glyphs },
    ])
    reorderRichCellSkeleton(editor)
    expect(visualContents(glyphs)).toEqual(['=', 'A', '1', '&', '"', AR_HEH, ' ', AR_PRICE, '"'])
    // formula rule: '=' pins the base even when the first strong char is Arabic
    const literal = `="${AR_PRICE}"\r\n`
    expect(paragraphIsRtl(literal, literal.length - 2, true)).toBe(false)
    // the same text painted in a cell is a plain string: first strong char rules
    expect(paragraphIsRtl(literal, literal.length - 2)).toBe(true)
    const text = `${AR_PRICE}=1\r\n`
    expect(paragraphIsRtl(text, text.length - 2, true)).toBe(true)
  })

  it('leaves LTR-only rich documents untouched', () => {
    const glyphs = univerGlyphs(['T', 'o', 't', 'a', 'l', ' ', '1', '4'])
    const before = glyphs.map((g) => ({ ...g }))
    const skeleton = makeSkeleton('rich-cell', 'Total 14\r\n', [{ paragraphIndex: 8, glyphs }])
    reorderRichCellSkeleton(skeleton)
    expect(glyphs).toEqual(before)
  })

  it('reorders an RTL run with trailing digits inside an LTR paragraph', () => {
    // UAX#9: the digits become AN after the Arabic word and sit LEFT of it
    // ("Pr 10 <AR>"), verified against python-bidi.
    const glyphs = univerGlyphs(['P', 'r', ' ', reverse(AR_PRICE), ' ', '1', '0'])
    const dataStream = `Pr ${AR_PRICE} 10\r\n`
    const skeleton = makeSkeleton('rich-cell', dataStream, [
      { paragraphIndex: dataStream.length - 2, glyphs },
    ])
    reorderRichCellSkeleton(skeleton)
    expect(visualContents(glyphs)).toEqual(['P', 'r', ' ', '1', '0', ' ', AR_PRICE])
  })
})

describe('editor caret geometry over reordered glyphs', () => {
  // Hebrew "שלום" typed into the editor: logical glyph i sits at visual slot
  // 3-i, every glyph 10px wide, line anchored at x=0.
  const he = [...HE_SHALOM].map((content, i) => ({ content, width: 10, left: (3 - i) * 10 }))
  const rtl = () => true

  it('puts the caret before an RTL glyph on its visual right edge', () => {
    expect(glyphBoundaryX(he[0]!, true, true)).toBe(40)
    expect(glyphBoundaryX(he[0]!, false, true)).toBe(30)
    // LTR glyphs keep the stock mapping
    expect(glyphBoundaryX({ left: 5, width: 10 }, true, false)).toBe(5)
    expect(glyphBoundaryX({ left: 5, width: 10 }, false, false)).toBe(15)
  })

  it('collapses a caret to the mapped boundary', () => {
    // offset 0 (before the first letter) = far right of the word
    expect(divideSelectionExtent(he, 0, true, 0, true, rtl)).toEqual({ startX: 40, endX: 40 })
    // after the last letter = far left
    expect(divideSelectionExtent(he, 3, false, 3, false, rtl)).toEqual({ startX: 0, endX: 0 })
  })

  it('boxes a logical range by the union of its glyphs', () => {
    // select the first two letters: visually the rightmost 20px
    expect(divideSelectionExtent(he, 0, true, 1, false, rtl)).toEqual({ startX: 20, endX: 40 })
    // Univer's "end before glyph 2" form of the same range
    expect(divideSelectionExtent(he, 0, true, 2, true, rtl)).toEqual({ startX: 20, endX: 40 })
    // whole word
    expect(divideSelectionExtent(he, 0, true, 3, false, rtl)).toEqual({ startX: 0, endX: 40 })
  })

  it('spans a mixed-direction range with one bounding box', () => {
    // "ab" + Hebrew word: a b at 0..20, word reversed at 20..60
    const mixed = [
      { width: 10, left: 0 },
      { width: 10, left: 10 },
      ...[...HE_SHALOM].map((_ch, i) => ({ width: 10, left: 20 + (3 - i) * 10 })),
    ]
    const isRtl = (glyph: { left: number }) => glyph.left >= 20
    expect(divideSelectionExtent(mixed, 1, true, 3, false, isRtl)).toEqual({ startX: 10, endX: 60 })
  })
})

describe('Arabic chunk measurement during layout', () => {
  it('measures the logical string only while a reordered document lays out', () => {
    const seen: string[] = []
    const fontCache = FontCache as unknown as {
      getTextSize(content: string, fontStyle: unknown): unknown
    }
    fontCache.getTextSize = (content: string) => {
      seen.push(content)
      return { width: content.length }
    }
    // stand-in for Univer's layout: ArabicHandler measures the reversed chunk
    const proto = DocumentSkeleton.prototype as unknown as { calculate(this: unknown): void }
    proto.calculate = function () {
      fontCache.getTextSize(reverse(AR_PRICE), {})
    }
    installRichTextBidiFix()

    proto.calculate.call(makeSkeleton('rich-cell', `${AR_PRICE}\r\n`, []))
    proto.calculate.call(makeSkeleton(REORDERED_EDITOR_IDS[0]!, `${AR_PRICE}\r\n`, []))
    // a document this module does not reorder keeps Univer's own widths
    proto.calculate.call(makeSkeleton('zen-editor', `${AR_PRICE}\r\n`, []))
    // whole-string callers outside layout (dropdown items) are untouched
    fontCache.getTextSize(reverse(AR_PRICE), {})
    // single letters and non-Arabic text never change
    proto.calculate = function () {
      fontCache.getTextSize('a', {})
      fontCache.getTextSize(reverse(HE_SHALOM), {})
    }
    proto.calculate.call(makeSkeleton('rich-cell', `${HE_SHALOM}\r\n`, []))

    expect(seen).toEqual([
      AR_PRICE,
      AR_PRICE,
      reverse(AR_PRICE),
      reverse(AR_PRICE),
      'a',
      reverse(HE_SHALOM),
    ])
  })
})

describe('removeCjkLatinSpacingFromSkeleton', () => {
  it('restores font advance and eliminates extra CJK-Latin spacing in editor', () => {
    // Simulate mixed CJK-Latin string where CJK preceding Latin has inflated width,
    // and CJK following Latin has inflated width and xOffset.
    const glyphs = [
      { content: '課', width: 14, left: 0, bBox: { width: 14 } },
      { content: '綱', width: 14, left: 14, bBox: { width: 14 } },
      { content: '導', width: 14, left: 28, bBox: { width: 14 } },
      { content: '入', width: 17.5, left: 42, bBox: { width: 14 } }, // width inflated by width/4
      { content: 'A', width: 9, left: 59.5, bBox: { width: 9 } },
      { content: 'I', width: 4, left: 68.5, bBox: { width: 4 } },
      { content: '數', width: 17.5, left: 72.5, xOffset: 3.5, bBox: { width: 14 } }, // inflated width & xOffset
      { content: '量', width: 14, left: 90, bBox: { width: 14 } },
    ]

    const divide = {
      glyphGroup: glyphs,
      glyphGroupWidth: 104,
      width: 200,
      paddingLeft: 0,
    }
    const line = { divides: [divide], width: 104, paragraphIndex: 0 }
    const skeleton = {
      getSkeletonData: () => ({
        pages: [{ sections: [{ columns: [{ lines: [line] }] }] }],
      }),
    }

    removeCjkLatinSpacingFromSkeleton(skeleton as any)

    // Check that inflated widths and xOffsets are restored
    expect(glyphs[3]!.width).toBe(14)
    expect((glyphs[3] as any).xOffset).toBe(0)
    expect(glyphs[6]!.width).toBe(14)
    expect((glyphs[6] as any).xOffset).toBe(0)

    // Check that left positions are recalculated continuously
    expect(glyphs[0]!.left).toBe(0)
    expect(glyphs[1]!.left).toBe(14)
    expect(glyphs[2]!.left).toBe(28)
    expect(glyphs[3]!.left).toBe(42)
    expect(glyphs[4]!.left).toBe(56) // 42 + 14 = 56 (was 59.5!)
    expect(glyphs[5]!.left).toBe(65) // 56 + 9 = 65 (was 68.5!)
    expect(glyphs[6]!.left).toBe(69) // 65 + 4 = 69 (was 72.5!)
    expect(glyphs[7]!.left).toBe(83) // 69 + 14 = 83 (was 90!)

    // Total width is restored to 83 + 14 = 97 (was 104)
    expect(divide.glyphGroupWidth).toBe(97)
    expect(line.width).toBe(97)
  })

  it('keeps divide.paddingLeft at 0 for left-aligned lines nearly filling the divide', () => {
    // When a wrapped line almost fills the divide (e.g. divide.width = 100,
    // divide.glyphGroupWidth was 99), divide.paddingLeft was 0 before fix.
    // The previous heuristic Math.abs(divide.paddingLeft - (width - glyphGroupWidth) / 2) < 1
    // falsely treated this left-aligned line as center-aligned, causing a ~3px indentation jump.
    const glyphs = [
      { content: 'A', width: 9, left: 0, bBox: { width: 9 } },
      { content: 'I', width: 4, left: 9, bBox: { width: 4 } },
      { content: '驅', width: 17.5, left: 13, xOffset: 3.5, bBox: { width: 14 } },
      { content: '動', width: 14, left: 30.5, bBox: { width: 14 } },
      { content: '金', width: 14, left: 44.5, bBox: { width: 14 } },
      { content: '融', width: 14, left: 58.5, bBox: { width: 14 } },
      { content: '商', width: 14, left: 72.5, bBox: { width: 14 } },
      { content: '管', width: 14, left: 86.5, bBox: { width: 14 } },
    ]
    const divide = {
      glyphGroup: glyphs,
      glyphGroupWidth: 100.5,
      width: 101, // divide width is only 0.5px wider than glyphGroupWidth
      paddingLeft: 0,
    }
    const line = { divides: [divide], width: 100.5, paragraphIndex: 0 }
    const skeleton = {
      getSkeletonData: () => ({
        pages: [{ sections: [{ columns: [{ lines: [line] }] }] }],
      }),
      getViewModel: () => ({
        getParagraph: () => ({ paragraphStyle: { horizontalAlign: HorizontalAlign.LEFT } }),
        getDataModel: () => ({ getSnapshot: () => ({ id: 'rich-cell' }) }),
      }),
    }

    removeCjkLatinSpacingFromSkeleton(skeleton as any)

    // Left alignment must be strictly preserved: paddingLeft must remain 0
    expect(divide.paddingLeft).toBe(0)
    expect(glyphs[0]!.left).toBe(0)
  })

  it('correctly updates divide.paddingLeft for center and right aligned text', () => {
    const makeLine = (hAlign: HorizontalAlign) => {
      const glyphs = [
        { content: 'A', width: 9, left: 0, bBox: { width: 9 } },
        { content: '文', width: 17.5, left: 9, xOffset: 3.5, bBox: { width: 14 } },
      ]
      const divide = {
        glyphGroup: glyphs,
        glyphGroupWidth: 26.5,
        width: 100,
        paddingLeft: hAlign === HorizontalAlign.CENTER ? (100 - 26.5) / 2 : 100 - 26.5,
      }
      const line = { divides: [divide], width: 26.5, paragraphIndex: 0 }
      const skeleton = {
        getSkeletonData: () => ({
          pages: [{ sections: [{ columns: [{ lines: [line] }] }] }],
        }),
        getViewModel: () => ({
          getParagraph: () => ({ paragraphStyle: { horizontalAlign: hAlign } }),
          getDataModel: () => ({ getSnapshot: () => ({ id: 'rich-cell' }) }),
        }),
      }
      removeCjkLatinSpacingFromSkeleton(skeleton as any)
      return { glyphs, divide }
    }

    // New glyphGroupWidth = 9 + 14 = 23
    const centerResult = makeLine(HorizontalAlign.CENTER)
    expect(centerResult.divide.glyphGroupWidth).toBe(23)
    expect(centerResult.divide.paddingLeft).toBe((100 - 23) / 2) // 38.5

    const rightResult = makeLine(HorizontalAlign.RIGHT)
    expect(rightResult.divide.glyphGroupWidth).toBe(23)
    expect(rightResult.divide.paddingLeft).toBe(100 - 23) // 77
  })

  it('resets xOffset and scales glyph back to bBox even when adjustGlyphsInDivide shrank width', () => {
    // When a line wraps, adjustGlyphsInDivide may shrink glyph.width down to <= bBox.width,
    // but xOffset remains non-zero.
    const glyphs = [
      { content: 'A', width: 9, left: 0, bBox: { width: 9 } },
      { content: '字', width: 13.8, left: 9, xOffset: 3.2, bBox: { width: 14 } },
    ]
    const divide = {
      glyphGroup: glyphs,
      glyphGroupWidth: 22.8,
      width: 50,
      paddingLeft: 0,
    }
    const line = { divides: [divide], width: 22.8, paragraphIndex: 0 }
    const skeleton = {
      getSkeletonData: () => ({
        pages: [{ sections: [{ columns: [{ lines: [line] }] }] }],
      }),
    }

    removeCjkLatinSpacingFromSkeleton(skeleton as any)

    expect(glyphs[1]!.width).toBe(14)
    expect((glyphs[1] as any).xOffset).toBe(0)
    expect(glyphs[1]!.left).toBe(9)
    expect(divide.glyphGroupWidth).toBe(23)
  })
})

