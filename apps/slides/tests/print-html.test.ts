import { describe, expect, it } from 'vitest'
import {
  buildPrintDocumentHtml,
  normalizePrintRatio,
  parsePrintRange,
  printPageCount,
} from '../src/shared/print-html'

const srcs = (n: number) => Array.from({ length: n }, (_x, i) => `blob:img-${i}`)

describe('parsePrintRange', () => {
  it('parses singles, ranges and mixed separators, deduped and sorted', () => {
    expect(parsePrintRange('1,3,5-8', 10)).toEqual([0, 2, 4, 5, 6, 7])
    expect(parsePrintRange('3，1、2 5;4', 5)).toEqual([0, 1, 2, 3, 4])
    expect(parsePrintRange('2-4, 3', 5)).toEqual([1, 2, 3])
  })

  it('rejects empty, malformed, out-of-bounds and inverted input', () => {
    expect(parsePrintRange('', 5)).toBeNull()
    expect(parsePrintRange('a-b', 5)).toBeNull()
    expect(parsePrintRange('0-2', 5)).toBeNull()
    expect(parsePrintRange('4-2', 5)).toBeNull()
    expect(parsePrintRange('6', 5)).toBeNull()
  })

  it('accepts spaces around dashes', () => {
    expect(parsePrintRange('1 - 3', 5)).toEqual([0, 1, 2])
    expect(parsePrintRange('1 – 3, 5', 5)).toEqual([0, 1, 2, 4])
  })

  it('accepts em-dash, tilde, fullwidth tilde and wave dash (IME variants)', () => {
    expect(parsePrintRange('1—3', 5)).toEqual([0, 1, 2])
    expect(parsePrintRange('1~3', 5)).toEqual([0, 1, 2])
    expect(parsePrintRange('1～3', 5)).toEqual([0, 1, 2])
    expect(parsePrintRange('1〜3', 5)).toEqual([0, 1, 2])
    expect(parsePrintRange('1 — 3, 5～6', 6)).toEqual([0, 1, 2, 4, 5])
  })

  it('accepts spaces around dashes', () => {
    expect(parsePrintRange('1 - 3', 5)).toEqual([0, 1, 2])
    expect(parsePrintRange('1 – 3, 5', 5)).toEqual([0, 1, 2, 4])
  })
})

describe('printPageCount', () => {
  it('follows the per-page thumbnail count of each layout', () => {
    expect(printPageCount(7, 'full')).toBe(7)
    expect(printPageCount(7, 'notes')).toBe(7)
    expect(printPageCount(7, 'handout2')).toBe(4)
    expect(printPageCount(7, 'handout3')).toBe(3)
    expect(printPageCount(7, 'handout6')).toBe(2)
  })
})

describe('buildPrintDocumentHtml', () => {
  it('full layout sizes pages by the slide ratio and can frame slides', () => {
    const html = buildPrintDocumentHtml({ srcs: srcs(2), ratio: 16 / 9, layout: 'full' })
    expect(html).toContain('@page { size: 13.333in 7.5in; margin: 0; }')
    expect(html.match(/class="page"/g)).toHaveLength(2)
    expect(html).not.toContain('.page > img { border')

    const framed = buildPrintDocumentHtml({
      srcs: srcs(1),
      ratio: 16 / 9,
      layout: 'full',
      frame: true,
    })
    expect(framed).toContain('.page > img { border: 1px solid #bbb; }')
  })

  it('handouts use A4 and swap dimensions in landscape', () => {
    const portrait = buildPrintDocumentHtml({ srcs: srcs(5), ratio: 16 / 9, layout: 'handout6' })
    expect(portrait).toContain('@page { size: 8.27in 11.69in; margin: 0; }')
    expect(portrait).toContain('grid-template-columns: 1fr 1fr;')

    const landscape = buildPrintDocumentHtml({
      srcs: srcs(5),
      ratio: 16 / 9,
      layout: 'handout6',
      orientation: 'landscape',
    })
    expect(landscape).toContain('@page { size: 11.69in 8.27in; margin: 0; }')
    expect(landscape).toContain('grid-template-columns: 1fr 1fr 1fr;')
  })

  it('notes layout escapes note text and pairs it with each slide', () => {
    const html = buildPrintDocumentHtml({
      srcs: srcs(2),
      ratio: 4 / 3,
      layout: 'notes',
      notes: ['a < b\nnext', ''],
    })
    expect(html.match(/class="page notes"/g)).toHaveLength(2)
    expect(html).toContain('a &lt; b<br>next')
  })

  it('escapes image sources for attribute context in every layout', () => {
    const full = buildPrintDocumentHtml({
      srcs: ['x" onerror="alert(1)'],
      ratio: 16 / 9,
      layout: 'full',
    })
    expect(full).toContain('src="x&quot; onerror=&quot;alert(1)"')
    expect(full).not.toContain('src="x" onerror=')
    const handout = buildPrintDocumentHtml({
      srcs: ["c'd", 'e&f<g>'],
      ratio: 16 / 9,
      layout: 'handout2',
    })
    expect(handout).toContain('src="c&#39;d"')
    expect(handout).toContain('src="e&amp;f&lt;g&gt;"')
  })

  it('preview mode adds page badges with the total page count', () => {
    const html = buildPrintDocumentHtml({
      srcs: srcs(5),
      ratio: 16 / 9,
      layout: 'handout2',
      preview: true,
    })
    expect(html).toContain("content: counter(pg) ' / 3';")
    expect(html).toContain('counter-reset: pg;')
  })
})

describe('normalizePrintRatio', () => {
  it('keeps real slide formats exact', () => {
    expect(normalizePrintRatio(16 / 9)).toBe(16 / 9)
    expect(normalizePrintRatio(4 / 3)).toBe(4 / 3)
    expect(normalizePrintRatio(16 / 10)).toBe(16 / 10)
  })

  it('falls back to 16:9 for non-finite and non-positive ratios', () => {
    for (const bad of [Number.NaN, Number.POSITIVE_INFINITY, 0, -1.5]) {
      expect(normalizePrintRatio(bad)).toBe(16 / 9)
      const html = buildPrintDocumentHtml({ srcs: srcs(1), ratio: bad, layout: 'full' })
      expect(html).toContain('@page { size: 13.333in 7.5in; margin: 0; }')
      expect(html).not.toContain('NaNin')
      expect(html).not.toContain('Infinityin')
    }
  })

  it('clamps absurd ratios so @page stays printable', () => {
    const tiny = buildPrintDocumentHtml({ srcs: srcs(1), ratio: 0.01, layout: 'full' })
    expect(tiny).toContain('@page { size: 1.5in 7.5in; margin: 0; }')
    const huge = buildPrintDocumentHtml({ srcs: srcs(1), ratio: 100, layout: 'full' })
    expect(huge).toContain('@page { size: 37.5in 7.5in; margin: 0; }')
  })
})

describe('multi-up full layout', () => {
  const base = { srcs: ['a', 'b', 'c', 'd', 'e'], ratio: 16 / 9, layout: 'full' as const }

  it('counts sheets by slides per sheet', () => {
    expect(printPageCount(5, 'full', 1)).toBe(5)
    expect(printPageCount(5, 'full', 4)).toBe(2)
    expect(printPageCount(17, 'full', 16)).toBe(2)
  })

  it('tiles 4 slides on a paper-sized page in Z order', () => {
    const html = buildPrintDocumentHtml({ ...base, perSheet: 4 })
    expect(html).toContain('size: 8.27in 11.69in')
    expect(html).toContain('repeat(2, 1fr)')
    expect(html).toContain('grid-row:1;grid-column:2')
    expect((html.match(/class="page nup"/g) ?? []).length).toBe(2)
  })

  it('reverse Z mirrors columns; N fills columns first', () => {
    const rz = buildPrintDocumentHtml({ ...base, perSheet: 4, direction: 'rz' })
    expect(rz.split('class="cell"')[1]).toContain('grid-row:1;grid-column:2')
    const n = buildPrintDocumentHtml({ ...base, perSheet: 4, direction: 'n' })
    expect(n.split('class="cell"')[2]).toContain('grid-row:2;grid-column:1')
  })

  it('rotates / flips and honors paper + orientation', () => {
    const html = buildPrintDocumentHtml({
      ...base,
      perSheet: 6,
      orientation: 'landscape',
      paper: 'letter',
      rotate180: true,
      flipH: true,
    })
    expect(html).toContain('size: 11in 8.5in')
    expect(html).toContain('repeat(3, 1fr)')
    expect(html).toContain('rotate(180deg) scaleX(-1)')
  })
})

describe('PowerPoint print layouts', () => {
  const base = { srcs: ['a', 'b', 'c', 'd', 'e', 'f', 'g'], ratio: 16 / 9 }

  it('counts pages for every handout size', () => {
    expect(printPageCount(7, 'handout1')).toBe(7)
    expect(printPageCount(7, 'handout4h')).toBe(2)
    expect(printPageCount(7, 'handout4v')).toBe(2)
    expect(printPageCount(7, 'handout6h')).toBe(2)
    expect(printPageCount(7, 'handout6v')).toBe(2)
    expect(printPageCount(10, 'handout9h')).toBe(2)
    expect(printPageCount(10, 'handout9v')).toBe(2)
  })

  it('horizontal 4-up fills rows first, vertical fills columns first', () => {
    const h = buildPrintDocumentHtml({ ...base, layout: 'handout4h' })
    const v = buildPrintDocumentHtml({ ...base, layout: 'handout4v' })
    // slide #2 (index 1): horizontal → row 1 / column 2; vertical → row 2 / column 1
    expect(h).toContain('style="grid-row:1;grid-column:2"')
    expect(v).toContain('style="grid-row:2;grid-column:1"')
    expect(h).toContain('grid-template-columns:repeat(2,1fr);grid-template-rows:repeat(2,1fr)')
  })

  it('9 per page is a 3×3 grid; 1 per page is a single cell', () => {
    const nine = buildPrintDocumentHtml({ ...base, layout: 'handout9h' })
    expect(nine).toContain('repeat(3,1fr);grid-template-rows:repeat(3,1fr)')
    const one = buildPrintDocumentHtml({ ...base, layout: 'handout1' })
    expect(one.match(/class="page handout h1/g)).toHaveLength(7)
  })

  it('6 slides vertical uses the 2×3 grid column-first', () => {
    const html = buildPrintDocumentHtml({ ...base, layout: 'handout6v' })
    expect(html).toContain('grid-template-columns:repeat(2,1fr);grid-template-rows:repeat(3,1fr)')
    expect(html).toContain('style="grid-row:1;grid-column:2"') // slide #4 → row 1, column 2
  })

  it('outline layout prints each slide number, title and body lines', () => {
    const html = buildPrintDocumentHtml({
      ...base,
      layout: 'outline',
      outline: [
        { number: 1, title: 'A < B', lines: ['one', 'two'] },
        { number: 3, title: 'Third', lines: [] },
      ],
    })
    expect(html).toContain('class="page flow outline"')
    expect(html).toContain('<div class="ol-title">A &lt; B</div>')
    expect(html).toContain('<div class="ol-num">3</div>')
    expect(html.match(/class="ol-line"/g)).toHaveLength(2)
    expect(html).toContain('@page flow {')
  })

  it('comment pages follow the slides and escape the text', () => {
    const html = buildPrintDocumentHtml({
      ...base,
      layout: 'full',
      comments: [{ number: 2, author: 'Ann', text: 'fix <b>this</b>\nplease' }],
      commentsTitle: 'Comments',
    })
    expect(html).toContain('class="page flow cmts"')
    expect(html).toContain('fix &lt;b&gt;this&lt;/b&gt;<br>please')
    expect(html.indexOf('class="page flow cmts"')).toBeGreaterThan(html.lastIndexOf('class="page"'))
  })

  it('a full-page slide on real paper centers on the paper and can scale to fit', () => {
    const plain = buildPrintDocumentHtml({ ...base, layout: 'full', paper: 'a4' })
    expect(plain).toContain('@page { size: 8.27in 11.69in; margin: 0; }')
    expect(plain).toContain('class="page onpaper"')
    const fit = buildPrintDocumentHtml({ ...base, layout: 'full', paper: 'a4', scaleFit: true })
    expect(fit).toContain('class="page onpaper fit"')
    const slidePaper = buildPrintDocumentHtml({ ...base, layout: 'full', paper: 'slide' })
    expect(slidePaper).toContain('@page { size: 13.333in 7.5in; margin: 0; }')
  })
})
