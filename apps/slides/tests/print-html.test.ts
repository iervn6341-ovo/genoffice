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
  it('omits disabled native notes slide/body regions while retaining master fields', () => {
    const html = buildPrintDocumentHtml({
      srcs: srcs(1),
      ratio: 16 / 9,
      layout: 'notes',
      notes: ['Kept in PPTX'],
      notesFields: [
        { slideImageVisible: false, notesBodyVisible: false, header: { text: 'Visible header' } },
      ],
    })
    const page = new DOMParser().parseFromString(html, 'text/html').querySelector('.page.notes')!
    expect(page.querySelector('.notes-slide-image')).toBeNull()
    expect(page.querySelector('.note')).toBeNull()
    expect(page.querySelector('.hf-header')?.textContent).toBe('Visible header')
    expect(page.textContent).not.toContain('Kept in PPTX')
  })
  it('renders native notes paragraph bullet, RTL, exact line spacing, margins, tab grid and run styles', () => {
    const html = buildPrintDocumentHtml({
      srcs: srcs(1),
      ratio: 16 / 9,
      layout: 'notes',
      notesFields: [
        {
          notesParagraphs: [
            {
              align: 'right',
              rtl: true,
              level: 1,
              lineExactPt: 18,
              spaceBeforePt: 6,
              spaceAfterPct: 50,
              marginLeftPt: 72,
              marginRightPt: 10,
              indentPt: -18,
              tabStopsPt: [36],
              defaultTabPt: 36,
              bullet: { type: 'char', char: '★', color: '#AA3300', fontFamily: 'Arial' },
              runs: [{ text: 'אב\tNotes', fontSizePt: 18, strike: true, highlight: '#FFFF00' }],
            },
          ],
        },
      ],
    })
    const page = new DOMParser().parseFromString(html, 'text/html')
    const paragraph = page.querySelector('.note p')!
    expect(paragraph.getAttribute('dir')).toBe('rtl')
    expect(paragraph.getAttribute('style')).toContain('line-height:18pt')
    expect(paragraph.getAttribute('style')).toContain('margin-top:6pt')
    expect(paragraph.getAttribute('style')).toContain('margin-left:72pt')
    expect(paragraph.getAttribute('style')).toContain('text-indent:-18pt')
    expect(paragraph.getAttribute('style')).toContain('tab-size:4')
    expect(paragraph.getAttribute('data-tab-stops')).toBe('36')
    expect(paragraph.querySelector('.note-bullet')?.textContent).toBe('★')
    expect(paragraph.querySelector('.note-bullet')?.getAttribute('style')).toContain(
      'color:#AA3300',
    )
    expect(
      paragraph.querySelector('.note-tab-segment > span:not(.note-bullet)')?.getAttribute('style'),
    ).toContain('line-through')
    expect(paragraph.querySelectorAll('.note-tab-segment')).toHaveLength(2)
    expect(paragraph.textContent).toContain('אבNotes')
  })

  it('positions multiple native tabs with right, center and decimal alignment', () => {
    const html = buildPrintDocumentHtml({
      srcs: srcs(1),
      ratio: 16 / 9,
      layout: 'notes',
      notesFields: [
        {
          notesParagraphs: [
            {
              tabStopsPt: [120, 240, 360],
              tabStops: [
                { posPt: 120, align: 'r' },
                { posPt: 240, align: 'ctr' },
                { posPt: 360, align: 'dec' },
              ],
              runs: [{ text: 'Label\tRight\tCenter\t12.34', fontSizePt: 12 }],
            },
          ],
        },
      ],
    })
    const page = new DOMParser().parseFromString(html, 'text/html')
    const segments = page.querySelectorAll('.note-tab-segment')
    expect(segments).toHaveLength(5)
    expect(segments[1]?.getAttribute('style')).toContain('left:120pt;transform:translateX(-100%)')
    expect(segments[2]?.getAttribute('style')).toContain('left:240pt;transform:translateX(-50%)')
    expect(segments[3]?.getAttribute('style')).toContain('left:360pt;transform:translateX(-100%)')
    expect(segments[3]?.textContent).toContain('12')
    expect(segments[4]?.getAttribute('style')).toContain('left:360pt')
    expect(segments[4]?.textContent).toContain('.34')
  })
  it('skips native stops covered by preceding text and continues on the default grid', () => {
    const html = buildPrintDocumentHtml({
      srcs: srcs(1),
      ratio: 16 / 9,
      layout: 'notes',
      notesFields: [
        {
          notesParagraphs: [
            {
              defaultTabPt: 72,
              tabStops: [
                { posPt: 72, align: 'l' },
                { posPt: 144, align: 'r' },
                { posPt: 216, align: 'ctr' },
                { posPt: 288, align: 'dec' },
              ],
              runs: [
                { text: 'TAB_ORIGIN\tTAB_LEFT\tTAB_RIGHT\tTAB_CENTER\t12.34', fontSizePt: 12 },
              ],
            },
            {
              rtl: true,
              defaultTabPt: 72,
              tabStops: [
                { posPt: 72, align: 'l' },
                { posPt: 108, align: 'ctr' },
              ],
              runs: [{ text: 'RTL_ORIGIN\tRTL_LEFT\tRTL_RIGHT', fontSizePt: 12 }],
            },
          ],
        },
      ],
    })
    const page = new DOMParser().parseFromString(html, 'text/html')
    const [ltr, rtl] = [...page.querySelectorAll('.note p')]
    const ltrSegments = [...ltr!.querySelectorAll('.note-tab-segment')]
    const rtlSegments = [...rtl!.querySelectorAll('.note-tab-segment')]
    expect(ltrSegments[1]?.getAttribute('style')).toContain('left:144pt')
    expect(ltrSegments[2]?.getAttribute('style')).toContain('left:216pt')
    expect(ltrSegments[3]?.getAttribute('style')).toContain('left:288pt')
    expect(ltrSegments[4]?.getAttribute('style')).toContain('left:432pt')
    expect(rtlSegments[1]?.getAttribute('style')).toContain('right:108pt')
    expect(rtlSegments[2]?.getAttribute('style')).toContain('right:144pt')
  })
  it('draws the resolved native picture bullet in notes print output', () => {
    const html = buildPrintDocumentHtml({
      srcs: srcs(1),
      ratio: 16 / 9,
      layout: 'notes',
      notesFields: [
        {
          notesParagraphs: [
            {
              bullet: { type: 'blip', imageDataUrl: 'data:image/png;base64,AAAA' },
              runs: [{ text: 'Bullet body' }],
            },
          ],
        },
      ],
    })
    const page = new DOMParser().parseFromString(html, 'text/html')
    expect(page.querySelector('.note-bullet img')?.getAttribute('src')).toBe(
      'data:image/png;base64,AAAA',
    )
    expect(page.querySelector('.note')?.textContent).toContain('Bullet body')
  })
  it('prints large authored native field and run sizes without an arbitrary 72pt cap', () => {
    const html = buildPrintDocumentHtml({
      srcs: srcs(1),
      ratio: 16 / 9,
      layout: 'notes',
      notesFields: [
        {
          header: { text: 'Large header', fontSizePt: 100 },
          notesParagraphs: [{ runs: [{ text: 'Large body', fontSizePt: 100 }] }],
        },
      ],
    })
    const page = new DOMParser().parseFromString(html, 'text/html')
    expect(page.querySelector('.hf-header span')?.getAttribute('style')).toContain(
      'font-size:100pt',
    )
    expect(page.querySelector('.note span')?.getAttribute('style')).toContain('font-size:100pt')
  })
  it('measures native stops from the frame after paragraph margins and mirrors RTL alignment', () => {
    const html = buildPrintDocumentHtml({
      srcs: srcs(1),
      ratio: 16 / 9,
      layout: 'notes',
      notesFields: [
        {
          notesParagraphs: [
            { marginLeftPt: 12, tabStops: [{ posPt: 36, align: 'r' }], runs: [{ text: 'A\tB' }] },
            {
              rtl: true,
              marginRightPt: 18,
              tabStops: [
                { posPt: 72, align: 'l' },
                { posPt: 108, align: 'ctr' },
              ],
              runs: [{ text: 'א\tב\tג' }],
            },
          ],
        },
      ],
    })
    const page = new DOMParser().parseFromString(html, 'text/html')
    const paragraphs = page.querySelectorAll('.note p')
    expect(
      paragraphs[0]?.querySelectorAll('.note-tab-segment')[1]?.getAttribute('style'),
    ).toContain('left:24pt;transform:translateX(-100%)')
    expect(
      paragraphs[1]?.querySelectorAll('.note-tab-segment')[1]?.getAttribute('style'),
    ).toContain('right:54pt;transform:translateX(100%)')
    expect(
      paragraphs[1]?.querySelectorAll('.note-tab-segment')[2]?.getAttribute('style'),
    ).toContain('right:90pt;transform:translateX(50%)')
  })
  it('layers native master artwork behind rotated fields and styled notes text in preview/PDF HTML', () => {
    const masterPath = 'ppt/notesMasters/notesMaster1.xml'
    const html = buildPrintDocumentHtml({
      srcs: srcs(1),
      ratio: 16 / 9,
      layout: 'notes',
      masterArtwork: { [masterPath]: 'data:image/png;base64,ARTWORK' },
      notesFields: [
        {
          masterPath,
          canvas: { cx: 6858000, cy: 9144000 },
          date: {
            text: 'native',
            box: { x: 457200, y: 8595360, cx: 1828800, cy: 320040, rotationDeg: 90, flipH: true },
          },
          notesBody: { x: 685800, y: 4400550, cx: 5486400, cy: 3600450, flipV: true },
          notesParagraphs: [
            {
              align: 'right',
              runs: [
                { text: 'Styled body', fontSizePt: 18, color: '#225588', fontFamily: 'Arial' },
              ],
            },
          ],
        },
      ],
    })
    const page = new DOMParser().parseFromString(html, 'text/html').querySelector('.page.notes')!
    expect(page.querySelector('.master-artwork')?.getAttribute('src')).toBe(
      'data:image/png;base64,ARTWORK',
    )
    expect(page.querySelector('.hf-date')?.getAttribute('style')).toContain(
      'rotate(90deg) scaleX(-1)',
    )
    expect(page.querySelector('.note')?.getAttribute('style')).toContain('scaleY(-1)')
    expect(page.querySelector('.note p')?.getAttribute('style')).toContain('text-align:right')
    expect(page.querySelector('.note span')?.getAttribute('style')).toContain('font-size:18pt')
    expect(page.querySelector('.note span')?.getAttribute('style')).toContain('color:#225588')
    expect(page.textContent).toContain('Styled body')
    expect(html).toContain('.page .master-artwork { position:absolute;')
  })
  it('maps native notes-master placeholder boxes to preview and PDF page coordinates', () => {
    const canvas = { cx: 6858000, cy: 9144000 }
    const html = buildPrintDocumentHtml({
      srcs: srcs(1),
      ratio: 16 / 9,
      layout: 'notes',
      notes: ['Native body'],
      notesFields: [
        {
          canvas,
          background: '#EEEEDD',
          header: {
            text: 'Positioned',
            box: { x: 685800, y: 914400, cx: 2743200, cy: 457200 },
            align: 'right',
            fontFamily: 'Arial',
          },
          slideImage: { x: 457200, y: 1371600, cx: 5943600, cy: 3200400 },
          notesBody: { x: 457200, y: 4800600, cx: 5943600, cy: 3657600 },
        },
      ],
    })
    const page = new DOMParser().parseFromString(html, 'text/html').querySelector('.page.notes')!
    expect(page.getAttribute('style')).toContain('background:#EEEEDD')
    expect(page.querySelector('.hf-header')?.getAttribute('style')).toContain('left:10%')
    expect(page.querySelector('.hf-header')?.getAttribute('style')).toContain('top:10%')
    expect(page.querySelector('.hf-header')?.getAttribute('style')).toContain('text-align:right')
    expect(page.querySelector('img')?.getAttribute('style')).toContain('top:15%')
    expect(page.querySelector('.note')?.getAttribute('style')).toContain('top:52.5%')
  })
  it('prints native notes overrides, dynamic dates and offset page fields without changing slide selection', () => {
    const html = buildPrintDocumentHtml({
      srcs: srcs(2),
      ratio: 16 / 9,
      layout: 'notes',
      notes: ['First', 'Second'],
      slideIndexes: [1, 4],
      firstSlideNum: 0,
      dateLocale: 'en-US',
      notesFields: [
        {
          header: { text: 'Local <header>', bold: true },
          date: { text: 'old cache', field: 'datetime1' },
          pageNum: { text: 'old number', field: 'slidenum' },
        },
        { footer: { text: 'Master & footer', italic: true } },
      ],
    })
    expect(html).toContain('Local &lt;header&gt;')
    expect(html).toContain('font-weight:700')
    expect(html).toContain('font-style:italic')
    expect(html).toContain('Master &amp; footer')
    expect(html).toContain('<div class="hf-field hf-pageNum"')
    expect(html).toContain('<span>1</span></div>')
    expect(html).not.toContain('old cache')
    expect(html).not.toContain('old number')
  })

  it('keeps notes-page numbers tied to slide positions and the presentation start', () => {
    const html = buildPrintDocumentHtml({
      srcs: srcs(5),
      ratio: 16 / 9,
      layout: 'notes',
      firstSlideNum: 0,
      slideIndexes: [0, 1, 2, 3, 4],
      notesFields: Array.from({ length: 5 }, () => ({
        pageNum: { text: 'old cache', field: 'slidenum' },
      })),
    })
    const pages = new DOMParser().parseFromString(html, 'text/html').querySelectorAll('.page.notes')
    expect([...pages].map((page) => page.querySelector('.hf-pageNum')?.textContent)).toEqual([
      '0',
      '1',
      '2',
      '3',
      '4',
    ])
  })

  it('lets an edited native footer extend beyond its placeholder box in print', () => {
    const html = buildPrintDocumentHtml({
      srcs: srcs(1),
      ratio: 16 / 9,
      layout: 'notes',
      notesFields: [
        {
          canvas: { cx: 6858000, cy: 9144000 },
          footer: {
            text: 'HANDOUT_FOOTER_NATIVE_2026_EDITED',
            box: { x: 2286000, y: 8595360, cx: 2286000, cy: 320040 },
          },
        },
      ],
    })
    const page = new DOMParser().parseFromString(html, 'text/html')
    const footer = page.querySelector('.hf-footer')!
    expect(footer.textContent).toBe('HANDOUT_FOOTER_NATIVE_2026_EDITED')
    expect(footer.getAttribute('style')).toContain('width:33.33333333333333%')
    expect(html).toContain(
      '.hf-field { position: absolute; z-index: 3; font-size: 10pt; overflow: visible;',
    )
  })

  it('prints handout page fields from 1 across sheets, independent of slide-number start', () => {
    const html = buildPrintDocumentHtml({
      srcs: srcs(5),
      ratio: 16 / 9,
      layout: 'handout2',
      firstSlideNum: 0,
      handoutFields: {
        header: { text: 'Handout heading' },
        pageNum: { text: 'cached', field: 'slidenum' },
      },
    })
    const pages = new DOMParser()
      .parseFromString(html, 'text/html')
      .querySelectorAll('.page.handout')
    expect(pages).toHaveLength(3)
    expect([...pages].map((page) => page.querySelector('.hf-pageNum')?.textContent)).toEqual([
      '1',
      '2',
      '3',
    ])
    expect(html.match(/Handout heading/g)).toHaveLength(3)
    for (const [layout, per] of [
      ['handout1', 1],
      ['handout2', 2],
      ['handout3', 3],
    ] as const) {
      const variant = buildPrintDocumentHtml({
        srcs: srcs(per * 2 + 1),
        ratio: 16 / 9,
        layout,
        firstSlideNum: -3,
        handoutFields: { pageNum: { text: 'old cache', field: 'slidenum' } },
      })
      const sheets = new DOMParser()
        .parseFromString(variant, 'text/html')
        .querySelectorAll('.page.handout')
      expect([...sheets].map((page) => page.querySelector('.hf-pageNum')?.textContent)).toEqual([
        '1',
        '2',
        '3',
      ])
    }
  })

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

  it('keeps native handout fields above a separate cell region on complete and partial sheets', () => {
    for (const layout of [
      'handout1',
      'handout2',
      'handout3',
      'handout4h',
      'handout4v',
      'handout6h',
      'handout6v',
      'handout9h',
      'handout9v',
    ] as const) {
      const html = buildPrintDocumentHtml({
        ...base,
        layout,
        handoutFields: {
          header: { text: 'Visible header' },
          footer: { text: 'Visible footer' },
          pageNum: { text: 'cached', field: 'slidenum' },
        },
      })
      const doc = new DOMParser().parseFromString(html, 'text/html')
      const pages = [...doc.querySelectorAll('.page.handout')]
      expect(pages).toHaveLength(printPageCount(base.srcs.length, layout))
      for (const page of pages) {
        const content = page.querySelector(':scope > .handout-content')
        expect(content).not.toBeNull()
        expect(content!.querySelectorAll('.cell').length).toBeGreaterThan(0)
        expect(content!.querySelector('.hf-field')).toBeNull()
        expect(page.querySelector(':scope > .hf-header')?.textContent).toBe('Visible header')
        expect(page.querySelector(':scope > .hf-footer')?.textContent).toContain('Visible footer')
      }
      expect(html).toContain('.page.handout { position: relative; padding: 0.42in 0.4in; }')
      expect(html).toContain('.hf-field { position: absolute; z-index: 3;')
    }
  })

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
