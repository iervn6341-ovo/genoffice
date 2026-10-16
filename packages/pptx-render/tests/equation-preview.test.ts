import { describe, expect, it } from 'vitest'
import { addElement, createBlankPptx, openPptx } from '@genoffice/pptx-engine'
import { latexToOmml } from '@genoffice/office-math'
import { buildRenderSlide } from '../src/build-slide'
import type { ShapeRenderNode } from '../src/render-tree'

const equation = (source: string) => ({
  text: source,
  rawXml: `<mc:AlternateContent><mc:Choice><a14:m><m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">${latexToOmml(source)}</m:oMath></a14:m></mc:Choice></mc:AlternateContent>`,
})

describe('structured equation render metadata', () => {
  it('exposes MathML and reversible source for a standalone native equation', async () => {
    const opened = await openPptx(await createBlankPptx())
    const slide = opened.deck.slides[0]!
    addElement(slide, {
      kind: 'textbox',
      offset: { x: 400000, y: 400000, cx: 3000000, cy: 700000 },
      paragraphs: [{ runs: [equation('\\frac{a}{b}')] }],
    })
    const node = buildRenderSlide(slide, opened.deck.size, { fitWidthPx: 960 }).nodes.at(
      -1,
    ) as ShapeRenderNode
    expect(node.hasEquation).toBe(true)
    expect(node.equationStandalone).toBe(true)
    expect(node.equationMathML).toContain('<mfrac>')
    expect(node.equationLatex).toContain('\\frac')
  })

  it('does not mark mixed text as safe for a covering overlay', async () => {
    const opened = await openPptx(await createBlankPptx())
    const slide = opened.deck.slides[0]!
    addElement(slide, {
      kind: 'textbox',
      offset: { x: 400000, y: 400000, cx: 3000000, cy: 700000 },
      paragraphs: [{ runs: [{ text: 'Before ' }, equation('x^2'), { text: ' after' }] }],
    })
    const node = buildRenderSlide(slide, opened.deck.size, { fitWidthPx: 960 }).nodes.at(
      -1,
    ) as ShapeRenderNode
    expect(node.hasEquation).toBe(true)
    expect(node.equationStandalone).toBe(false)
    const runs = node.text!.lines.flatMap((line) => line.runs)
    expect(runs.map((run) => run.text).join('')).toContain('Before ')
    expect(runs.map((run) => run.text).join('')).toContain(' after')
    const math = runs.filter((run) => run.mathML)
    expect(math).toHaveLength(1)
    expect(math[0]).toMatchObject({ srcParagraphIdx: 0, srcRunIdx: 1 })
    expect(math[0]!.mathML).toContain('<msup>')
    const prefix = runs.find((run) => run.text.includes('Before'))!
    const suffix = runs.find((run) => run.text.includes('after'))!
    expect(prefix.x + prefix.widthPx).toBeLessThanOrEqual(math[0]!.x)
    expect(math[0]!.x + math[0]!.widthPx).toBeLessThanOrEqual(suffix.x)
    expect(node.text!.lines[0]!.height).toBeGreaterThanOrEqual(
      math[0]!.mathAscentPx! + math[0]!.mathDescentPx!,
    )
  })

  it('keeps repeated run indexes distinct across paragraphs and retains rotation', async () => {
    const opened = await openPptx(await createBlankPptx())
    const slide = opened.deck.slides[0]!
    const el = addElement(slide, {
      kind: 'textbox',
      offset: { x: 400000, y: 400000, cx: 3500000, cy: 1500000 },
      paragraphs: [
        { runs: [equation('\\frac{a}{\\sqrt{b}}')] },
        { runs: [equation('x_1 + \\sum_{i=1}^{n} i')] },
      ],
    })
    el.transform.rot = 30 * 60000
    const node = buildRenderSlide(slide, opened.deck.size, { fitWidthPx: 960 }).nodes.at(
      -1,
    ) as ShapeRenderNode
    expect(node.box.rotationDeg).toBe(30)
    const math = node.text!.lines.flatMap((line) => line.runs).filter((run) => run.mathML)
    expect(math.map((run) => [run.srcParagraphIdx, run.srcRunIdx])).toEqual([
      [0, 0],
      [1, 0],
    ])
    expect(math[0]!.mathML).toContain('<msqrt>')
    expect(math[1]!.mathML).toContain('<msub>')
    expect(math[1]!.mathML).toContain('<munderover>')
  })

  it('reserves one upright equation cell in East Asian vertical text', async () => {
    const opened = await openPptx(await createBlankPptx())
    const slide = opened.deck.slides[0]!
    const el = addElement(slide, {
      kind: 'textbox',
      offset: { x: 400000, y: 400000, cx: 1600000, cy: 2500000 },
      paragraphs: [{ runs: [{ text: '上' }, equation('\\frac{a}{b}'), { text: '下' }] }],
    })
    el.text!.vert = 'eaVert'
    const node = buildRenderSlide(slide, opened.deck.size, { fitWidthPx: 960 }).nodes.at(
      -1,
    ) as ShapeRenderNode
    const math = node.text!.lines.flatMap((line) => line.runs).filter((run) => run.mathML)
    expect(math).toHaveLength(1)
    expect(math[0]!.rotate90).toBeUndefined()
    expect(math[0]!.mathAscentPx! + math[0]!.mathDescentPx!).toBeGreaterThan(math[0]!.fontSizePx)
  })

  it('keeps an equation atomic between right-to-left text runs', async () => {
    const opened = await openPptx(await createBlankPptx())
    const slide = opened.deck.slides[0]!
    addElement(slide, {
      kind: 'textbox',
      offset: { x: 400000, y: 400000, cx: 4000000, cy: 700000 },
      paragraphs: [
        { rtl: true, runs: [{ text: 'שלום ' }, equation('\\frac{a}{b}'), { text: ' עולם' }] },
      ],
    })
    const node = buildRenderSlide(slide, opened.deck.size, { fitWidthPx: 960 }).nodes.at(
      -1,
    ) as ShapeRenderNode
    const runs = node.text!.lines.flatMap((line) => line.runs)
    const math = runs.filter((run) => run.mathML)
    expect(math).toHaveLength(1)
    expect(math[0]!.text).toBe('\\frac{a}{b}')
    expect(math[0]).toMatchObject({ srcParagraphIdx: 0, srcRunIdx: 1 })
  })

  it('reserves the tallest math branch, not the total count of parallel fractions', async () => {
    const opened = await openPptx(await createBlankPptx())
    const slide = opened.deck.slides[0]!
    for (const source of ['\\frac{a}{b}', '\\frac{a}{b}+\\frac{c}{d}', '\\frac{a}{\\frac{b}{c}}']) {
      addElement(slide, {
        kind: 'textbox',
        offset: { x: 400000, y: 400000, cx: 5000000, cy: 1000000 },
        paragraphs: [{ runs: [equation(source)] }],
      })
    }
    const heights = buildRenderSlide(slide, opened.deck.size, { fitWidthPx: 960 })
      .nodes.slice(-3)
      .map((renderNode) => {
        const math = (renderNode as ShapeRenderNode)
          .text!.lines.flatMap((line) => line.runs)
          .find((run) => run.mathML)!
        return math.mathAscentPx! + math.mathDescentPx!
      })
    expect(heights[1]).toBe(heights[0])
    expect(heights[2]).toBeGreaterThan(heights[0]!)
  })

  it('reserves each row of an imported equation array instead of shrinking it into one line', async () => {
    const opened = await openPptx(await createBlankPptx())
    const slide = opened.deck.slides[0]!
    const row = (text: string) => `<m:e><m:r><m:t>${text}</m:t></m:r></m:e>`
    const wrap = (inner: string) =>
      `<mc:AlternateContent><mc:Choice><a14:m><m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">${inner}</m:oMath></a14:m></mc:Choice></mc:AlternateContent>`
    for (const rows of [[row('a')], [row('a'), row('b'), row('c')]]) {
      addElement(slide, {
        kind: 'textbox',
        offset: { x: 400000, y: 400000, cx: 5000000, cy: 1500000 },
        paragraphs: [
          {
            runs: [
              {
                text: 'imported equation array',
                rawXml: wrap(`<m:eqArr>${rows.join('')}</m:eqArr>`),
              },
            ],
          },
        ],
      })
    }
    const math = buildRenderSlide(slide, opened.deck.size, { fitWidthPx: 960 })
      .nodes.slice(-2)
      .map((renderNode) =>
        (renderNode as ShapeRenderNode)
          .text!.lines.flatMap((line) => line.runs)
          .find((run) => run.mathML),
      )
    expect(math[0]!.mathML).toContain('<mtable>')
    expect(math[1]!.mathML!.match(/<mtr>/g)).toHaveLength(3)
    expect(math[1]!.mathAscentPx! + math[1]!.mathDescentPx!).toBeGreaterThan(
      (math[0]!.mathAscentPx! + math[0]!.mathDescentPx!) * 2,
    )
  })

  it('sizes stacked fractions by their widest branch, but adds parallel fractions', async () => {
    const opened = await openPptx(await createBlankPptx())
    const slide = opened.deck.slides[0]!
    const sources = [
      '\\frac{a}{b}',
      '\\frac{a}{b}+\\frac{c}{d}',
      '\\frac{\\frac{a}{b}}{\\frac{c}{d}}',
    ]
    for (const source of sources) {
      addElement(slide, {
        kind: 'textbox',
        offset: { x: 400000, y: 400000, cx: 5000000, cy: 1500000 },
        paragraphs: [{ runs: [equation(source)] }],
      })
    }
    addElement(slide, {
      kind: 'textbox',
      offset: { x: 400000, y: 400000, cx: 5000000, cy: 1500000 },
      paragraphs: [{ runs: [{ ...equation(sources[0]!), text: 'long linear fallback (a)/(b)' }] }],
    })
    const widths = buildRenderSlide(slide, opened.deck.size, { fitWidthPx: 960 })
      .nodes.slice(-4)
      .map(
        (renderNode) =>
          (renderNode as ShapeRenderNode)
            .text!.lines.flatMap((line) => line.runs)
            .find((run) => run.mathML)!.widthPx,
      )
    expect(widths[1]).toBeGreaterThan(widths[0]! * 1.7)
    expect(widths[2]).toBeLessThan(widths[1]!)
    expect(widths[3]).toBeCloseTo(widths[0]!, 6)
  })
})
