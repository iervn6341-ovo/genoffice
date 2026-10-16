import { describe, expect, it } from 'vitest'
import {
  addElement,
  createBlankPptx,
  openPptx,
  savePptx,
  type TextElement,
} from '@genoffice/pptx-engine'
import { EQUATION_STRUCTURES, latexToOmml, ommlToLatex } from '@genoffice/office-math'
import { runTxn } from '../src/ops/executor'
import '../src/ops/index'
import { equationRun } from '../src/ops/equation-ops'

const box = { x: 800000, y: 900000, cx: 3600000, cy: 700000 }

describe('editable native slide equations', () => {
  it('inserts, saves, reopens and replaces OMML without changing the shape box or id', async () => {
    const opened = await openPptx(await createBlankPptx())
    const insert = runTxn(opened, {
      ops: [{ op: 'insertEquation', target: { slide: 0 }, latex: '\\frac{a}{b}', box }],
    })
    expect(insert.applied).toBe(true)
    const original = opened.deck.slides[0]!.elements.at(-1) as TextElement
    const sourceId = original.id
    expect(original.text!.paragraphs[0]!.runs[0]!.rawXml).toContain('<a14:m>')
    const reopened = await openPptx(await savePptx(opened))
    const math = reopened.deck.slides[0]!.elements.find(
      (e) =>
        e.type === 'text' &&
        (e as TextElement).text?.paragraphs.some((p) =>
          p.runs.some((r) => r.rawXml?.includes('<a14:m')),
        ),
    ) as TextElement
    expect(math).toBeTruthy()
    const originalBox = structuredClone(math.transform)
    const replacement = runTxn(reopened, {
      ops: [{ op: 'replaceEquation', target: { slide: 0, el: math.id }, latex: '\\sqrt{x}' }],
    })
    expect(replacement.applied).toBe(true)
    expect(math.id).toBeTruthy()
    expect(math.transform).toEqual(originalBox)
    const rereopened = await openPptx(await savePptx(reopened))
    const after = rereopened.deck.slides[0]!.elements.find(
      (e) =>
        e.type === 'text' &&
        (e as TextElement).text?.paragraphs.some((p) =>
          p.runs.some((r) => r.rawXml?.includes('<a14:m')),
        ),
    ) as TextElement
    expect(after.text!.paragraphs[0]!.runs[0]!.rawXml).toContain('<m:rad>')
    expect(after.transform).toEqual(originalBox)
    expect(sourceId).toBeTruthy()
  })

  it('keeps all structure gallery formulas reversible', () => {
    for (const category of EQUATION_STRUCTURES) {
      for (const template of category.templates) {
        const source = template.latex.replace('{{selection}}', 'x')
        const omml = `<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">${latexToOmml(source)}</m:oMath>`
        expect(ommlToLatex(omml), `${category.id}/${template.id}`).not.toBeNull()
        const run = equationRun(source)
        expect(run.rawXml, `${category.id}/${template.id}`).toContain('<a14:m>')
        expect(run.text, `${category.id}/${template.id}`).not.toBe('')
      }
    }
  })

  it('replaces only the math run in mixed rich text and retains unrelated elements', async () => {
    const opened = await openPptx(await createBlankPptx())
    const slide = opened.deck.slides[0]!
    const el = addElement(slide, {
      kind: 'textbox',
      offset: box,
      paragraphs: [
        {
          runs: [
            { text: 'Before ', bold: true, color: '#B02020' },
            equationRun('x^2'),
            { text: ' after', italic: true, fontFamily: 'Aptos' },
          ],
        },
      ],
    })
    const neighbor = addElement(slide, {
      kind: 'textbox',
      offset: { ...box, y: 1800000 },
      paragraphs: [{ runs: [{ text: 'Untouched' }] }],
    })
    const originalNeighbor = neighbor.anchor.originalXml
    const loaded = await openPptx(await savePptx(opened))
    const target = loaded.deck.slides[0]!.elements.find((item) =>
      item.anchor.originalXml.includes('Before '),
    ) as TextElement
    const neighborAfter = loaded.deck.slides[0]!.elements.find((item) =>
      item.anchor.originalXml.includes('Untouched'),
    )!
    expect(target).toBeTruthy()
    const r = runTxn(loaded, {
      ops: [{ op: 'replaceEquation', target: { slide: 0, el: target.id }, latex: '\\frac{p}{q}' }],
    })
    expect(r.applied).toBe(true)
    const after = await openPptx(await savePptx(loaded))
    const xml = after.archive.readText(after.deck.slides[0]!.path)!
    expect(xml).toContain('<m:f>')
    expect(xml).toContain('Before ')
    expect(xml).toContain(' after')
    expect(
      (
        after.deck.slides[0]!.elements.find((item) =>
          item.anchor.originalXml.includes('Untouched'),
        ) as TextElement
      ).anchor.originalXml,
    ).toBe(neighborAfter.anchor.originalXml)
    expect(originalNeighbor).toContain('Untouched')
    expect(el.id).toBeTruthy()
  })

  it('rejects unsupported imported OMML without changing its raw XML', async () => {
    const opened = await openPptx(await createBlankPptx())
    runTxn(opened, { ops: [{ op: 'insertEquation', target: { slide: 0 }, latex: 'x', box }] })
    const math = opened.deck.slides[0]!.elements.at(-1) as TextElement
    const raw = math.text!.paragraphs[0]!.runs[0]!.rawXml!
    math.text!.paragraphs[0]!.runs[0]!.rawXml = raw
      .replace('<m:r>', '<m:phant><m:r>')
      .replace('</m:r>', '</m:r></m:phant>')
    expect(ommlToLatex(math.text!.paragraphs[0]!.runs[0]!.rawXml!)).toBeNull()
    const result = runTxn(opened, {
      ops: [{ op: 'replaceEquation', target: { slide: 0, el: math.id }, latex: 'y' }],
    })
    expect(result.applied).toBe(false)
    expect(math.text!.paragraphs[0]!.runs[0]!.rawXml).toContain('<m:phant>')
  })
})
