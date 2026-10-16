import { describe, expect, it } from 'vitest'
import { latexToOmml } from '@genoffice/office-math'
import { HeuristicMetrics } from '../src/metrics'
import { makeViewport } from '../src/coords'
import { layoutText } from '../src/text-layout'
import type { Paragraph, TextBody, TextRun } from '@genoffice/pptx-engine'

const vp = makeViewport({ cx: 9_525_000, cy: 9_525_000 }, 1000)

function mathRun(latex: string, fallback = latex): TextRun {
  return {
    text: fallback,
    fontSize: 18,
    fontFamily: 'Arial',
    color: '#111111',
    rawXml: `<mc:AlternateContent><mc:Choice><a14:m><m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">${latexToOmml(latex)}</m:oMath></a14:m></mc:Choice><mc:Fallback><a:r><a:t>${fallback}</a:t></a:r></mc:Fallback></mc:AlternateContent>`,
  }
}

function layout(paragraphs: Paragraph[]) {
  const body: TextBody = {
    paragraphs,
    anchor: 'top',
    insets: { l: 0, t: 0, r: 0, b: 0 },
    wrap: true,
  }
  return layoutText({
    body,
    boxWidthPx: 1000,
    boxHeightPx: 600,
    metrics: new HeuristicMetrics(),
    vp,
  })
}

describe('native math inline layout review regressions', () => {
  it('measures parallel fractions by their maximum vertical structure, not their count', () => {
    const one = layout([{ runs: [mathRun('\\frac{a}{b}', 'a/b')] }])
    const parallelRun = mathRun('\\frac{a}{b}+\\frac{c}{d}', 'a/b+c/d')
    expect(parallelRun.rawXml?.match(/<m:f(?=[\s>])/g)).toHaveLength(2)
    const parallel = layout([{ runs: [parallelRun] }])
    expect(parallel.lines[0]!.height).toBeCloseTo(one.lines[0]!.height, 0)
  })

  it('keeps an RTL formula atomic while retaining the neighboring paragraph direction', () => {
    const rtlFormula = layout([
      {
        rtl: true,
        runs: [
          { text: 'قبل ', fontSize: 18, fontFamily: 'Arial' },
          mathRun('\\text{אבג}+x', 'אבג+x'),
          { text: ' بعد', fontSize: 18, fontFamily: 'Arial' },
        ],
      },
    ])
    const mathRuns = rtlFormula.lines.flatMap((line) => line.runs).filter((run) => run.mathML)
    expect(mathRuns).toHaveLength(1)
    expect(mathRuns[0]!.srcParagraphIdx).toBe(0)
  })

  it('identifies equation runs by paragraph and run index across multiple paragraphs', () => {
    const multiParagraph = layout([
      { runs: [{ text: 'A ', fontSize: 18 }, mathRun('x^2', 'x^2')] },
      { runs: [{ text: 'B ', fontSize: 18 }, mathRun('y^2', 'y^2')] },
    ])
    const mathRuns = multiParagraph.lines.flatMap((line) => line.runs).filter((run) => run.mathML)
    expect(mathRuns.map((run) => [run.srcParagraphIdx, run.srcRunIdx])).toEqual([
      [0, 1],
      [1, 1],
    ])
  })
})
