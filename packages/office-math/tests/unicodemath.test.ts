import { describe, expect, it } from 'vitest'
import {
  EQUATION_STRUCTURES,
  latexToOmml,
  ommlToLatex,
  ommlToUnicodeMath,
  unicodeMathToOmml,
} from '../src/index'

const outer = (inner: string) => `<m:oMath>${inner}</m:oMath>`
const linear = (latex: string) => ommlToUnicodeMath(outer(latexToOmml(latex)))

describe('UnicodeMath serializer', () => {
  const cases: Array<[string, string]> = [
    ['\\frac{a}{b}', 'a/b'],
    ['\\frac{a+b}{c}', '(a+b)/c'],
    ['{x}^{2}', 'x^2'],
    ['{x}_{i}^{n}', 'x_i^n'],
    ['{x}^{2n}', 'x^(2n)'],
    ['\\sqrt{x}', '√(x)'],
    ['\\sqrt[3]{x}', '∛(x)'],
    ['\\sqrt[n]{x}', '√(n&x)'],
    ['\\sum_{i=1}^{n} {a}', '∑_(i=1)^n▒a'],
    ['\\int_{0}^{1} {f}', '∫_0^1▒f'],
    ['\\left( a+b \\right)', '(a+b)'],
    ['\\left[ a \\middle| b \\right]', '[a|b]'],
    ['\\sin {x}', 'sin⁡x'],
    ['\\hat{x}', 'x̂'],
    ['\\overline{x}', 'x̅'],
    ['\\begin{matrix} a & b \\\\ c & d \\end{matrix}', '■(a&b@c&d)'],
    ['\\begin{cases} x & y \\\\ z & w \\end{cases}', '{█(x&y@z&w)┤'],
    ['\\overset{def}{=}', '=┴(def)'],
    ['\\boxed{x}', '▭(x)'],
    ['\\nicefrac{a}{b}', 'a⊘b'],
    ['\\genfrac{}{}{0pt}{}{n}{k}', 'n¦k'],
  ]
  for (const [latex, expected] of cases) {
    it(`${latex} -> ${expected}`, () => {
      expect(linear(latex)).toBe(expected)
    })
  }

  it('refuses structures it cannot express', () => {
    expect(linear('\\mathbf{v}')).toBeNull()
    expect(linear('\\int\\limits_{0}^{1} {x}')).toBeNull()
    expect(linear('{}_{a}^{b}{x}')).toBeNull()
  })
})

describe('UnicodeMath parser', () => {
  it('reads fractions with grouped operands without adding brackets', () => {
    const omml = unicodeMathToOmml('(a+b)/(c+d)')
    expect(omml).toContain('<m:f>')
    expect(omml).not.toContain('<m:d>')
    expect(ommlToLatex(outer(omml))).toBe('\\frac{a+b}{c+d}')
  })

  it('keeps a bracket when no operator follows the group', () => {
    expect(ommlToLatex(outer(unicodeMathToOmml('(a+b)')))).toBe('\\left( a+b \\right)')
    expect(ommlToLatex(outer(unicodeMathToOmml('((a+b))/c')))).toBe(
      '\\frac{\\left( a+b \\right)}{c}',
    )
  })

  it('reads scripts, radicals, n-ary operators, functions and matrices', () => {
    expect(ommlToLatex(outer(unicodeMathToOmml('x_i^2')))).toBe('{x}_{i}^{2}')
    expect(ommlToLatex(outer(unicodeMathToOmml('√(n&x)')))).toBe('\\sqrt[n]{x}')
    expect(ommlToLatex(outer(unicodeMathToOmml('∑_(i=1)^n▒a_i')))).toBe('\\sum_{i=1}^{n} {{a}_{i}}')
    expect(ommlToLatex(outer(unicodeMathToOmml('sin⁡x')))).toBe('\\sin {x}')
    expect(ommlToLatex(outer(unicodeMathToOmml('■(1&0@0&1)')))).toBe(
      '\\begin{matrix} 1 & 0 \\\\ 0 & 1 \\end{matrix}',
    )
    expect(ommlToLatex(outer(unicodeMathToOmml('{█(x&y@z&w)┤')))).toBe(
      '\\begin{cases} x& y \\\\ z& w \\end{cases}',
    )
  })

  it('reads quoted text, accents and over/under limits', () => {
    expect(unicodeMathToOmml('"if" x')).toContain('<m:nor/>')
    expect(unicodeMathToOmml('x̂')).toContain('<m:acc>')
    expect(unicodeMathToOmml('a┬b')).toContain('<m:limLow>')
    expect(unicodeMathToOmml('a┴b')).toContain('<m:limUpp>')
  })

  it('reports malformed input', () => {
    expect(() => unicodeMathToOmml('(a+b')).toThrow()
    expect(() => unicodeMathToOmml('■(a&b')).toThrow()
    expect(() => unicodeMathToOmml('a/')).toThrow()
  })
})

describe('Professional <-> Linear round trip over the gallery', () => {
  it('re-reads every linear-capable template to the same structure', () => {
    let covered = 0
    for (const category of EQUATION_STRUCTURES) {
      for (const template of category.templates) {
        const omml = outer(latexToOmml(template.latex.replace('{{selection}}', 'x')))
        const text = ommlToUnicodeMath(omml)
        // literal bracket characters inside a run read back as brackets (linear form has no escape)
        if (text === null || /<m:t[^>]*>[^<]*[()[\]{}][^<]*<\/m:t>/.test(omml)) continue
        covered++
        const back = unicodeMathToOmml(text)
        const normalize = (xml: string) => ommlToLatex(outer(xml)) ?? xml
        const strip = (v: string) => v.replace(/\s+/g, '')
        expect(strip(normalize(back)), `${template.id}: ${text}`).toBe(
          strip(normalize(omml.slice(9, -10))),
        )
      }
    }
    expect(covered).toBeGreaterThan(120)
  })
})
