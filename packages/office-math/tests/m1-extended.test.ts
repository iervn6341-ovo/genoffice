import { describe, expect, it } from 'vitest'
import { latexToOmml, ommlToLatex, ommlToMathML } from '../src/index'

const outer = (inner: string) => `<m:oMath>${inner}</m:oMath>`
const W = '<w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/></w:rPr>'
const WI = '<w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/><w:i/></w:rPr>'
const ctrl = `<m:ctrlPr>${WI}</m:ctrlPr>`
const wr = (text: string) => `<m:r>${W}<m:t>${text}</m:t></m:r>`

/** LaTeX -> OMML -> LaTeX -> OMML must settle on the same OMML (the structure survives an edit). */
function settles(latex: string): { omml: string; reopened: string } {
  const omml = latexToOmml(latex)
  const reopened = ommlToLatex(outer(omml))
  expect(reopened, `${latex} should reopen as source`).not.toBeNull()
  expect(latexToOmml(reopened!), `${latex} -> ${reopened}`).toBe(omml)
  expect(ommlToMathML(outer(omml)), latex).toContain('<math')
  return { omml, reopened: reopened! }
}

describe('M1 structures write native OMML and reopen as source', () => {
  const cases: Array<[string, string, RegExp]> = [
    ['stacked fraction', '\\genfrac{}{}{0pt}{}{n}{k}', /<m:type m:val="noBar"\/>/],
    ['binomial with brackets', '\\genfrac{[}{]}{0pt}{}{n}{k}', /<m:begChr m:val="\["\/>/],
    ['skewed fraction', '\\nicefrac{a}{b}', /<m:type m:val="skw"\/>/],
    ['linear fraction', '\\linfrac{a}{b}', /<m:type m:val="lin"\/>/],
    ['left subscript and superscript', '{}_{1}^{2}{x}', /<m:sPre>/],
    ['left subscript only', '{}_{n}{C}', /<m:sPre>/],
    [
      'limits above and below an integral',
      '\\int\\limits_{0}^{1} {x}',
      /<m:limLoc m:val="undOvr"\/>/,
    ],
    ['side limits on a sum', '\\sum\\nolimits_{i=1}^{n} {x}', /<m:limLoc m:val="subSup"\/>/],
    ['surface integral', '\\oiint_{S} {F}', /m:val="∯"/],
    ['volume integral', '\\oiiint_{V} {F}', /m:val="∰"/],
    ['clockwise contour integral', '\\varointclockwise_{C} {F}', /m:val="∲"/],
    ['logical or operator', '\\bigvee_{i=1}^{n} {p}', /m:val="⋁"/],
    ['circled plus operator', '\\bigoplus_{i} {V}', /m:val="⨁"/],
    [
      'bracket with separator',
      '\\left\\langle a \\middle| b \\right\\rangle',
      /<m:sepChr m:val="\|"\/>/,
    ],
    ['three-part set', '\\left\\{ x \\middle| y \\middle| z \\right\\}', /<m:e>.*<m:e>.*<m:e>/s],
    ['double bracket', '\\left\\llbracket x \\right\\rrbracket', /m:val="⟦"/],
    ['piecewise', '\\begin{cases} x & x>0 \\\\ -x & x<0 \\end{cases}', /<m:eqArr>.*<m:aln\/>/s],
    ['aligned equations', '\\begin{aligned} a &= b \\\\ c &= d \\end{aligned}', /<m:eqArr>/],
    ['inverse trig', '\\sin^{-1}{x}', /<m:fName><m:sSup>/],
    ['logarithm base', '\\log_{2}{x}', /<m:fName><m:sSub>/],
    ['limit function', '\\lim_{n\\to\\infty}{a}', /<m:func><m:fName><m:limLow>/],
    ['maximum over a set', '\\max_{x\\in A}{f}', /<m:fName><m:limLow>/],
    ['operatorname scripts', '\\operatorname{arsinh}^{2}{x}', /<m:fName><m:sSup>/],
    ['wide hat', '\\widehat{xy}', /m:val="̂"/],
    ['acute accent', '\\acute{e}', /m:val="́"/],
    ['triple dot', '\\dddot{x}', /m:val="⃛"/],
    ['over parenthesis', '\\overparen{abc}', /m:val="⏜"/],
    ['under bracket', '\\underbracket{abc}', /m:val="⎵"/],
    ['extensible arrow', '\\xrightarrow{yields}', /<m:limUpp>/],
    ['phantom', '\\phantom{x}', /<m:show m:val="0"\/>/],
    ['cancel', '\\cancel{x}', /<m:strikeTLBR m:val="1"\/>/],
    ['back cancel', '\\bcancel{x}', /<m:strikeBLTR m:val="1"\/>/],
    ['cross out', '\\xcancel{x}', /<m:strikeBLTR m:val="1"\/><m:strikeTLBR m:val="1"\/>/],
    ['strike through', '\\hcancel{x}', /<m:strikeH m:val="1"\/>/],
    ['bold math', '\\mathbf{v}', /<m:sty m:val="b"\/>/],
    ['bold italic math', '\\boldsymbol{v}', /<m:sty m:val="bi"\/>/],
    ['script letter', '\\mathcal{L}', /<m:scr m:val="script"\/>/],
    ['fraktur letter', '\\mathfrak{g}', /<m:scr m:val="fraktur"\/>/],
    ['double-struck letter', '\\mathbb{R}', /<m:scr m:val="double-struck"\/>/],
    ['sans-serif letter', '\\mathsf{A}', /<m:scr m:val="sans-serif"\/>/],
    ['monospace letter', '\\mathtt{A}', /<m:scr m:val="monospace"\/>/],
  ]
  for (const [label, latex, pattern] of cases) {
    it(`${label}: ${latex}`, () => {
      const { omml } = settles(latex)
      expect(omml).toMatch(pattern)
    })
  }

  it('keeps a command that merely starts like \\right inside a delimiter group', () => {
    const { omml } = settles('\\left( a \\rightarrow b \\right)')
    expect(omml).toContain('→')
  })

  it('rejects mixed separators and unsupported slots instead of guessing', () => {
    expect(() => latexToOmml('\\left( a \\middle| b \\middle/ c \\right)')).toThrow()
  })
})

describe('M1 imported Word equations stay re-editable', () => {
  it('opens a typical Word quadratic formula (Cambria Math run fonts and control characters)', () => {
    const xml = outer(
      `${wr('x=')}<m:f><m:fPr>${ctrl}</m:fPr><m:num>${wr('−b±')}<m:rad><m:radPr>${ctrl}</m:radPr><m:deg/>` +
        `<m:e>${wr('b')}<m:sSup><m:sSupPr>${ctrl}</m:sSupPr><m:e>${wr('b')}</m:e><m:sup>${wr('2')}</m:sup></m:sSup>` +
        `${wr('−4ac')}</m:e></m:rad></m:num><m:den>${wr('2a')}</m:den></m:f>`,
    )
    const source = ommlToLatex(xml)
    expect(source).toBe('x=\\frac{−b\\pm \\sqrt{b{b}^{2}−4ac}}{2a}')
    expect(latexToOmml(source!)).toContain('<m:rad>')
  })

  it('opens Word n-ary, function and delimiter structures that carry ctrlPr', () => {
    const sum = outer(
      `<m:nary><m:naryPr><m:chr m:val="∑"/><m:ctrlPr>${WI}</m:ctrlPr></m:naryPr><m:sub>${wr('i=1')}</m:sub><m:sup>${wr('n')}</m:sup><m:e>${wr('i')}</m:e></m:nary>`,
    )
    expect(ommlToLatex(sum)).toBe('\\sum_{i=1}^{n} {i}')
    const sin = outer(
      `<m:func><m:funcPr>${ctrl}</m:funcPr><m:fName><m:sSup><m:sSupPr>${ctrl}</m:sSupPr><m:e><m:r><m:rPr><m:sty m:val="p"/></m:rPr>${W}<m:t>sin</m:t></m:r></m:e><m:sup>${wr('−1')}</m:sup></m:sSup></m:fName><m:e>${wr('x')}</m:e></m:func>`,
    )
    expect(ommlToLatex(sin)).toBe('\\sin^{−1} {x}')
    const set = outer(
      `<m:d><m:dPr><m:begChr m:val="{"/><m:sepChr m:val="|"/><m:endChr m:val="}"/>${ctrl}</m:dPr><m:e>${wr('x')}</m:e><m:e>${wr('x&gt;0')}</m:e></m:d>`,
    )
    expect(ommlToLatex(set)).toBe('\\left\\{ x \\middle| x>0 \\right\\}')
  })

  it('keeps native Word column specifications on matrices editable only when nothing is lost', () => {
    const centered =
      '<m:m><m:mPr><m:mcs><m:mc><m:mcPr><m:count m:val="2"/><m:mcJc m:val="center"/></m:mcPr></m:mc></m:mcs></m:mPr>' +
      '<m:mr><m:e><m:r><m:t>a</m:t></m:r></m:e><m:e><m:r><m:t>b</m:t></m:r></m:e></m:mr></m:m>'
    expect(ommlToLatex(outer(centered))).toBe('\\begin{matrix} a & b \\end{matrix}')
    expect(ommlToLatex(outer(centered.replace('center', 'left')))).toBeNull()
  })

  it('refuses Word run formatting that regeneration would drop', () => {
    for (const properties of [
      '<w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/></w:rPr>',
      '<w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/><w:color w:val="FF0000"/></w:rPr>',
      '<w:rPr><w:b/></w:rPr>',
    ]) {
      expect(ommlToLatex(outer(`<m:r>${properties}<m:t>x</m:t></m:r>`)), properties).toBeNull()
    }
    // tracked changes on a control character are Office data too
    expect(
      ommlToLatex(
        outer(
          `<m:sSup><m:sSupPr><m:ctrlPr><w:ins w:id="1" w:author="a"/></m:ctrlPr></m:sSupPr><m:e>${wr('x')}</m:e><m:sup>${wr('2')}</m:sup></m:sSup>`,
        ),
      ),
    ).toBeNull()
  })

  it('refuses structures whose properties the source cannot express', () => {
    const fixedBracket =
      '<m:d><m:dPr><m:grow m:val="0"/></m:dPr><m:e><m:r><m:t>x</m:t></m:r></m:e></m:d>'
    expect(ommlToLatex(outer(fixedBracket))).toBeNull()
    const topArrow =
      '<m:groupChr><m:groupChrPr><m:chr m:val="⏞"/><m:pos m:val="top"/><m:vertJc m:val="top"/></m:groupChrPr><m:e><m:r><m:t>x</m:t></m:r></m:e></m:groupChr>'
    expect(ommlToLatex(outer(topArrow))).toBeNull()
    const partialBox =
      '<m:borderBox><m:borderBoxPr><m:hideTop m:val="1"/><m:strikeH m:val="1"/></m:borderBoxPr><m:e><m:r><m:t>x</m:t></m:r></m:e></m:borderBox>'
    expect(ommlToLatex(outer(partialBox))).toBeNull()
  })

  it('does not accept an alignment point outside an equation array', () => {
    expect(ommlToLatex(outer('<m:r><m:rPr><m:aln/></m:rPr><m:t>x</m:t></m:r>'))).toBeNull()
    const array =
      '<m:eqArr><m:e><m:r><m:t>a</m:t></m:r><m:r><m:rPr><m:aln/></m:rPr><m:t>=b</m:t></m:r></m:e></m:eqArr>'
    expect(ommlToLatex(outer(array))).toBe('\\begin{aligned} a& =b \\end{aligned}')
  })
})

describe('gallery size matches the Word for Mac equation tab', () => {
  it('has the same template count per category (230 in total)', async () => {
    const { EQUATION_STRUCTURES } = await import('../src/index')
    const counts = Object.fromEntries(EQUATION_STRUCTURES.map((c) => [c.id, c.templates.length]))
    expect(counts).toEqual({
      fraction: 9,
      script: 8,
      radical: 6,
      integral: 21,
      largeOperator: 40,
      delimiter: 41,
      function: 27,
      accent: 27,
      limitLog: 8,
      operator: 21,
      matrix: 22,
    })
    const ids = EQUATION_STRUCTURES.flatMap((c) => c.templates.map((t) => t.id))
    expect(new Set(ids).size).toBe(230)
  })
})
