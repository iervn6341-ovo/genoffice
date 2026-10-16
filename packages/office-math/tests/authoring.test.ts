import { describe, expect, it } from 'vitest'
import {
  EQUATION_STRUCTURES,
  SYMBOL_SUBSETS,
  findSymbolByCodePoint,
  latexToOmml,
  ommlToLatex,
  ommlToMathML,
  symbolCodePoint,
  symbolSearchText,
} from '../src/index'

const outer = (inner: string) => `<m:oMath>${inner}</m:oMath>`
const structuresOf = (xml: string) =>
  xml
    .match(
      /<m:(?:f|sSup|sSub|sSubSup|rad|nary|d|m|acc|bar|groupChr|borderBox|limLow|limUpp|func)(?=[\s>])/g,
    )
    ?.sort() ?? []

describe('authoring catalog survives native math roundtrips', () => {
  for (const category of EQUATION_STRUCTURES) {
    for (const template of category.templates) {
      it(`${category.id}/${template.id} is editable after writing OMML`, () => {
        const latex = template.latex.replace('{{selection}}', 'x')
        const xml = outer(latexToOmml(latex))
        expect(ommlToMathML(xml)).toContain('<math')
        const reopened = ommlToLatex(xml)
        expect(reopened).not.toBeNull()
        const rewritten = outer(latexToOmml(reopened!))
        expect(structuresOf(rewritten)).toEqual(structuresOf(xml))
      })
    }
  }

  it('preserves nested structures through a second edit', () => {
    const first = outer(latexToOmml('\\boxed{\\frac{\\sqrt{x}}{1+x_{i}^{2}}}'))
    const reopened = ommlToLatex(first)!
    const edited = outer(latexToOmml(`${reopened} + \\overset{def}{y}`))
    expect(ommlToMathML(edited)).toContain('<mfrac')
    expect(ommlToMathML(edited)).toContain('<msqrt')
    expect(ommlToMathML(edited)).toContain('<mover')
    expect(ommlToLatex(edited)).toContain('\\boxed')
  })

  it('preserves native named functions when importing and editing', () => {
    const xml =
      '<m:oMath><m:func><m:fName><m:r><m:rPr><m:sty m:val="p"/></m:rPr><m:t>sin</m:t></m:r></m:fName><m:e><m:r><m:t>x</m:t></m:r></m:e></m:func></m:oMath>'
    const source = ommlToLatex(xml)!
    expect(latexToOmml(source)).toContain('<m:func>')
    expect(ommlToMathML(outer(latexToOmml(source)))).toContain('sin')
  })

  it('does not normalize imported italic function names or limits to upright text', () => {
    for (const properties of ['', '<m:rPr><m:sty m:val="i"/></m:rPr>']) {
      const name = `<m:r>${properties}<m:t>sin</m:t></m:r>`
      expect(
        ommlToLatex(
          outer(`<m:func><m:fName>${name}</m:fName><m:e><m:r><m:t>x</m:t></m:r></m:e></m:func>`),
        ),
      ).toBeNull()
      const limit = `<m:limLow><m:e><m:r>${properties}<m:t>lim</m:t></m:r></m:e><m:lim><m:r><m:t>x</m:t></m:r></m:lim></m:limLow>`
      const source = ommlToLatex(outer(limit))!
      expect(source).toContain('\\underset')
      expect(latexToOmml(source)).not.toContain('<m:sty m:val="p"/>')
    }
  })

  it('does not flatten structural function names or ignored child elements', () => {
    const run = '<m:r><m:t>x</m:t></m:r>'
    expect(
      ommlToLatex(
        outer(
          `<m:func><m:fName><m:f><m:num>${run}</m:num><m:den>${run}</m:den></m:f></m:fName><m:e>${run}</m:e></m:func>`,
        ),
      ),
    ).toBeNull()
    expect(ommlToLatex(outer(`<m:r><m:t>x</m:t><m:unsupportedNativeContent/></m:r>`))).toBeNull()
    expect(
      ommlToLatex(outer(`<m:m><m:mr><m:e>${run}</m:e><m:unsupportedNativeContent/></m:mr></m:m>`)),
    ).toBeNull()
  })

  it('rejects malformed XML before reverse conversion', () => {
    expect(ommlToLatex('<m:oMath><m:r><m:t>x</m:r></m:oMath>')).toBeNull()
    expect(ommlToMathML('<m:oMath><m:r><m:t>x</m:r></m:oMath>')).toBe('')
  })

  it('rejects unsupported structures rather than losing them in an edit', () => {
    for (const tag of ['box', 'phant', 'unknown']) {
      expect(
        ommlToLatex(outer(`<m:${tag}><m:e><m:r><m:t>x</m:t></m:r></m:e></m:${tag}>`)),
      ).toBeNull()
    }
    expect(
      ommlToLatex(
        outer(
          '<m:borderBox><m:borderBoxPr><m:hideTop m:val="1"/></m:borderBoxPr><m:e><m:r><m:t>x</m:t></m:r></m:e></m:borderBox>',
        ),
      ),
    ).toBeNull()
    expect(() => latexToOmml('\\unknown{x}')).toThrow('Unsupported command')
  })

  it('rejects OMML properties that the editable LaTeX subset cannot preserve', () => {
    const cases = [
      [
        'formatted script',
        outer(
          '<m:sSup><m:sSupPr><m:ctrlPr><w:rPr><w:b/></w:rPr></m:ctrlPr></m:sSupPr><m:e><m:r><m:t>x</m:t></m:r></m:e><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup>',
        ),
      ],
      [
        'small fraction',
        outer(
          '<m:f><m:fPr><m:smallFrac/></m:fPr><m:num><m:r><m:t>a</m:t></m:r></m:num><m:den><m:r><m:t>b</m:t></m:r></m:den></m:f>',
        ),
      ],
      [
        'hidden limit with content',
        outer(
          '<m:nary><m:naryPr><m:chr m:val="∑"/><m:subHide m:val="1"/></m:naryPr><m:sub><m:r><m:t>i</m:t></m:r></m:sub><m:e><m:r><m:t>x</m:t></m:r></m:e></m:nary>',
        ),
      ],
      [
        'hidden radical degree with content',
        outer(
          '<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg><m:r><m:t>3</m:t></m:r></m:deg><m:e><m:r><m:t>x</m:t></m:r></m:e></m:rad>',
        ),
      ],
      [
        'non-growing delimiter',
        outer(
          '<m:d><m:dPr><m:begChr m:val="("/><m:endChr m:val=")"/><m:grow m:val="0"/></m:dPr><m:e><m:r><m:t>x</m:t></m:r></m:e></m:d>',
        ),
      ],
      [
        'matrix alignment',
        outer(
          '<m:m><m:mPr><m:mcs><m:mc><m:mcPr><m:mcJc m:val="left"/></m:mcPr></m:mc></m:mcs></m:mPr><m:mr><m:e><m:r><m:t>x</m:t></m:r></m:e></m:mr></m:m>',
        ),
      ],
      [
        'group character position mismatch',
        outer(
          '<m:groupChr><m:groupChrPr><m:chr m:val="⏟"/><m:pos m:val="top"/></m:groupChrPr><m:e><m:r><m:t>x</m:t></m:r></m:e></m:groupChr>',
        ),
      ],
    ] as const

    for (const [label, xml] of cases) expect(ommlToLatex(xml), label).toBeNull()
  })
})

describe('symbol authoring data', () => {
  it('finds names and exact Unicode codepoints across subsets', () => {
    expect(SYMBOL_SUBSETS.flatMap((s) => [...s.symbols])).toContain('∫')
    expect(symbolSearchText('∫')).toContain('INTEGRAL')
    expect(symbolCodePoint('≤')).toBe('2264')
    expect(findSymbolByCodePoint('U+2264')).toBe('≤')
    expect(findSymbolByCodePoint('0x1D6FC')).toBe('𝛼')
  })
  it('does not insert invalid scalars or XML control characters', () => {
    for (const invalid of ['D800', '110000', '0000', '001F', 'FFFE', 'xyz']) {
      expect(findSymbolByCodePoint(invalid)).toBeNull()
    }
  })
})
