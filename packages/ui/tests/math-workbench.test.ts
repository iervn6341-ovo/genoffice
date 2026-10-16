import { describe, expect, it } from 'vitest'
import { EQUATION_STRUCTURES, SYMBOL_SUBSETS, ommlToLatex } from '@genoffice/office-math'
import { insertMathTemplate, mathPreview } from '../src/MathWorkbench'
import { latexToOmml } from '@genoffice/office-math'

describe('math structure workbench', () => {
  it('shows every catalog template as supported math and round-trips editable source', () => {
    expect(EQUATION_STRUCTURES).toHaveLength(11)
    const templates = EQUATION_STRUCTURES.flatMap((group) => group.templates)
    expect(templates).toHaveLength(230)
    for (const template of templates) {
      const inserted = insertMathTemplate('', 0, 0, template.latex)
      const preview = mathPreview(inserted.latex)
      expect(preview, template.id).toHaveProperty('mathml')
      const reversed = ommlToLatex(`<m:oMath>${latexToOmml(inserted.latex)}</m:oMath>`)
      expect(reversed, template.id).toBeTruthy()
      expect(mathPreview(reversed!), template.id).toHaveProperty('mathml')
    }
  })

  it('wraps the selected source and selects an empty placeholder for editing', () => {
    const template = EQUATION_STRUCTURES[0].templates[0].latex
    expect(insertMathTemplate('a + bc', 4, 6, template).latex).toBe('a + \\frac{bc}{y}')
    const empty = insertMathTemplate('a + ', 4, 4, template)
    expect(empty.latex.slice(empty.selectionStart, empty.selectionEnd)).toBe('x')
    expect(mathPreview(empty.latex)).toHaveProperty('mathml')
  })

  it('inserts symbols from the equation gallery at the source selection', () => {
    for (const subset of SYMBOL_SUBSETS.filter(
      (item) => item.id === 'math' || item.id === 'greek' || item.id === 'arrows',
    )) {
      for (const char of subset.symbols) {
        const source = insertMathTemplate('a + b', 4, 5, char)
        expect(source.latex, `${subset.id}: ${char}`).toBe(`a + ${char}`)
        expect(mathPreview(source.latex), `${subset.id}: ${char}`).toHaveProperty('mathml')
      }
    }
  })
})

describe('linear (UnicodeMath) conversion helpers', () => {
  it('converts both ways and reports errors', async () => {
    const { latexToLinear, linearToLatex } = await import('../src/MathWorkbench')
    expect(latexToLinear('\\frac{a+b}{c}')).toBe('(a+b)/c')
    expect(latexToLinear('\\mathbf{v}')).toBeNull()
    expect(linearToLatex('(a+b)/c')).toEqual({ latex: '\\frac{a+b}{c}' })
    expect('error' in linearToLatex('(a+b')).toBe(true)
  })
})
