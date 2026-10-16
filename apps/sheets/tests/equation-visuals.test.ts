import { createHash } from 'node:crypto'
import { describe, expect, it } from 'vitest'
import { latexToOmml } from '@genoffice/office-math'
import { equationXml } from '@genoffice/xlsx-gateway/gateway/xlsx-drawing-add'
import { equationAtDrawingIndex } from '../src/main/equation-visuals'

function drawing(latex: string): string {
  const omml = latexToOmml(latex)
  return `<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><xdr:twoCellAnchor><xdr:from/><xdr:to/>${equationXml(2, { latex, omml })}<xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>`
}

describe('editable Office equation rehydration', () => {
  it('recovers exact author source when the native OMML is unchanged', () => {
    const latex = 'x = \\frac{1}{2}'
    const found = equationAtDrawingIndex(drawing(latex), 0)
    expect(found?.latex).toBe(latex)
    expect(found?.omml).toContain('<m:f>')
  })

  it('derives new source from Office-edited OMML instead of stale metadata', () => {
    const original = drawing('x + 1')
    const changed = original.replace(latexToOmml('x + 1'), latexToOmml('x + 2'))
    expect(equationAtDrawingIndex(changed, 0)?.latex).toContain('2')
    expect(equationAtDrawingIndex(changed, 0)?.latex).not.toBe('x + 1')
  })

  it('does not classify an ordinary drawing or a missing anchor as an equation', () => {
    expect(equationAtDrawingIndex('<xdr:wsDr/>', 0)).toBeNull()
    expect(equationAtDrawingIndex(drawing('x'), 1)).toBeNull()
  })

  it('recognizes a direct a14:m shape from an external Office producer', () => {
    const direct = `<xdr:wsDr><xdr:twoCellAnchor><xdr:sp><xdr:txBody><a:p><a14:m><m:oMath>${latexToOmml('x+1')}</m:oMath></a14:m></a:p></xdr:txBody></xdr:sp></xdr:twoCellAnchor></xdr:wsDr>`
    expect(equationAtDrawingIndex(direct, 0)?.latex).toContain('x')
  })

  it('derives source from native OMML when metadata source disagrees with its digest-bound math', () => {
    const native = drawing('x')
    const tampered = native.replace(
      `GenOfficeEquationV2:${Buffer.from('x').toString('base64')}:`,
      `GenOfficeEquationV2:${Buffer.from('y').toString('base64')}:`,
    )
    const found = equationAtDrawingIndex(tampered, 0)
    expect(found?.latex).toBe('x')
    expect(found?.omml).toBe(latexToOmml('x'))
  })

  it('classifies a native equation with a picture fallback as read-only before editing', () => {
    const pictorial = drawing('x').replace(
      /<mc:Fallback>[\s\S]*?<\/mc:Fallback>/,
      '<mc:Fallback><xdr:pic><a:blip r:embed="rId7"/></xdr:pic></mc:Fallback>',
    )
    expect(equationAtDrawingIndex(pictorial, 0)).toMatchObject({
      latex: null,
      omml: latexToOmml('x'),
    })
  })

  it('classifies an ambiguous text fallback as read-only before editing', () => {
    const ambiguous = drawing('x').replace(
      '<a:r><a:t>x</a:t></a:r>',
      '<a:r><a:t>x</a:t></a:r><a:r><a:t>x</a:t></a:r>',
    )
    expect(equationAtDrawingIndex(ambiguous, 0)).toMatchObject({ latex: null })
  })

  it('does not treat a substring inside an unrelated fallback run as the equation span', () => {
    const ambiguous = drawing('x+1').replace(
      '<a:r><a:t>x+1</a:t></a:r>',
      '<a:r><a:t>x+10</a:t></a:r>',
    )
    expect(equationAtDrawingIndex(ambiguous, 0)).toMatchObject({ latex: null })
  })

  it('allows a custom-described native equation when its text fallback has one matched math span', () => {
    const imported = drawing('q')
      .replaceAll(/descr="GenOfficeEquationV2:[^"]*"/g, 'descr="Analyst annotation"')
      .replace(
        '<a:r><a:t>q</a:t></a:r>',
        '<a:r><a:t>before </a:t></a:r><a:r><a:t>q</a:t></a:r><a:r><a:t> after</a:t></a:r>',
      )
    expect(equationAtDrawingIndex(imported, 0)).toMatchObject({
      latex: 'q',
      omml: latexToOmml('q'),
    })
  })

  it('does not offer stale source when Office adds an unsupported native structure', () => {
    const changed = drawing('x').replace(latexToOmml('x'), '<m:unsupportedNativeMath/>')
    expect(equationAtDrawingIndex(changed, 0)).toMatchObject({ latex: null })
  })

  it('does not trust stored source for unsupported math even if the digest was updated', () => {
    const original = latexToOmml('x')
    const unsupported = '<m:unsupportedNativeMath/>'
    const changed = drawing('x')
      .replace(original, unsupported)
      .replaceAll(
        createHash('sha256').update(original).digest('hex'),
        createHash('sha256').update(unsupported).digest('hex'),
      )
    expect(equationAtDrawingIndex(changed, 0)).toMatchObject({ latex: null, omml: unsupported })
  })

  it('keeps a drawing with two native math zones read-only', () => {
    const omml = latexToOmml('x')
    const math = `<a14:m><m:oMath>${omml}</m:oMath></a14:m>`
    const doubled = drawing('x').replace(math, math + math)
    expect(equationAtDrawingIndex(doubled, 0)).toMatchObject({ latex: null, omml })
  })
})
