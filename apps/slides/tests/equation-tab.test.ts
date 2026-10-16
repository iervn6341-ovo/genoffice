import { describe, expect, it } from 'vitest'
import { EQUATION_STRUCTURES, latexToOmml } from '@genoffice/office-math'
import { contextTabsForElement } from '../src/renderer/components/context-tabs'
import { appendToEquationSource } from '../src/renderer/components/RibbonEquationTab'

describe('equation contextual tab', () => {
  it('adds the Equation tab for a selected native equation, whatever the element type', () => {
    expect(contextTabsForElement(null, true)).toEqual(['equationTools'])
    expect(contextTabsForElement('textShape', true)).toEqual(['shapeFormat', 'equationTools'])
    expect(contextTabsForElement('textShape')).toEqual(['shapeFormat'])
  })

  it('appends structures to the source and keeps every catalog template convertible', () => {
    expect(appendToEquationSource('a+', '\\frac{{{selection}}}{y}')).toBe('a+\\frac{x}{y}')
    expect(appendToEquationSource('', '\\sqrt{{{selection}}}')).toBe('\\sqrt{x}')
    expect(appendToEquationSource('E=', 'π')).toBe('E=π')
    expect(EQUATION_STRUCTURES).toHaveLength(11)
    for (const cat of EQUATION_STRUCTURES) {
      for (const template of cat.templates) {
        const source = appendToEquationSource('a+', template.latex)
        expect(() => latexToOmml(source)).not.toThrow()
      }
    }
  })
})
