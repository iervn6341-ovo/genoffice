/**
 * Word sizes a header/footer line by the fonts its text uses (Word probe
 * 2026-10-01): a page-number-only footer line under an Arial/標楷體 Normal
 * takes Arial's height, and a Footer-styled paragraph lays out its unsized
 * runs at the style's 10pt. Each strip paragraph therefore carries its own
 * per-script --doc-line-factor and the style size.
 */
import { describe, expect, it } from 'vitest'
import { makeGapHfEl } from '../src/renderer/editor/hf-dom'

const paras = (els: HTMLElement) => Array.from(els.querySelectorAll<HTMLElement>('.page-hf-para'))

describe('header/footer strip paragraph line factor', () => {
  it('picks the factor by each paragraph text and applies the style size', () => {
    const el = makeGapHfEl({
      kind: 'footer',
      value: {
        text: '',
        paras: [
          { styleSizeHalfPoints: 20, runs: [{ text: '2' }] },
          { runs: [{ text: '第 2 頁' }] },
          { runs: [] },
        ],
      },
      images: [],
      pageNo: 2,
      pageTotal: 9,
    })
    const [num, cjk, blank] = paras(el)
    expect(num.style.getPropertyValue('--doc-line-factor')).toBe('var(--doc-line-factor-latin,1.2)')
    expect(num.style.fontSize).toBe('10pt')
    expect(cjk.style.getPropertyValue('--doc-line-factor')).toBe('var(--doc-line-factor-cjk,1.7)')
    expect(cjk.style.fontSize).toBe('')
    expect(blank.style.getPropertyValue('--doc-line-factor')).toBe(
      'var(--doc-line-factor-latin,1.2)',
    )
  })
})
