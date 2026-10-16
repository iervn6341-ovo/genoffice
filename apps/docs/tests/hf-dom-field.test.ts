import { describe, expect, it } from 'vitest'
import { PAGE_MARK, type HeaderFooter } from '@genoffice/docx-engine'
import { makeGapHfEl } from '../src/renderer/editor/hf-dom'

describe('paginated header field cache', () => {
  it('renders mixed result formatting without displaying the protected editor token', () => {
    const value: HeaderFooter = {
      text: 'October 13',
      paras: [
        {
          runs: [
            {
              text: 'October 13',
              hfField: {
                xml: '<w:fldSimple w:instr="DATE"/>',
                instruction: 'DATE',
                token: '⟦DATE·1⟧',
                displayText: 'October 13',
                resultRuns: [
                  { text: 'October', bold: true },
                  { text: ' 13', italic: true },
                ],
              },
            },
          ],
        },
      ],
    }
    const el = makeGapHfEl({ kind: 'header', value, pageNo: 2, pageTotal: 4 })
    const spans = [...el.querySelectorAll('.page-hf-para > span')]
    expect(spans.map((span) => span.textContent)).toEqual(['October', ' 13'])
    expect((spans[0] as HTMLElement).style.fontWeight).toBe('600')
    expect((spans[1] as HTMLElement).style.fontStyle).toBe('italic')
    expect(el.textContent).not.toContain('⟦DATE')
  })

  it('still renders a native page number instead of its stale cached value', () => {
    const value: HeaderFooter = {
      text: PAGE_MARK,
      pageNumber: true,
      paras: [
        {
          runs: [
            {
              text: PAGE_MARK,
              hfField: {
                xml: '<w:fldSimple w:instr="PAGE"/>',
                instruction: 'PAGE',
                token: '⟦PAGE·1⟧',
                displayText: PAGE_MARK,
                resultRuns: [{ text: '1' }],
              },
            },
          ],
        },
      ],
    }
    const el = makeGapHfEl({ kind: 'footer', value, pageNo: 7, pageTotal: 9 })
    expect(el.textContent).toBe('7')
  })
})
