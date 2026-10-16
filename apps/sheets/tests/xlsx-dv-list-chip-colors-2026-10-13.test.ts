import { describe, expect, it } from 'vitest'

import { applyDvRules } from '@genoffice/xlsx-gateway/gateway/xlsx-dv'

const emptyWorksheet = '<worksheet><sheetData/></worksheet>'
const range = { startRow: 1, endRow: 3, startColumn: 1, endColumn: 1 }

describe('UI dropdown chip-color metadata export probe (2026-10-13)', () => {
  it.each([
    ['default color join', ','],
    ['JSON color array with an empty slot', JSON.stringify(['#f8696b', ''])],
  ])('serializes only formula1 for list validation with %s', (_caseName, formula2) => {
    const xml = applyDvRules(emptyWorksheet, [
      {
        ranges: [range],
        rule: {
          type: 'list',
          formula1: JSON.stringify(['  Open  ', 'Closed ']),
          formula2,
        },
      },
    ])

    expect(xml).toContain('sqref="B2:B4"')
    expect(xml).toContain('<formula1>"  Open  ,Closed "</formula1>')
    expect(xml).not.toContain('<formula2>')
  })
})
