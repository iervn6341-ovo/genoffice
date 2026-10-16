import { describe, expect, it } from 'vitest'
import { formatHeaderFooterDate } from '../src/renderer/header-footer-date'

describe('PowerPoint date field examples', () => {
  const date = new Date(2026, 9, 13, 16, 28, 34)
  it('matches the 13 documented field patterns for an English locale', () => {
    expect(
      Array.from({ length: 13 }, (_, i) =>
        formatHeaderFooterDate(date, `datetime${i + 1}`, 'en-US'),
      ),
    ).toEqual([
      '10/13/2026',
      'Tuesday, October 13, 2026',
      '13 October 2026',
      'October 13, 2026',
      '13-Oct-26',
      'October 26',
      'Oct-26',
      '10/13/2026 4:28 PM',
      '10/13/2026 4:28:34 PM',
      '16:28',
      '16:28:34',
      '4:28 PM',
      '4:28:34 PM',
    ])
  })
  it('keeps an unknown imported type unformatted', () => {
    expect(formatHeaderFooterDate(date, 'datetime42', 'en-US')).toBeNull()
  })
})
