import { describe, expect, it } from 'vitest'
import { transformCase } from '../src/renderer/editor/case-transform'

describe('transformCase', () => {
  it('tOGGLE cASE swaps every letter and leaves the rest alone', () => {
    expect(transformCase('Hello World 123!', 'toggle')).toBe('hELLO wORLD 123!')
    expect(transformCase('已有 MiXed', 'toggle')).toBe('已有 mIxED')
  })
  it('toggling twice returns the original', () => {
    const s = 'Quarterly Report, Q3 2026'
    expect(transformCase(transformCase(s, 'toggle'), 'toggle')).toBe(s)
  })
})
