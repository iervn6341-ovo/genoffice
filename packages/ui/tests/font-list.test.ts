import { describe, expect, it } from 'vitest'
import {
  BUILTIN_FONT_FAMILIES,
  fontFamiliesFor,
  LATIN_FONT_FAMILIES,
  TRADITIONAL_CHINESE_FONT_FAMILIES,
  SIMPLIFIED_CHINESE_FONT_FAMILIES,
  JAPANESE_FONT_FAMILIES,
  KOREAN_FONT_FAMILIES,
  sortFontFamilies,
} from '../src/font-list'

describe('font-list catalog and sorting', () => {
  it('replenishes candidate fonts across Latin and CJK scripts', () => {
    expect(LATIN_FONT_FAMILIES.length).toBeGreaterThanOrEqual(30)
    expect(TRADITIONAL_CHINESE_FONT_FAMILIES.length).toBeGreaterThanOrEqual(15)
    expect(SIMPLIFIED_CHINESE_FONT_FAMILIES.length).toBeGreaterThanOrEqual(20)
    expect(JAPANESE_FONT_FAMILIES.length).toBeGreaterThanOrEqual(10)
    expect(KOREAN_FONT_FAMILIES.length).toBeGreaterThanOrEqual(10)
    expect(BUILTIN_FONT_FAMILIES.length).toBeGreaterThanOrEqual(100)
  })

  it('contains no duplicate font names (case-insensitive)', () => {
    const seen = new Set<string>()
    for (const font of BUILTIN_FONT_FAMILIES) {
      const lower = font.toLowerCase()
      expect(seen.has(lower), `Duplicate font name found: ${font}`).toBe(false)
      seen.add(lower)
    }
  })

  it('includes essential Traditional Chinese office fonts in BUILTIN catalog', () => {
    expect(BUILTIN_FONT_FAMILIES).toContain('微軟正黑體')
    expect(BUILTIN_FONT_FAMILIES).toContain('新細明體')
    expect(BUILTIN_FONT_FAMILIES).toContain('細明體')
    expect(BUILTIN_FONT_FAMILIES).toContain('標楷體')
    expect(BUILTIN_FONT_FAMILIES).toContain('DFKai-SB')
    expect(BUILTIN_FONT_FAMILIES).toContain('PMingLiU')
    expect(BUILTIN_FONT_FAMILIES).toContain('Microsoft JhengHei')
  })

  it('includes essential Latin office and modern Microsoft fonts', () => {
    expect(BUILTIN_FONT_FAMILIES).toContain('Aptos')
    expect(BUILTIN_FONT_FAMILIES).toContain('Arial')
    expect(BUILTIN_FONT_FAMILIES).toContain('Calibri')
    expect(BUILTIN_FONT_FAMILIES).toContain('Times New Roman')
    expect(BUILTIN_FONT_FAMILIES).toContain('Segoe UI')
  })

  it('filters candidates per locale so zh-TW only shows Traditional Chinese and Latin fonts', () => {
    const zhTW = fontFamiliesFor('zh-TW')

    // Essential Traditional Chinese fonts must be present
    expect(zhTW).toContain('微軟正黑體')
    expect(zhTW).toContain('新細明體')
    expect(zhTW).toContain('細明體')
    expect(zhTW).toContain('標楷體')
    expect(zhTW).toContain('Arial')
    expect(zhTW).toContain('Calibri')

    // Simplified Chinese fonts must NOT be shown in zh-TW menu
    expect(zhTW).not.toContain('等线')
    expect(zhTW).not.toContain('微软雅黑')
    expect(zhTW).not.toContain('方正小标宋简体')
    expect(zhTW).not.toContain('华文中宋')
    expect(zhTW).not.toContain('苹方-简')

    // Verify zh-TW is sorted according to zh-TW collation
    for (let i = 1; i < zhTW.length; i++) {
      const cmp = zhTW[i - 1]!.localeCompare(zhTW[i]!, 'zh-TW', {
        sensitivity: 'base',
        numeric: true,
      })
      expect(cmp).toBeLessThanOrEqual(0)
    }
  })

  it('filters candidates per locale so zh only shows Simplified Chinese and Latin fonts', () => {
    const zh = fontFamiliesFor('zh')

    // Essential Simplified Chinese fonts must be present
    expect(zh).toContain('等线')
    expect(zh).toContain('微软雅黑')
    expect(zh).toContain('宋体')
    expect(zh).toContain('Calibri')

    // Traditional Chinese specific fonts must not clutter zh menu
    expect(zh).not.toContain('微軟正黑體')
    expect(zh).not.toContain('新細明體')
    expect(zh).not.toContain('標楷體')
  })

  it('sortFontFamilies sorts arbitrary family arrays', () => {
    const sample = ['Times New Roman', 'Arial', '微軟正黑體', '新細明體', '標楷體']
    const sorted = sortFontFamilies(sample, 'zh-TW')
    expect(sorted).toEqual([
      '微軟正黑體',
      '新細明體',
      '標楷體',
      'Arial',
      'Times New Roman',
    ])
  })
})
