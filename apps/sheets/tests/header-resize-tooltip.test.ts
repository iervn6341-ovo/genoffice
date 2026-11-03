// @vitest-environment jsdom
import { describe, expect, it } from 'vitest'
import {
  formatResizeTooltip,
  hideResizeTooltip,
  pixelsToExcelColWidth,
  pixelsToExcelRowHeight,
  showResizeTooltip,
} from '../src/renderer/header-resize-tooltip'

describe('header-resize-tooltip calculations', () => {
  it('converts pixels to Excel column width characters matching Excel', () => {
    // 89 pixels at MDW 6 (Mac Excel default): (89 - 5) / 6 = 14.00
    expect(pixelsToExcelColWidth(89, 6)).toBe(14)

    // 72 pixels (classic Excel column width) at MDW 7: (72 - 5) / 7 = 9.57
    expect(pixelsToExcelColWidth(72, 7)).toBe(9.57)

    // 0 pixels
    expect(pixelsToExcelColWidth(0, 7)).toBe(0)
    expect(pixelsToExcelColWidth(-10, 7)).toBe(0)

    // Narrow column (< mdw + 5)
    expect(pixelsToExcelColWidth(6, 7)).toBeCloseTo(6 / 12, 2)
  })

  it('converts pixels to Excel row height points matching Excel', () => {
    // 20 pixels = 15.00 pt (Excel default row height)
    expect(pixelsToExcelRowHeight(20)).toBe(15)

    // 27 pixels = 20.25 pt
    expect(pixelsToExcelRowHeight(27)).toBe(20.25)

    // 32 pixels = 24.00 pt
    expect(pixelsToExcelRowHeight(32)).toBe(24)

    // 0 pixels
    expect(pixelsToExcelRowHeight(0)).toBe(0)
    expect(pixelsToExcelRowHeight(-5)).toBe(0)
  })

  it('formats column resize tooltip per locale', () => {
    // English
    expect(formatResizeTooltip('column', 89, 'en', 6)).toBe('Width: 14.00 (89 pixels)')
    expect(formatResizeTooltip('column', 1, 'en', 6)).toBe('Width: 0.09 (1 pixel)')

    // Traditional Chinese (zh-TW)
    expect(formatResizeTooltip('column', 89, 'zh-TW', 6)).toBe('寬度: 14.00 (89 像素)')

    // Simplified Chinese (zh)
    expect(formatResizeTooltip('column', 89, 'zh', 6)).toBe('宽度: 14.00 (89 像素)')

    // Japanese
    expect(formatResizeTooltip('column', 89, 'ja', 6)).toBe('幅: 14.00 (89 ピクセル)')
  })

  it('formats row resize tooltip per locale', () => {
    // English
    expect(formatResizeTooltip('row', 27, 'en')).toBe('Height: 20.25 (27 pixels)')
    expect(formatResizeTooltip('row', 1, 'en')).toBe('Height: 0.75 (1 pixel)')

    // Traditional Chinese (zh-TW)
    expect(formatResizeTooltip('row', 20, 'zh-TW')).toBe('高度: 15.00 (20 像素)')

    // Simplified Chinese (zh)
    expect(formatResizeTooltip('row', 20, 'zh')).toBe('高度: 15.00 (20 像素)')

    // Japanese
    expect(formatResizeTooltip('row', 20, 'ja')).toBe('高さ: 15.00 (20 ピクセル)')
  })

  it('manages tooltip DOM element show and hide', () => {
    showResizeTooltip({
      kind: 'column',
      pixels: 89,
      clientX: 250,
      clientY: 100,
      canvasRect: { top: 80, left: 50, right: 800, bottom: 600, width: 750, height: 520, x: 50, y: 80, toJSON: () => ({}) },
      lang: 'en',
      mdw: 6,
    })

    const el = document.getElementById('sheets-header-resize-tooltip')
    expect(el).not.toBeNull()
    expect(el?.textContent).toBe('Width: 14.00 (89 pixels)')
    expect(el?.style.display).toBe('block')
    expect(el?.style.left).toBe('250px')

    hideResizeTooltip()
    expect(el?.style.display).toBe('none')
  })
})
