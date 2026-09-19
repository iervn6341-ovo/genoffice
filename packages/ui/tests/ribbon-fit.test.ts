import { describe, expect, it } from 'vitest'
import { fitRibbonLabels, labelFromTip, type RibbonFitTarget } from '../src/ribbon-fit'

/** an element whose content width grows with the applied label level, like a real ribbon body */
function ribbon(clientWidth: number, widthAt: (level: number) => number): RibbonFitTarget {
  const dataset: DOMStringMap = {}
  return {
    dataset,
    clientWidth,
    get scrollWidth() {
      return widthAt(Number(dataset.labels ?? '0'))
    },
  }
}

// icons only = 1000px, each label tier adds 100px
const grows = (level: number) => 1000 + level * 100

describe('fitRibbonLabels', () => {
  it('shows every tier when the window is wide enough', () => {
    const el = ribbon(1400, grows)
    expect(fitRibbonLabels(el)).toBe(3)
    expect(el.dataset.labels).toBe('3')
  })

  it('backs off tier by tier as the window narrows', () => {
    // content is 1300 / 1200 / 1100 / 1000px at levels 3 / 2 / 1 / 0 (1px tolerance)
    expect(fitRibbonLabels(ribbon(1300, grows))).toBe(3)
    expect(fitRibbonLabels(ribbon(1298, grows))).toBe(2)
    expect(fitRibbonLabels(ribbon(1198, grows))).toBe(1)
    const el = ribbon(1098, grows)
    expect(fitRibbonLabels(el)).toBe(0)
    expect(el.dataset.labels).toBe('0')
  })

  it('falls back to icon-only when even that overflows (the ribbon then scrolls as before)', () => {
    const el = ribbon(600, grows)
    expect(fitRibbonLabels(el)).toBe(0)
    expect(el.dataset.labels).toBe('0')
  })

  it('tolerates a 1px sub-pixel overflow', () => {
    expect(fitRibbonLabels(ribbon(1300, (l) => 1000 + l * 100 + 0.5))).toBe(3)
    expect(fitRibbonLabels(ribbon(1300, (l) => 1000 + l * 100 + 2))).toBe(2)
  })

  it('re-fits from the top each time, so widening the window brings labels back', () => {
    const el = ribbon(1000, grows)
    expect(fitRibbonLabels(el)).toBe(0)
    el.clientWidth = 1500
    expect(fitRibbonLabels(el)).toBe(3)
  })

  it('honours a smaller maxLevel', () => {
    expect(fitRibbonLabels(ribbon(5000, grows), 1)).toBe(1)
  })
})

describe('labelFromTip', () => {
  it.each([
    ['Cut (⌘X)', 'Cut'],
    ['Copy (Ctrl+C)', 'Copy'],
    ['剪下 (⌘X)', '剪下'],
    ['複製格式：先按這裡，再選取要套用格式的文字', '複製格式'],
    ['Format Painter: click here, then select the text to format', 'Format Painter'],
    ['قص (⌘X)', 'قص'],
    ['切り取り（⌘X）', '切り取り'],
    ['Format Painter (⌘⌥C): click here', 'Format Painter'],
    ['  Paste  ', 'Paste'],
    ['Bold', 'Bold'],
  ])('%s -> %s', (tip, label) => {
    expect(labelFromTip(tip)).toBe(label)
  })
})
