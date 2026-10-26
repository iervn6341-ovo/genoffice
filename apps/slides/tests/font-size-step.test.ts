/**
 * Grow/shrink font while editing: sizes step along the PowerPoint ladder (no ×1.2
 * fractional drift) and the DOM selection survives the span replacement so the
 * highlight stays and repeated clicks keep working.
 */
import { describe, it, expect, beforeAll, afterEach } from 'vitest'
import {
  liveFontSize,
  resizeSelectionFont,
  setSelectionFontSizePt,
} from '../src/renderer/TextEditOverlay'

// jsdom has no execCommand: emulate the browser product — wrap the selection in <font size="7">.
// An element-only selection (what reselectSpans produces) wraps that element's contents, matching
// how Chromium nests the font inside the styled span.
beforeAll(() => {
  document.execCommand = (cmd: string) => {
    if (cmd !== 'fontSize') return true
    const sel = window.getSelection()
    if (!sel?.rangeCount || sel.getRangeAt(0).collapsed) return false
    const range = sel.getRangeAt(0)
    const onlyChild =
      range.startContainer === range.endContainer &&
      range.endOffset - range.startOffset === 1 &&
      range.startContainer.childNodes[range.startOffset]
    const wrap = document.createRange()
    if (onlyChild instanceof HTMLElement) wrap.selectNodeContents(onlyChild)
    else wrap.setStart(range.startContainer, range.startOffset)
    if (!(onlyChild instanceof HTMLElement)) wrap.setEnd(range.endContainer, range.endOffset)
    const font = document.createElement('font')
    font.setAttribute('size', '7')
    wrap.surroundContents(font)
    return true
  }
})

afterEach(() => {
  document.body.innerHTML = ''
  window.getSelection()?.removeAllRanges()
})

function setupEditor(px: number): { root: HTMLElement; span: HTMLElement } {
  const root = document.createElement('div')
  root.tabIndex = 0
  Object.defineProperty(root, 'isContentEditable', { value: true })
  const span = document.createElement('span')
  span.style.fontSize = `${px}px`
  span.textContent = 'Annual timeline'
  root.appendChild(span)
  document.body.appendChild(root)
  root.focus()
  const range = document.createRange()
  range.selectNodeContents(span)
  const sel = window.getSelection()!
  sel.removeAllRanges()
  sel.addRange(range)
  return { root, span }
}

function currentSizePx(root: HTMLElement): number {
  const spans = root.querySelectorAll('span')
  return parseFloat((spans[spans.length - 1] as HTMLElement).style.fontSize)
}

describe('resizeSelectionFont ladder stepping', () => {
  it('grows along the ladder: 45pt → 48 → 54 (not ×1.2)', () => {
    const { root } = setupEditor(60) // 60px = 45pt
    resizeSelectionFont(1)
    expect(currentSizePx(root)).toBe(64) // 48pt
    resizeSelectionFont(1)
    expect(currentSizePx(root)).toBe(72) // 54pt
  })

  it('shrinks along the ladder: 45pt → 44', () => {
    const { root } = setupEditor(60)
    resizeSelectionFont(-1)
    expect(currentSizePx(root)).toBe((44 * 96) / 72)
  })

  it('steps ±10pt beyond the ladder and clamps at 1pt', () => {
    const { root } = setupEditor((100 * 96) / 72)
    resizeSelectionFont(1)
    expect(currentSizePx(root)).toBe((110 * 96) / 72)

    const { root: small } = setupEditor((8 * 96) / 72)
    resizeSelectionFont(-1)
    expect(currentSizePx(small)).toBe((7 * 96) / 72)

    const { root: tiny } = setupEditor((1 * 96) / 72)
    resizeSelectionFont(-1)
    expect(currentSizePx(tiny)).toBe((1 * 96) / 72)
  })

  it('a size with px rounding noise still moves a whole stop on every click', () => {
    // 12pt authored, read back from layout as 15.999px → must not read as "just under 12 → 12"
    const { root } = setupEditor(15.999)
    resizeSelectionFont(1)
    expect(currentSizePx(root)).toBe((14 * 96) / 72)
    resizeSelectionFont(1)
    expect(currentSizePx(root)).toBe((16 * 96) / 72)
  })

  it('keeps the text selected after each step', () => {
    const { root } = setupEditor(60)
    resizeSelectionFont(1)
    const sel = window.getSelection()!
    expect(sel.isCollapsed).toBe(false)
    expect(sel.toString()).toBe('Annual timeline')
    expect(root.contains(sel.getRangeAt(0).commonAncestorContainer)).toBe(true)
  })
})

describe('setSelectionFontSizePt', () => {
  it('applies the absolute size and keeps the selection', () => {
    const { root } = setupEditor(60)
    setSelectionFontSizePt(28)
    expect(currentSizePx(root)).toBe((28 * 96) / 72)
    const sel = window.getSelection()!
    expect(sel.isCollapsed).toBe(false)
    expect(sel.toString()).toBe('Annual timeline')
  })
})

describe('liveFontSize (ribbon size box while editing)', () => {
  // The editor DOM carries [data-src-para] paragraph blocks; norm = viewport px per model px
  function paraEditor(
    runs: Array<{ text: string; px: number }>,
    norm = 1,
  ): { root: HTMLElement; spans: HTMLElement[] } {
    const root = document.createElement('div')
    root.tabIndex = 0
    root.dataset.norm = String(norm)
    Object.defineProperty(root, 'isContentEditable', { value: true })
    const para = document.createElement('div')
    para.setAttribute('data-src-para', '0')
    const spans = runs.map((r) => {
      const span = document.createElement('span')
      span.style.fontSize = `${r.px}px`
      span.textContent = r.text
      para.appendChild(span)
      return span
    })
    root.appendChild(para)
    document.body.appendChild(root)
    root.focus()
    return { root, spans }
  }
  const select = (from: Node, to: Node) => {
    const range = document.createRange()
    range.setStartBefore(from)
    range.setEndAfter(to)
    const sel = window.getSelection()!
    sel.removeAllRanges()
    sel.addRange(range)
  }

  it('reads the selected run, not the element the selection happens to sit in', () => {
    const { spans } = paraEditor([{ text: 'Alpha', px: 16 }]) // 12pt
    select(spans[0]!, spans[0]!)
    expect(liveFontSize()).toEqual({ sizePt: 12, mixed: false })
    resizeSelectionFont(1)
    // the selection now sits at the paragraph (parent of the new span): still the new size
    expect(liveFontSize()).toEqual({ sizePt: 14, mixed: false })
    resizeSelectionFont(1)
    expect(liveFontSize()).toEqual({ sizePt: 16, mixed: false })
    resizeSelectionFont(-1)
    expect(liveFontSize()).toEqual({ sizePt: 14, mixed: false })
  })

  it('divides the viewport scale out (norm) and rounds to half points', () => {
    const { spans } = paraEditor([{ text: 'Alpha', px: 16 * 1.5 + 0.1 }], 1.5)
    select(spans[0]!, spans[0]!)
    expect(liveFontSize()).toEqual({ sizePt: 12, mixed: false })
  })

  it('a mixed selection reports the smallest size and flags it', () => {
    const { spans } = paraEditor([
      { text: 'Alpha', px: 16 },
      { text: 'Beta', px: 24 },
    ])
    select(spans[0]!, spans[1]!)
    expect(liveFontSize()).toEqual({ sizePt: 12, mixed: true })
  })

  it('a caret reads the size it will type with', () => {
    const { spans } = paraEditor([{ text: 'Alpha', px: 24 }])
    const range = document.createRange()
    range.setStart(spans[0]!.firstChild!, 2)
    range.collapse(true)
    const sel = window.getSelection()!
    sel.removeAllRanges()
    sel.addRange(range)
    expect(liveFontSize()).toEqual({ sizePt: 18, mixed: false })
  })
})
