/**
 * Home → Font extras while editing text: Character Spacing, Text Highlight and Change Case act
 * on the selection only, keep the surrounding formatting, and reach the commit payload —
 * while untouched text commits nothing new.
 */
import { describe, it, expect, afterEach } from 'vitest'
import {
  changeSelectionCase,
  extractParagraphs,
  setSelectionHighlight,
  setSelectionLetterSpacingPt,
} from '../src/renderer/TextEditOverlay'
import { applyEditParagraphs } from '@genoffice/pptx-ops'

afterEach(() => {
  document.body.innerHTML = ''
  window.getSelection()?.removeAllRanges()
})

/** One paragraph "hello world" in a 24px run, as populateEditorDom would produce it. */
function editor(): { root: HTMLElement; text: Text } {
  const root = document.createElement('div')
  root.tabIndex = 0
  Object.defineProperty(root, 'isContentEditable', { value: true })
  root.dataset.norm = '1'
  const p = document.createElement('div')
  p.dataset.srcPara = '0'
  const span = document.createElement('span')
  span.dataset.srcRun = '0'
  span.style.fontSize = '24px'
  span.textContent = 'hello world'
  p.appendChild(span)
  root.appendChild(p)
  document.body.appendChild(root)
  root.focus()
  return { root, text: span.firstChild as Text }
}

function select(node: Text, start: number, end: number) {
  const r = document.createRange()
  r.setStart(node, start)
  r.setEnd(node, end)
  const sel = window.getSelection()!
  sel.removeAllRanges()
  sel.addRange(r)
}

describe('editing-mode font extras', () => {
  it('an untouched round trip carries no spacing or highlight', () => {
    const { root } = editor()
    const runs = extractParagraphs(root, 1)[0]!.runs
    expect(runs.every((r) => r.letterSpacing === undefined && r.highlight === undefined)).toBe(true)
  })

  it('highlight marks only the selected word and keeps its font size', () => {
    const { root, text } = editor()
    select(text, 6, 11) // "world"
    setSelectionHighlight('#FFFF00')
    const runs = extractParagraphs(root, 1)[0]!.runs
    expect(runs.map((r) => [r.text, r.highlight, r.fontSize])).toEqual([
      ['hello ', undefined, 18],
      ['world', '#FFFF00', 18],
    ])
    // the selection survives, so a second click (e.g. No Color) hits the same text
    expect(window.getSelection()!.toString()).toBe('world')
    setSelectionHighlight(null)
    expect(extractParagraphs(root, 1)[0]!.runs.at(-1)!.highlight).toBeNull()
  })

  it('character spacing marks the selection in pt and previews it in px', () => {
    const { root, text } = editor()
    select(text, 0, 5)
    setSelectionLetterSpacingPt(3)
    const span = root.querySelector<HTMLElement>('[data-spc]')!
    expect(span.style.letterSpacing).toBe('4px') // 3pt at 96dpi
    const runs = extractParagraphs(root, 1)[0]!.runs
    expect(runs[0]).toMatchObject({ text: 'hello', letterSpacing: 3 })
    expect(runs[1]!.letterSpacing).toBeUndefined()
  })

  it('Change Case rewrites the selected text in place', () => {
    const { root, text } = editor()
    select(text, 0, 11)
    changeSelectionCase('title')
    expect(root.textContent).toBe('Hello World')
    changeSelectionCase('upper')
    expect(root.textContent).toBe('HELLO WORLD')
    // run structure is untouched: still one run traced to the source run
    const runs = extractParagraphs(root, 1)[0]!.runs
    expect(runs).toHaveLength(1)
    expect(runs[0]).toMatchObject({ text: 'HELLO WORLD', srcRun: 0 })
  })

  it('the commit maps the marks onto the model runs', () => {
    const { root, text } = editor()
    select(text, 6, 11)
    setSelectionHighlight('#00FF00')
    setSelectionLetterSpacingPt(-1.5)
    const edited = extractParagraphs(root, 1)
    const out = applyEditParagraphs([{ runs: [{ text: 'hello world', fontSize: 18 }] }], edited)
    const world = out[0]!.runs.find((r) => r.text === 'world')!
    expect(world).toMatchObject({
      highlight: '#00FF00',
      highlightEdited: true,
      letterSpacing: -1.5,
    })
    const hello = out[0]!.runs.find((r) => r.text === 'hello ')!
    expect(hello.highlight).toBeUndefined()
    expect(hello.highlightEdited).toBeUndefined()
  })
})
