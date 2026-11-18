import { afterEach, expect, it } from 'vitest'
import { selectCollapsedFieldText } from '../src/renderer/field-format-selection'
afterEach(() => {
  window.getSelection()?.removeAllRanges()
  document.body.innerHTML = ''
})
function editor(field = true) {
  const root = document.createElement('div')
  root.contentEditable = 'true'
  root.dataset.dynamicField = String(field)
  root.innerHTML = '<span>Page </span><span>2</span>'
  document.body.append(root)
  const range = document.createRange()
  range.selectNodeContents(root.lastChild!)
  range.collapse(false)
  const selection = window.getSelection()!
  selection.removeAllRanges()
  selection.addRange(range)
  return { root, selection }
}
it('selects a generated field for color changes without making it plain text', () => {
  const { root, selection } = editor()
  const html = root.innerHTML
  expect(selectCollapsedFieldText(root, selection)).toBe(true)
  expect(selection.toString()).toBe('Page 2')
  expect(root.innerHTML).toBe(html)
})
it('keeps insertion color semantics for ordinary text', () => {
  const { root, selection } = editor(false)
  expect(selectCollapsedFieldText(root, selection)).toBe(false)
  expect(selection.isCollapsed).toBe(true)
})
it('keeps an explicit text selection and a selection outside the field', () => {
  const { root, selection } = editor()
  const range = document.createRange()
  range.selectNodeContents(root.lastChild!)
  selection.removeAllRanges()
  selection.addRange(range)
  expect(selectCollapsedFieldText(root, selection)).toBe(false)
  expect(selection.toString()).toBe('2')
  range.selectNodeContents(document.body)
  range.collapse(true)
  selection.removeAllRanges()
  selection.addRange(range)
  expect(selectCollapsedFieldText(root, selection)).toBe(false)
})
