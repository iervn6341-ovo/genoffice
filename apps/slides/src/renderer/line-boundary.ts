/**
 * Cmd+←/→ (macOS) and Home/End: move the caret to the start / end of the visual line.
 *
 * The editing overlay wraps every word in an `inline-block` layout fragment (so the text sits where the
 * canvas laid it out), and Chromium's own line-boundary movement treats each fragment as its own line:
 * Cmd+← stopped at the start of the current word and Cmd+→ did not move at all. The visual line is
 * found here instead: every text node whose box sits at the caret's height belongs to it.
 */
const visualLineOf = (block: HTMLElement, caretMidY: number): Text[] => {
  const nodes: Text[] = []
  const walker = document.createTreeWalker(block, NodeFilter.SHOW_TEXT)
  const range = document.createRange()
  for (let n = walker.nextNode() as Text | null; n; n = walker.nextNode() as Text | null) {
    if (!n.data.length) continue
    range.selectNodeContents(n)
    const rect = range.getClientRects()[0]
    if (!rect || rect.height === 0) continue
    if (Math.abs((rect.top + rect.bottom) / 2 - caretMidY) <= rect.height / 2) nodes.push(n)
  }
  return nodes
}

/** The paragraph div (a direct child of the editor) holding a node. */
const blockOf = (root: HTMLElement, node: Node): HTMLElement | null => {
  let cur: Node | null = node
  while (cur && cur.parentNode !== root) cur = cur.parentNode
  return cur instanceof HTMLElement ? cur : null
}

/** True when handled; false leaves the key to the browser (right-to-left, vertical, no layout). */
export function moveToLineBoundary(
  root: HTMLElement,
  where: 'start' | 'end',
  extend: boolean,
): boolean {
  const sel = window.getSelection()
  if (!sel || !sel.rangeCount || !sel.focusNode || !root.contains(sel.focusNode)) return false
  const block = blockOf(root, sel.focusNode)
  if (!block) return false
  const style = window.getComputedStyle(block)
  if (style.direction === 'rtl' || style.writingMode.startsWith('vertical')) return false

  const caret = document.createRange()
  caret.setStart(sel.focusNode, sel.focusOffset)
  caret.collapse(true)
  let rect = caret.getClientRects()[0]
  if (!rect || rect.height === 0) {
    // a caret in an empty or element position: its own box gives the line
    const el = sel.focusNode instanceof HTMLElement ? sel.focusNode : sel.focusNode.parentElement
    rect = (el ?? block).getBoundingClientRect()
  }
  const line = visualLineOf(block, (rect.top + rect.bottom) / 2)
  if (!line.length) return false

  let target: Text
  let offset: number
  if (where === 'start') {
    target = line[0]!
    offset = 0
  } else {
    // the end of the last word: a wrapped line's trailing space stays off the caret's reach
    const last = [...line].reverse().find((n) => n.data.trim().length > 0) ?? line[line.length - 1]!
    target = last
    offset = last.data.replace(/\s+$/, '').length || last.data.length
  }
  if (extend) {
    sel.setBaseAndExtent(sel.anchorNode ?? target, sel.anchorOffset, target, offset)
  } else {
    sel.collapse(target, offset)
  }
  return true
}
