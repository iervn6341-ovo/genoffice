/** A generated field is formatted as a unit when only its caret is active.
 * Ordinary text keeps the browser's insertion-format behavior. */
export function selectCollapsedFieldText(
  root = document.querySelector<HTMLElement>('.slide-text-editor[data-dynamic-field="true"]'),
  selection = window.getSelection(),
): boolean {
  if (
    !root ||
    root.dataset.dynamicField !== 'true' ||
    !root.textContent ||
    !selection?.isCollapsed ||
    !selection.anchorNode ||
    !root.contains(selection.anchorNode)
  )
    return false
  const range = document.createRange()
  range.selectNodeContents(root)
  selection.removeAllRanges()
  selection.addRange(range)
  return true
}
