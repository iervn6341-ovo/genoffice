import { KEEP_EDIT_ATTR, classifyEditPress } from '@genoffice/ui'

/**
 * Excel's ribbon never takes keyboard focus from the workbook: after Bold, a fill colour or any
 * other ribbon button the next keystroke still goes to the grid (or to the in-cell editor, with its
 * text selection intact), so formatting can be chained and typing carries on.
 *
 * Here that is the shared keep-edit rule (packages/ui keep-edit): the ribbon, its dropdowns and the
 * colour popups are `data-keep-edit` zones; a press on a plain control in a zone must not move
 * focus (mousedown default prevented - the click still fires), a press on a field (the font name
 * and size boxes, selects, links) lets the field take focus as usual.
 *
 * The in-cell editor is Univer's, with its own selection model: nothing here can restore a DOM
 * range, so the way to keep its selection alive is not to blur it in the first place.
 */

const TEXT_FIELD = 'input, textarea, select, [contenteditable]'

/** Whether a mousedown on `target` must leave focus where it is (a plain control in a zone). */
export function shouldKeepGridFocus(target: EventTarget | null): boolean {
  if (classifyEditPress(null, target) !== 'neutral') return false
  // native colour pickers hang off a <label>: preventing the press would not stop them but
  // is not worth the risk either (same exemption as keepEditorFocusOnRibbonPress)
  return !(target as Element).closest('label')
}

/**
 * The ribbon box (font name / size) holding a typed-but-uncommitted value: the next ribbon press
 * must let its blur commit the value, and afterwards focus returns to the grid like Excel's.
 */
function typingInZone(active: Element | null): boolean {
  return (
    !!active &&
    active !== document.body &&
    !!active.closest(TEXT_FIELD) &&
    !!active.closest(`[${KEEP_EDIT_ATTR}]`)
  )
}

export interface KeepGridFocusOptions {
  /** Hand keyboard focus back to the grid (the shell's `focus-grid` command) */
  focusGrid: () => void
  readonly doc?: Document
}

export function installKeepGridFocus(options: KeepGridFocusOptions): () => void {
  const doc = options.doc ?? document
  const onDown = (event: MouseEvent): void => {
    if (event.button !== 0) return
    if (!shouldKeepGridFocus(event.target)) return
    if (typingInZone(doc.activeElement)) {
      // let the field blur (commits its draft), then give the grid focus once the click is done
      doc.addEventListener('mouseup', () => setTimeout(options.focusGrid, 0), { once: true })
      return
    }
    event.preventDefault()
  }
  doc.addEventListener('mousedown', onDown, true)
  return () => doc.removeEventListener('mousedown', onDown, true)
}
