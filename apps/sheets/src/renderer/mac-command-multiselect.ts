/**
 * ⌘+click adds a range to the selection on a Mac (Excel); Ctrl+click is the right button there.
 *
 * Univer's selection layer only knows the Windows convention and reads `ctrlKey` off the pointer
 * event to decide "add to the selection instead of replacing it" (cells, column and row headers).
 * So on a Mac a ⌘ press on the grid is presented to Univer as the additive modifier: in the capture
 * phase, before its canvas listeners run, the event object's `ctrlKey` reads true for that event.
 * Only presses on the grid canvas itself, and only the pointer events of the selection gesture
 * (down / move / up), are touched; keyboard shortcuts and every other surface keep their keys.
 */

const POINTER_EVENTS = ['pointerdown', 'pointermove', 'pointerup'] as const

export interface CommandClickLike {
  readonly metaKey: boolean
  readonly ctrlKey: boolean
  readonly altKey: boolean
  readonly target: EventTarget | null
}

/** Whether Univer should see this pointer event as Ctrl-modified. */
export function shouldPresentAsCtrl(event: CommandClickLike, isMac: boolean): boolean {
  if (!isMac || !event.metaKey || event.ctrlKey || event.altKey) return false
  const target = event.target as { tagName?: string; closest?(selector: string): unknown } | null
  return (
    !!target &&
    typeof target.closest === 'function' &&
    target.tagName === 'CANVAS' &&
    !!target.closest('#univer-container')
  )
}

export function installMacCommandMultiSelect(
  isMac: boolean,
  target: Pick<Window, 'addEventListener' | 'removeEventListener'> = window,
): () => void {
  if (!isMac) return () => {}
  const onPointer = (event: Event): void => {
    if (!shouldPresentAsCtrl(event as unknown as CommandClickLike, true)) return
    Object.defineProperty(event, 'ctrlKey', { value: true, configurable: true })
  }
  for (const type of POINTER_EVENTS) target.addEventListener(type, onPointer, true)
  return () => {
    for (const type of POINTER_EVENTS) target.removeEventListener(type, onPointer, true)
  }
}
