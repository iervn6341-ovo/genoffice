import { useEffect, useRef } from 'react'

/// The open dialogs, the one on top last: one Esc closes that one only.
const open: { close: () => void }[] = []
let listening = false

function onKeyDown(event: KeyboardEvent): void {
  if (event.key !== 'Escape' || event.defaultPrevented || event.isComposing) return
  const top = open[open.length - 1]
  if (!top) return
  event.preventDefault()
  top.close()
}

/// Esc closes a modal dialog wherever the focus is. A dialog that opens without taking the focus (the
/// key still goes to the grid or to the button that opened it) never sees the key on its own element, so
/// the listener sits on the window. A dialog that handles Esc itself, with the focus inside it, stops the
/// event first and this does not run twice.
export function useDialogEscape(close: () => void, enabled = true): void {
  const closeRef = useRef(close)
  useEffect(() => {
    closeRef.current = close
  })
  useEffect(() => {
    if (!enabled) return
    const entry = { close: () => closeRef.current() }
    open.push(entry)
    if (!listening) {
      window.addEventListener('keydown', onKeyDown)
      listening = true
    }
    return () => {
      open.splice(open.indexOf(entry), 1)
      if (open.length === 0) {
        window.removeEventListener('keydown', onKeyDown)
        listening = false
      }
    }
  }, [enabled])
}
