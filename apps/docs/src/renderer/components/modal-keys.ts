import { useEffect, useRef } from 'react'
import type { KeyboardEvent as ReactKeyboardEvent } from 'react'

/** The main document surface (tiptap's editorProps tag it `doc-page`) */
const EDITOR_SELECTOR = '.ProseMirror.doc-page'

/**
 * Give focus back after a modal closes. A dialog that was open on top of the document leaves
 * `document.activeElement` on BODY once it unmounts, so typing goes nowhere (Word returns to the
 * document with the selection intact and the next keystroke replaces it). Only acts when focus was
 * actually lost: a close that deliberately moved focus somewhere else is left alone.
 */
export function restoreFocusAfterModal(previous: Element | null): void {
  const active = document.activeElement
  if (active && active !== document.body) return
  const target =
    previous instanceof HTMLElement && previous.isConnected && previous !== document.body
      ? previous
      : document.querySelector<HTMLElement>(EDITOR_SELECTOR)
  target?.focus({ preventScroll: true })
}

/**
 * Shared modal keyboard behavior: Esc closes (stopped so it never reaches
 * global listeners like "exit read mode"), the first form control gets
 * focus on mount unless something inside is already focused (autoFocus),
 * and closing returns focus to where it was (the document).
 * Spread the returned ref/onKeyDown onto the modal backdrop element.
 */
export function useModalKeys(onClose: () => void) {
  const ref = useRef<HTMLDivElement>(null)
  const previousFocus = useRef<Element | null>(null)

  useEffect(() => {
    const el = ref.current
    if (previousFocus.current === null && !el?.contains(document.activeElement))
      previousFocus.current = document.activeElement
    if (!el || el.contains(document.activeElement)) return
    const first = el.querySelector<HTMLElement>('input, textarea, select, button')
    ;(first ?? el).focus()
  }, [])

  useEffect(
    () => () => {
      restoreFocusAfterModal(previousFocus.current)
    },
    [],
  )

  const onKeyDown = (e: ReactKeyboardEvent) => {
    if (e.key !== 'Escape') return
    e.preventDefault()
    e.stopPropagation()
    onClose()
  }

  return { ref, onKeyDown }
}
