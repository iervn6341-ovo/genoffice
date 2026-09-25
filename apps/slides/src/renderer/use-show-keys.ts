import { useEffect, useRef } from 'react'
import { showKeyCommand, type ShowKeyCommand } from './slideshow-utils'

/** What a show window does for each key command; the latest handlers are read on every key. */
export interface ShowKeyHandlers {
  onCommand: (command: ShowKeyCommand) => void
  /** 1-based slide number typed before Enter */
  onGoto: (slideNumber: number) => void
}

/** A typed slide number is forgotten after this long without another digit (PowerPoint drops it too). */
const GOTO_TIMEOUT_MS = 3000

/**
 * Capture-phase key handling shared by the show, presenter and audience windows
 * (capture beats the editor's generic shortcuts).
 */
export function useShowKeys(handlers: ShowKeyHandlers): void {
  const ref = useRef(handlers)
  ref.current = handlers
  useEffect(() => {
    let buffer = ''
    let timer: ReturnType<typeof setTimeout> | undefined
    const onKey = (e: KeyboardEvent) => {
      const r = showKeyCommand(e.key, buffer, {
        meta: e.metaKey,
        ctrl: e.ctrlKey,
        alt: e.altKey,
      })
      if (!r.handled) return
      e.preventDefault()
      buffer = r.buffer
      clearTimeout(timer)
      if (buffer) timer = setTimeout(() => (buffer = ''), GOTO_TIMEOUT_MS)
      if (r.command) ref.current.onCommand(r.command)
      else if (r.goto != null) ref.current.onGoto(r.goto)
    }
    window.addEventListener('keydown', onKey, true)
    return () => {
      window.removeEventListener('keydown', onKey, true)
      clearTimeout(timer)
    }
  }, [])
}
