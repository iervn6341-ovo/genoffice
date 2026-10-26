/**
 * Paste from the system clipboard when there is no native paste event (ribbon
 * Paste button, context menu): one path for both, running the editor's full
 * paste machinery (payload parse, paste modes, image priority) exactly like
 * Ctrl+V. The synthesized event carries a real clipboardData so the
 * handlePaste lanes behave as on a native paste.
 */
import type { Editor } from '@tiptap/core'
import { insertImageFromDataUrl } from '../components/ribbon-tabs'
import {
  beginForeignPaste,
  defaultPasteMode,
  forceNextPasteMode,
  stashPastePayload,
  type PasteMode,
} from './paste-options'

function pasteEvent(html: string | null, text: string): ClipboardEvent {
  const data = new DataTransfer()
  if (html) data.setData('text/html', html)
  if (text) data.setData('text/plain', text)
  return new ClipboardEvent('paste', { clipboardData: data })
}

export async function pasteFromSystemClipboard(ed: Editor, mode?: PasteMode): Promise<void> {
  try {
    for (const item of await navigator.clipboard.read()) {
      const text = item.types.includes('text/plain')
        ? await (await item.getType('text/plain')).text()
        : ''
      if (item.types.includes('text/html')) {
        const html = await (await item.getType('text/html')).text()
        if (html) {
          // arm the foreign-paste handshake exactly like a Ctrl+V — the
          // synthetic pasteHTML never fires the DOM paste handler, so
          // ribbon pastes skipped the paste mode and the r181 fill
          // Paste ▾ → Keep Source / Merge / Text Only decides this one paste
          if (mode) forceNextPasteMode(mode)
          if (beginForeignPaste(html)) {
            stashPastePayload({ html, text, mode: mode ?? defaultPasteMode() })
          }
          ed.view.pasteHTML(html, pasteEvent(html, text))
          ed.commands.focus()
          return
        }
      }
      // image priority mirrors Ctrl+V: an image wins over missing or
      // whitespace-only plain text (OS clipboards often advertise an
      // empty text/plain beside image/png)
      const imageType = item.types.find((type) => type.startsWith('image/'))
      if (imageType && !text.trim()) {
        const blob = await item.getType(imageType)
        const reader = new FileReader()
        reader.onload = () => {
          if (typeof reader.result === 'string') {
            void insertImageFromDataUrl(ed, reader.result, 'Image (pasted)')
          }
        }
        reader.readAsDataURL(blob)
        return
      }
    }
  } catch {
    /* clipboard.read unavailable/denied: plain-text fallback below */
  }
  const text = await navigator.clipboard.readText()
  if (text) {
    ed.view.pasteText(text, pasteEvent(null, text))
    ed.commands.focus()
  }
}

/** Paste Special ▸ Unformatted Text */
export async function pastePlainFromSystemClipboard(ed: Editor): Promise<void> {
  const text = await navigator.clipboard.readText()
  if (text) {
    ed.view.pasteText(text, pasteEvent(null, text))
    ed.commands.focus()
  }
}
