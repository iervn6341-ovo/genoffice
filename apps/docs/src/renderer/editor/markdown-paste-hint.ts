/**
 * Handshake between the paste handler and the "Markdown detected" chip: handlePaste inserts the
 * clipboard text literally and stashes it here; the chip takes it when that paste's transaction
 * lands (same pattern as the paste-options payload).
 */
let pending: string | null = null

export function stashMarkdownPaste(text: string): void {
  pending = text
}

export function takeMarkdownPaste(): string | null {
  const text = pending
  pending = null
  return text
}
