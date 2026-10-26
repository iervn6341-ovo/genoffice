/**
 * Keep an in-place text edit (and its selection) alive while the user works in the ribbon or a task
 * pane — the behavior of Office apps, shared by every editor that owns a contentEditable surface.
 *
 * The failure it prevents: the ribbon is full of controls that take real focus (font name / size
 * boxes, colour inputs, menu items). The moment focus leaves the editor its selection is gone and a
 * naive "blur = commit" ends the edit, so a font change either hits nothing or drops the user out of
 * the cell / shape. The fix has three cooperating parts:
 *
 *  1. Neutral zones — mark the ribbon root, task panes and text dialogs with `data-keep-edit`.
 *  2. `installKeepEditGuard` — a capture-phase press handler for one live edit session. A press on a
 *     plain control inside a zone keeps (or hands back) focus on the editor; a press on a field lets
 *     the field take focus; only a press outside every zone commits.
 *  3. A selection keeper — `handleEditorBlur` snapshots the selection when focus leaves the editor
 *     for a zone, paints it with the CSS Custom Highlight API so the text still looks selected, and
 *     `restore()` puts it back right before a command runs.
 *
 * Nothing here knows about slides, cells or shapes; the host supplies the editor element and what
 * "commit" means. `keepEditorFocusOnRibbonPress` (ribbon-focus.ts) is the older, button-only variant.
 */

export const KEEP_EDIT_ATTR = 'data-keep-edit'

/** Controls that take real focus on press (text entry, pickers); everything else in a zone leaves focus alone */
const FOCUS_FIELD = 'input, textarea, select, [contenteditable], a[href]'

export type EditPress = 'in-editor' | 'neutral' | 'neutral-field' | 'outside'

/**
 * Where a mouse press landed relative to an edit session. `frame` is the editor's own box (the press
 * is the editor's business); a press inside a `data-keep-edit` zone is neutral — on a field the field
 * takes focus, on anything else the editor must keep (or get back) focus.
 */
export function classifyEditPress(
  frame: HTMLElement | null,
  target: EventTarget | null,
): EditPress {
  const t = target instanceof Element ? target : null
  if (!t) return 'outside'
  if (frame?.contains(t)) return 'in-editor'
  if (!t.closest(`[${KEEP_EDIT_ATTR}]`)) return 'outside'
  return t.closest(FOCUS_FIELD) ? 'neutral-field' : 'neutral'
}

interface HighlightRegistry {
  set(name: string, value: unknown): void
  delete(name: string): void
}
const highlights = (): HighlightRegistry | undefined =>
  (globalThis.CSS as unknown as { highlights?: HighlightRegistry } | undefined)?.highlights
const HighlightCtor = (): (new (...r: Range[]) => unknown) | undefined =>
  (globalThis as unknown as { Highlight?: new (...r: Range[]) => unknown }).Highlight

export interface SelectionKeeper {
  /** Snapshot the selection as focus leaves `root` (document selection must still sit inside it) */
  remember(root: HTMLElement): void
  /** Snapshot the live selection of the focused contentEditable (before a field takes focus) */
  save(): void
  /** Re-focus the editor and re-select the snapshot; false when there is none (or the editor is gone) */
  restore(): boolean
  /** Replace the snapshot (the editor re-populated and carried the selection over) */
  set(root: HTMLElement, range: Range): void
  /** The current snapshot, if any */
  peek(): { root: HTMLElement; range: Range } | null
  /** Drop the snapshot (the session ended) */
  forget(): void
  /** Remove the "kept selection" paint (the editor has focus again) */
  clearPaint(): void
}

/**
 * One keeper per editor kind. `highlightName` is the `::highlight(<name>)` rule the host styles
 * (background Highlight, color HighlightText) so the kept selection stays visible while a ribbon
 * field holds focus.
 */
export function createSelectionKeeper(highlightName = 'go-kept-selection'): SelectionKeeper {
  let saved: { root: HTMLElement; range: Range } | null = null
  const paint = (range: Range): void => {
    const reg = highlights()
    const H = HighlightCtor()
    if (!reg || !H || range.collapsed) return
    reg.set(highlightName, new H(range))
  }
  const clearPaint = (): void => highlights()?.delete(highlightName)
  return {
    remember(root) {
      const sel = window.getSelection()
      if (!sel?.rangeCount) return
      const range = sel.getRangeAt(0)
      if (!root.contains(range.startContainer) || !root.contains(range.endContainer)) return
      saved = { root, range: range.cloneRange() }
      paint(saved.range)
    },
    save() {
      const root = document.activeElement
      const sel = window.getSelection()
      if (!(root instanceof HTMLElement) || !root.isContentEditable || !sel?.rangeCount) return
      saved = { root, range: sel.getRangeAt(0).cloneRange() }
    },
    restore() {
      if (!saved?.root.isConnected) return false
      clearPaint()
      saved.root.focus()
      const sel = window.getSelection()
      if (!sel) return false
      sel.removeAllRanges()
      sel.addRange(saved.range)
      return true
    },
    set(root, range) {
      saved = { root, range }
      paint(range)
    },
    peek: () => saved,
    forget() {
      saved = null
      clearPaint()
    },
    clearPaint,
  }
}

/**
 * The editor's blur handler: focus moving into a keep-edit zone parks the selection instead of
 * committing. Returns true when the blur was absorbed (the host must not commit).
 */
export function handleEditorBlur(
  e: { relatedTarget: EventTarget | null; currentTarget: EventTarget | null },
  keeper: SelectionKeeper,
): boolean {
  const to = e.relatedTarget instanceof HTMLElement ? e.relatedTarget : null
  if (!to?.closest(`[${KEEP_EDIT_ATTR}]`)) return false
  if (e.currentTarget instanceof HTMLElement) keeper.remember(e.currentTarget)
  return true
}

export interface KeepEditGuardOptions {
  /** The contentEditable surface of the live session (null while it is unmounting) */
  editor: () => HTMLElement | null
  /** The editor's own box (presses inside are left alone); defaults to the editor's parent element */
  frame?: () => HTMLElement | null
  keeper: SelectionKeeper
  /** A press outside every zone while focus is parked off the editor: end the edit */
  commit: () => void
}

/**
 * Install the capture-phase press guard for one live edit session; returns the uninstaller (which
 * also forgets the kept selection). While the editor holds focus an outside press ends the edit
 * through the editor's own blur, so the guard only commits when focus is already parked.
 */
export function installKeepEditGuard(opts: KeepEditGuardOptions): () => void {
  const onDown = (ev: MouseEvent): void => {
    const div = opts.editor()
    if (!div) return
    const frame = opts.frame ? opts.frame() : div.parentElement
    const kind = classifyEditPress(frame, ev.target)
    if (kind === 'in-editor' || kind === 'neutral-field') return
    const focusInEditor = div.contains(document.activeElement)
    if (kind === 'outside') {
      if (!focusInEditor) opts.commit()
      return
    }
    if (!focusInEditor) opts.keeper.restore()
    // buttons, menu items and blank panel area must not steal focus from the editor
    ev.preventDefault()
  }
  document.addEventListener('mousedown', onDown, true)
  return () => {
    document.removeEventListener('mousedown', onDown, true)
    opts.keeper.forget()
  }
}
