/**
 * A mouse selection must never rewrite cells (Excel).
 *
 * Univer arms two drag handles on the selection frame: the 6 px fill handle at
 * its bottom-right corner (drag = AutoFill: the source text is copied over the
 * cells dragged across) and the frame's border strip (drag = move the cell
 * contents there). Both sit under the pointer when a new selection drag starts
 * a pixel or two from the old frame - on the gridline next to the selected
 * cell, the common place to begin selecting the next range - so a plain
 * select-by-dragging silently copied or moved the old cell's text into the
 * range ("column A overwrites the others").
 *
 * Univer shows the grab in the cursor: crosshair over the fill handle, move
 * over the border. A press only grabs a handle when the pointer has rested on
 * it long enough for that cursor to have been seen; a press that follows a
 * quick pass over the handle is an ordinary selection start. For such a press
 * the handles are switched off (not evented) until the button is released, so
 * the press falls through to the cell below.
 *
 * Scope: only the ordinary selection's controls (ISheetSelectionRenderService). The reference
 * boxes drawn while a formula is being edited belong to a separate render service
 * (RefSelectionsRenderService), whose border drag moves a reference - Excel behavior the user
 * grabs deliberately and quickly - so they are deliberately left alone.
 */

/** Cursors Univer shows over the fill handle and over the frame border. */
const GRAB_CURSORS: ReadonlySet<string> = new Set(['crosshair', 'move'])

/** How long the pointer must rest on a handle before a press grabs it. */
export const GRAB_DWELL_MS = 150

interface EventedLike {
  evented: boolean
}

/** The slice of Univer's SelectionControl the guard touches. */
export interface GrabControlLike {
  readonly fillControl: EventedLike
  readonly leftControl: EventedLike
  readonly rightControl: EventedLike
  readonly topControl: EventedLike
  readonly bottomControl: EventedLike
}

type Listenable = Pick<Window, 'addEventListener' | 'removeEventListener'>

export interface SelectionGrabGuardOptions {
  readonly getControls: () => readonly GrabControlLike[]
  readonly target?: Listenable
  readonly dwellMs?: number
  readonly now?: () => number
}

function handlesOf(control: GrabControlLike): EventedLike[] {
  return [
    control.fillControl,
    control.leftControl,
    control.rightControl,
    control.topControl,
    control.bottomControl,
  ]
}

function cursorOf(event: Event): string {
  const element = event.target as { style?: { cursor?: string } } | null
  return element?.style?.cursor ?? ''
}

export function installSelectionGrabGuard(options: SelectionGrabGuardOptions): {
  dispose(): void
} {
  const target = options.target ?? window
  const dwellMs = options.dwellMs ?? GRAB_DWELL_MS
  const now = options.now ?? (() => performance.now())
  /** when the pointer came to rest on a handle (null: not on one) */
  let overSince: number | null = null
  /** handles switched off for the press in progress, with what to restore */
  let suspended: Array<{ rect: EventedLike; evented: boolean }> | null = null

  const restore = (): void => {
    if (!suspended) return
    for (const { rect, evented } of suspended) rect.evented = evented
    suspended = null
  }

  // A capture listener on the window also sees every ELEMENT blur (the cell editor, the formula
  // bar...) that happens mid-press; only the window itself losing focus ends a press.
  const onBlur = (event: Event): void => {
    if (event.target && event.target !== target) return
    restore()
  }

  const onMove = (event: Event): void => {
    if (GRAB_CURSORS.has(cursorOf(event))) overSince ??= now()
    else overSince = null
  }

  const onDown = (event: Event): void => {
    const button = (event as MouseEvent).button
    if (button !== undefined && button !== 0) return
    const deliberate =
      overSince !== null && GRAB_CURSORS.has(cursorOf(event)) && now() - overSince >= dwellMs
    overSince = null
    if (deliberate) return
    restore()
    suspended = []
    for (const control of options.getControls()) {
      for (const rect of handlesOf(control)) {
        if (rect.evented) suspended.push({ rect, evented: true })
        rect.evented = false
      }
    }
  }

  // capture: runs before Univer's own listeners on the canvas hit-test the press
  target.addEventListener('pointerdown', onDown, true)
  target.addEventListener('pointermove', onMove, false)
  target.addEventListener('pointerup', restore, true)
  target.addEventListener('pointercancel', restore, true)
  target.addEventListener('blur', onBlur, true)
  return {
    dispose() {
      target.removeEventListener('pointerdown', onDown, true)
      target.removeEventListener('pointermove', onMove, false)
      target.removeEventListener('pointerup', restore, true)
      target.removeEventListener('pointercancel', restore, true)
      target.removeEventListener('blur', onBlur, true)
      restore()
    },
  }
}
