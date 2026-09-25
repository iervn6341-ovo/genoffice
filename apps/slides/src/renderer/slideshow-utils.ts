/**
 * Pure logic for slide shows: playback sequence computation + rehearsal timing accumulation.
 * Extracted from SlideShowView for unit testing (no React/DOM dependency).
 */

/** Custom show: subset of slides in user-specified order (original indexes). App persists it per document to localStorage. */
export interface CustomShow {
  id: string
  name: string
  slideIndices: number[]
}

/**
 * Compute the playback sequence (array of original indexes).
 * - Default: all slides in order, skipping hidden ones (starting from a hidden slide still plays it)
 * - Non-empty customOrder: play in its order (out-of-range slides filtered; hidden slides still skipped, except the start slide)
 * - Fallback: when the result is empty, at least play the start slide
 */
export function computePlayOrder(
  slides: ReadonlyArray<{ hidden?: boolean }>,
  startAt: number,
  customOrder?: readonly number[],
): number[] {
  const playable = (i: number) => slides[i] != null && (!slides[i]!.hidden || i === startAt)
  const o =
    customOrder && customOrder.length > 0
      ? customOrder.filter(playable)
      : slides.map((_, i) => i).filter(playable)
  return o.length > 0 ? [...o] : [startAt]
}

// ── Rehearsal timing ─────────────────────────────────────────────────────────────

/** Rehearsal timing state: perPageMs accumulates dwell milliseconds by original slide index. */
export interface RehearseTiming {
  perPageMs: number[]
  /** Slide currently dwelt on (original index; -1 = finished) */
  currentIndex: number
  /** Timestamp of entering the current slide (ms) */
  enteredAt: number
}

/** Start rehearsal: begin timing from startIndex. */
export function startRehearse(slideCount: number, startIndex: number, now: number): RehearseTiming {
  return {
    perPageMs: new Array(Math.max(0, slideCount)).fill(0),
    currentIndex: startIndex,
    enteredAt: now,
  }
}

/** Page turn: accumulate the current slide's dwell into perPageMs, then switch to nextIndex and restart timing (revisiting a slide keeps accumulating). */
export function switchRehearsePage(
  t: RehearseTiming,
  nextIndex: number,
  now: number,
): RehearseTiming {
  const perPageMs = t.perPageMs.slice()
  if (t.currentIndex >= 0 && t.currentIndex < perPageMs.length) {
    perPageMs[t.currentIndex]! += Math.max(0, now - t.enteredAt)
  }
  return { perPageMs, currentIndex: nextIndex, enteredAt: now }
}

/** End rehearsal: accumulate the last slide's dwell, then convert to seconds per slide (rounded; visited slides count at least 1 second). */
export function finishRehearse(t: RehearseTiming, now: number): number[] {
  const final = switchRehearsePage(t, -1, now)
  return final.perPageMs.map((ms) => (ms > 0 ? Math.max(1, Math.round(ms / 1000)) : 0))
}

/** m:ss clock display (rehearsal timer bar / save confirmation dialog). */
export function formatClock(ms: number): string {
  const sec = Math.max(0, Math.floor(ms / 1000))
  return `${Math.floor(sec / 60)}:${String(sec % 60).padStart(2, '0')}`
}

// ── Show keyboard (PowerPoint's slide-show shortcuts) ────────────────────────────

/** Screen cover: B / . blanks to black, W / , to white; the same key (or navigating) brings the slide back. */
export type ShowScreen = 'none' | 'black' | 'white'

export type ShowKeyCommand = 'next' | 'prev' | 'first' | 'last' | 'exit' | 'black' | 'white'

export interface ShowKeyResult {
  command: ShowKeyCommand | null
  /** 1-based slide number typed before Enter (PowerPoint: digits, then Enter, jumps to that slide) */
  goto: number | null
  /** digits typed so far */
  buffer: string
  /** the key was a show key, so the caller must preventDefault */
  handled: boolean
}

const MAX_GOTO_DIGITS = 4

/**
 * Map one keydown to a show command, following PowerPoint:
 * next = N / Enter / Space / → / ↓ / PageDown; previous = P / Backspace / ← / ↑ / PageUp;
 * Home / End = first / last; Esc = end show; B / . black; W / , white;
 * typing a number then Enter goes to that slide (Backspace edits the number, Esc cancels it).
 * Keys held with ⌘ / Ctrl / Alt are left alone.
 */
export function showKeyCommand(
  key: string,
  buffer: string,
  mods: { meta?: boolean; ctrl?: boolean; alt?: boolean } = {},
): ShowKeyResult {
  const none: ShowKeyResult = { command: null, goto: null, buffer, handled: false }
  if (mods.meta || mods.ctrl || mods.alt) return none
  const cmd = (command: ShowKeyCommand): ShowKeyResult => ({
    command,
    goto: null,
    buffer: '',
    handled: true,
  })
  if (/^[0-9]$/.test(key)) {
    return {
      command: null,
      goto: null,
      buffer: (buffer + key).slice(-MAX_GOTO_DIGITS),
      handled: true,
    }
  }
  switch (key) {
    case 'Enter':
      return buffer
        ? { command: null, goto: Number(buffer), buffer: '', handled: true }
        : cmd('next')
    case 'Backspace':
      return buffer
        ? { command: null, goto: null, buffer: buffer.slice(0, -1), handled: true }
        : cmd('prev')
    case 'Escape':
      return buffer ? { command: null, goto: null, buffer: '', handled: true } : cmd('exit')
    case 'n':
    case 'N':
    case ' ':
    case 'ArrowRight':
    case 'ArrowDown':
    case 'PageDown':
      return cmd('next')
    case 'p':
    case 'P':
    case 'ArrowLeft':
    case 'ArrowUp':
    case 'PageUp':
      return cmd('prev')
    case 'Home':
      return cmd('first')
    case 'End':
      return cmd('last')
    case 'b':
    case 'B':
    case '.':
      return cmd('black')
    case 'w':
    case 'W':
    case ',':
      return cmd('white')
    default:
      return none
  }
}

/** Toggle a screen cover: pressing the same key again restores the slide. */
export function toggleShowScreen(current: ShowScreen, pressed: 'black' | 'white'): ShowScreen {
  return current === pressed ? 'none' : pressed
}
