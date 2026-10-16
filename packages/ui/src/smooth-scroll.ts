/**
 * Eased wheel scrolling for every renderer ("Smooth scrolling" in Settings), modelled on how Office
 * scrolls: the content starts moving on the very first frame and glides to rest with an ease-out.
 *
 * Every wheel / trackpad delta is accumulated into a target and the content chases it with an
 * exponential ease. The time constant adapts to the input: isolated notches (a mouse wheel) get a
 * longer, more visible glide; a rapid stream of small deltas (a trackpad, or a wheel spun fast)
 * gets a short one so the scrolling stays under the finger instead of lagging behind it.
 * Zoom gestures (⌘/Ctrl + wheel) and fields' own scrolling keep the browser default.
 *
 * Two targets: a DOM scroller gets its scrollTop/scrollLeft animated; a canvas surface (Univer's
 * grid, the slide stage) has no scroller, so the smoothed deltas are re-dispatched to it as small
 * wheel events and the app's own wheel handling does the scrolling.
 */

/** Glide time constant for an isolated notch (ms): ~95% of the distance in 3 × this. */
export const SMOOTH_TAU_MS = 80
/** Time constant while deltas stream in rapidly (trackpad / fast spin). */
export const SMOOTH_STREAM_TAU_MS = 40
/** A gap shorter than this between wheel events counts as a continuous stream. */
export const STREAM_GAP_MS = 30
const FIRST_FRAME_MS = 16
const STOP_PX = 0.5
const LINE_PX = 40

export interface WheelLike {
  readonly deltaX: number
  readonly deltaY: number
  readonly deltaMode: number
  readonly ctrlKey: boolean
  readonly metaKey: boolean
  readonly wheelDeltaY?: number
}

/** True for a notched mouse wheel (line-based, or a pixel delta that is a multiple of 120). */
export function isNotchedWheel(event: WheelLike): boolean {
  if (event.ctrlKey || event.metaKey) return false
  if (event.deltaMode !== 0) return true
  const wheel = event.wheelDeltaY ?? 0
  return wheel !== 0 && Math.abs(wheel) % 120 === 0 && Number.isInteger(event.deltaY)
}

/** Pixel size of a wheel event on one axis. */
export function wheelPixels(delta: number, deltaMode: number, pageSize: number): number {
  if (deltaMode === 1) return delta * LINE_PX
  if (deltaMode === 2) return delta * pageSize
  return delta
}

/** Time constant for a wheel event given the gap since the previous one. */
export function tauForGap(gapMs: number): number {
  return gapMs < STREAM_GAP_MS ? SMOOTH_STREAM_TAU_MS : SMOOTH_TAU_MS
}

/** Wheel events the smoother leaves alone: zoom gestures and horizontal-less zero deltas. */
export function isSmoothable(event: WheelLike): boolean {
  if (event.ctrlKey || event.metaKey) return false
  return event.deltaX !== 0 || event.deltaY !== 0
}

/** How far to move this frame: exponential ease toward the remaining distance. */
export function easeStep(remaining: number, dtMs: number, tauMs: number = SMOOTH_TAU_MS): number {
  const fraction = 1 - Math.exp(-Math.max(dtMs, 0) / tauMs)
  const step = remaining * fraction
  return Math.abs(remaining) <= STOP_PX ? remaining : step
}

function canScroll(el: Element, axis: 'x' | 'y', direction: number): boolean {
  const style = getComputedStyle(el)
  const overflow = axis === 'y' ? style.overflowY : style.overflowX
  if (overflow !== 'auto' && overflow !== 'scroll' && overflow !== 'overlay') return false
  const size = axis === 'y' ? el.clientHeight : el.clientWidth
  const total = axis === 'y' ? el.scrollHeight : el.scrollWidth
  const pos = axis === 'y' ? el.scrollTop : el.scrollLeft
  if (total <= size + 1) return false
  return direction > 0 ? pos + size < total - 1 : pos > 0
}

function scrollerFor(
  start: EventTarget | null,
  axis: 'x' | 'y',
  direction: number,
): Element | null {
  let el = start instanceof Element ? start : null
  while (el) {
    if (canScroll(el, axis, direction)) return el
    el = el.parentElement
  }
  return null
}

interface DomGlide {
  el: Element
  tau: number
  targetY: number
  targetX: number
  raf: number
  last: number
}

let installed = false
let enabled = false
let redispatching = false
let lastWheelAt = -Infinity

/** Turns the behaviour on or off; installs the single window listener on first use. */
export function setSmoothScroll(on: boolean): void {
  enabled = on
  if (installed || typeof window === 'undefined') return
  installed = true
  const glides = new WeakMap<Element, DomGlide>()
  let canvasGlide: {
    target: EventTarget
    x: number
    y: number
    remY: number
    remX: number
    tau: number
    raf: number
    last: number
  } | null = null

  const runDom = (g: DomGlide, now: number): void => {
    // the frame timestamp can predate the glide's own start: wait for a frame that moves time on
    if (now <= g.last) {
      g.raf = requestAnimationFrame((t) => runDom(g, t))
      return
    }
    const dt = now - g.last
    g.last = now
    const stepY = easeStep(g.targetY - g.el.scrollTop, dt, g.tau)
    const stepX = easeStep(g.targetX - g.el.scrollLeft, dt, g.tau)
    g.el.scrollTop += stepY
    g.el.scrollLeft += stepX
    const doneY = Math.abs(g.targetY - g.el.scrollTop) <= STOP_PX
    const doneX = Math.abs(g.targetX - g.el.scrollLeft) <= STOP_PX
    // a scroller that stopped short (hit its end) must not spin forever
    const stuck = stepY === 0 && stepX === 0
    if ((doneY && doneX) || stuck) {
      glides.delete(g.el)
      return
    }
    g.raf = requestAnimationFrame((t) => runDom(g, t))
  }

  const runCanvas = (now: number): void => {
    const g = canvasGlide
    if (!g) return
    if (now <= g.last) {
      g.raf = requestAnimationFrame(runCanvas)
      return
    }
    const dt = now - g.last
    g.last = now
    const stepY = easeStep(g.remY, dt, g.tau)
    const stepX = easeStep(g.remX, dt, g.tau)
    g.remY -= stepY
    g.remX -= stepX
    redispatching = true
    try {
      g.target.dispatchEvent(
        new WheelEvent('wheel', {
          deltaX: stepX,
          deltaY: stepY,
          deltaMode: 0,
          clientX: g.x,
          clientY: g.y,
          bubbles: true,
          cancelable: true,
          composed: true,
        }),
      )
    } finally {
      redispatching = false
    }
    if (Math.abs(g.remY) <= STOP_PX && Math.abs(g.remX) <= STOP_PX) {
      canvasGlide = null
      return
    }
    g.raf = requestAnimationFrame(runCanvas)
  }

  window.addEventListener(
    'wheel',
    (event) => {
      if (!enabled || redispatching || event.defaultPrevented || !event.isTrusted) return
      if (!isSmoothable(event)) return
      const gap = event.timeStamp - lastWheelAt
      lastWheelAt = event.timeStamp
      const tau = tauForGap(gap)
      const target = event.target
      if (!(target instanceof Element)) return
      // fields and selects scroll themselves; leave them alone
      if (target.closest('select, input, textarea')) return
      const pageH = window.innerHeight
      const dy = wheelPixels(event.deltaY, event.deltaMode, pageH)
      const dx = wheelPixels(event.deltaX, event.deltaMode, window.innerWidth)
      const axis: 'x' | 'y' = Math.abs(dy) >= Math.abs(dx) ? 'y' : 'x'
      const direction = axis === 'y' ? dy : dx
      if (direction === 0) return

      const scroller = scrollerFor(target, axis, direction)
      if (scroller) {
        event.preventDefault()
        let g = glides.get(scroller)
        if (!g) {
          g = {
            el: scroller,
            tau,
            targetY: scroller.scrollTop,
            targetX: scroller.scrollLeft,
            raf: 0,
            // start as if one frame has already elapsed: the first frame already moves
            last: performance.now() - FIRST_FRAME_MS,
          }
          glides.set(scroller, g)
          g.raf = requestAnimationFrame((t) => runDom(g!, t))
        }
        const maxY = scroller.scrollHeight - scroller.clientHeight
        const maxX = scroller.scrollWidth - scroller.clientWidth
        g.tau = tau
        g.targetY = Math.min(Math.max(g.targetY + dy, 0), Math.max(maxY, 0))
        g.targetX = Math.min(Math.max(g.targetX + dx, 0), Math.max(maxX, 0))
        return
      }
      // no DOM scroller: a canvas surface that scrolls from its own wheel handler
      if (target instanceof HTMLCanvasElement || target.closest('canvas')) {
        event.preventDefault()
        event.stopPropagation()
        if (!canvasGlide || canvasGlide.target !== target) {
          if (canvasGlide) cancelAnimationFrame(canvasGlide.raf)
          canvasGlide = {
            target,
            x: event.clientX,
            y: event.clientY,
            remY: 0,
            remX: 0,
            tau,
            raf: 0,
            last: performance.now() - FIRST_FRAME_MS,
          }
          canvasGlide.raf = requestAnimationFrame(runCanvas)
        }
        canvasGlide.x = event.clientX
        canvasGlide.y = event.clientY
        canvasGlide.tau = tau
        canvasGlide.remY += dy
        canvasGlide.remX += dx
      }
    },
    { capture: true, passive: false },
  )
}
