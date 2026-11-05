import { useEffect, useState, type RefObject } from 'react'

export interface OutlineAnchor {
  id: string
  /** Pixel offset in the document scroller, independent of its current scroll. */
  top: number
}

/** Anchors must be sorted by top. Equal destinations keep the first (parent) entry. */
export function activeOutlineId(anchors: readonly OutlineAnchor[], top: number): string | null {
  if (anchors.length === 0) return null
  let low = 0
  let high = anchors.length
  while (low < high) {
    const mid = (low + high) >>> 1
    if (anchors[mid]!.top <= top) low = mid + 1
    else high = mid
  }
  let index = Math.max(0, low - 1)
  while (index > 0 && anchors[index - 1]!.top === anchors[index]!.top) index--
  return anchors[index]!.id
}

/** Read layout only when it changes; scroll events do one binary search per frame. */
export function observeOutlineScroll(
  scroller: HTMLElement,
  content: HTMLElement | null,
  getAnchors: () => OutlineAnchor[],
  onChange: (id: string | null) => void,
): () => void {
  let frame = 0
  let dirty = true
  let measureTimer = 0
  let stopped = false
  let anchors: OutlineAnchor[] = []
  const update = () => {
    frame = 0
    if (stopped) return
    if (dirty) {
      anchors = getAnchors()
        .filter((anchor) => Number.isFinite(anchor.top))
        .sort((a, b) => a.top - b.top)
      dirty = false
    }
    // The reading line is just below the viewport top, matching click navigation.
    let top = scroller.scrollTop + Math.min(32, scroller.clientHeight * 0.15)
    if (
      scroller.scrollTop > 0 &&
      scroller.scrollTop + scroller.clientHeight >= scroller.scrollHeight - 2
    ) {
      top = Math.max(top, scroller.scrollTop + scroller.clientHeight - 1)
    }
    onChange(activeOutlineId(anchors, top))
  }
  const schedule = () => {
    if (!frame && !stopped) frame = requestAnimationFrame(update)
  }
  const invalidate = () => {
    clearTimeout(measureTimer)
    measureTimer = window.setTimeout(() => {
      dirty = true
      schedule()
    }, 80)
  }
  const observer = new ResizeObserver(invalidate)
  observer.observe(scroller)
  if (content) observer.observe(content)
  scroller.addEventListener('scroll', schedule, { passive: true })
  scroller.addEventListener('load', invalidate, true)
  document.fonts?.addEventListener('loadingdone', invalidate)
  schedule()
  return () => {
    stopped = true
    if (frame) cancelAnimationFrame(frame)
    clearTimeout(measureTimer)
    observer.disconnect()
    scroller.removeEventListener('scroll', schedule)
    scroller.removeEventListener('load', invalidate, true)
    document.fonts?.removeEventListener('loadingdone', invalidate)
  }
}

export function useOutlineScrollSpy(
  scrollRef: RefObject<HTMLElement | null>,
  content: HTMLElement | null,
  getAnchors: () => OutlineAnchor[],
): string | null {
  const [activeId, setActiveId] = useState<string | null>(null)
  useEffect(() => {
    const scroller = scrollRef.current
    if (!scroller) return
    return observeOutlineScroll(scroller, content, getAnchors, setActiveId)
  }, [scrollRef, content, getAnchors])
  return activeId
}
