import { useLayoutEffect, type RefObject } from 'react'

/**
 * Adaptive ribbon labels. Buttons carry optional text labels tagged with a tier
 * (`<span class="rb-lbl" data-tier="1">…</span>`); the ribbon body gets a
 * `data-labels="N"` attribute and CSS (ribbon-fit.css) shows every label whose
 * tier is <= N. N is the highest level at which the tab's content still fits the
 * available width, so a wide window shows text next to the icons and a narrow one
 * falls back to icon-only, for every tab and every UI language (German labels are
 * longer than Chinese ones, so fixed pixel breakpoints would be wrong).
 */
export const RIBBON_LABEL_MAX_LEVEL = 3

/** the slice of an element the fit needs; lets the algorithm be tested without a DOM */
export interface RibbonFitTarget {
  dataset: DOMStringMap
  scrollWidth: number
  clientWidth: number
}

/**
 * Pick the richest label level whose content fits without horizontal overflow, apply it
 * to `el.dataset.labels` and return it. Level 0 (icons only) is the fallback even when
 * it still overflows: the ribbon then scrolls, exactly as it did before labels existed.
 */
export function fitRibbonLabels(
  el: RibbonFitTarget,
  maxLevel: number = RIBBON_LABEL_MAX_LEVEL,
): number {
  for (let level = maxLevel; level > 0; level--) {
    el.dataset.labels = String(level)
    // reading scrollWidth forces layout, so it reflects the level just applied
    if (el.scrollWidth <= el.clientWidth + 1) return level
  }
  el.dataset.labels = '0'
  return 0
}

/**
 * A short visible caption from a tooltip string. Tips read "Cut (⌘X)" or
 * "Format Painter: click here, then select the text": keep the name, drop the
 * shortcut and the explanation. Reusing the tips means every UI language already
 * has a reviewed wording, with no separate label table to keep in sync.
 */
export function labelFromTip(tip: string): string {
  const beforeColon = tip.split(/[:：]/)[0] ?? tip
  return beforeColon.replace(/\s*[(（][^)）]*[)）]\s*$/, '').trim()
}

/**
 * Keep `data-labels` on the ribbon body in step with its width and content. Re-fits when
 * the body is resized and when its children change (tab switch, context tabs appearing,
 * fonts loading), coalesced to one pass per frame.
 */
export function useRibbonLabelFit(ref: RefObject<HTMLElement | null>): void {
  useLayoutEffect(() => {
    const el = ref.current
    if (!el) return
    let frame = 0
    const run = () => {
      frame = 0
      fitRibbonLabels(el)
    }
    const schedule = () => {
      if (frame === 0) frame = requestAnimationFrame(run)
    }
    fitRibbonLabels(el)
    // no observers (jsdom, very old webviews): the one-time fit above still applies
    if (typeof ResizeObserver === 'undefined' || typeof MutationObserver === 'undefined') return
    const resize = new ResizeObserver(schedule)
    resize.observe(el)
    // childList only: fitting writes an attribute, which must not retrigger itself
    const mutation = new MutationObserver(schedule)
    mutation.observe(el, { childList: true, subtree: true })
    return () => {
      resize.disconnect()
      mutation.disconnect()
      if (frame !== 0) cancelAnimationFrame(frame)
    }
  }, [ref])
}
