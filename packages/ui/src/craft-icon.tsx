import type { CSSProperties } from 'react'

/**
 * Renders an icon from one of the Craft sets (craft-icons/*.ts): the same 20 / 24 / 16 unit
 * design grids the sets were drawn on, with the stroke pinned to a painted pixel width like the
 * rest of the ribbon (a proportional stroke would be hairline when small and heavy when large).
 */
export interface CraftSet {
  /** Side of the square design grid */
  readonly grid: number
  readonly icons: Readonly<Record<string, string>>
}

export interface CraftIconProps {
  size?: number
  /** Renders the glyph in one muted ink instead of its accent colors (disabled buttons) */
  muted?: boolean
}

const htmlByMarkup = new Map<string, { __html: string }>()

/**
 * `dangerouslySetInnerHTML` object with a stable identity per markup string. React 19 compares the
 * prop by identity and rewrites the element's children whenever it changes, so a fresh `{ __html }`
 * on every render replaced every icon's nodes each time a parent re-rendered — a click whose
 * mousedown landed on a glyph (e.g. the Paste Options chip, which takes focus from the editor)
 * lost its target before mouseup and never fired.
 */
export function stableHtml(markup: string): { __html: string } {
  let html = htmlByMarkup.get(markup)
  if (!html) {
    html = { __html: markup }
    htmlByMarkup.set(markup, html)
  }
  return html
}

/** Painted stroke width in px, the same pinning the other ribbon glyphs use */
function paintedStroke(size: number): number {
  return size >= 20 ? 1.5 : size >= 13 ? 1.25 : 1.1
}

export function CraftIcon({
  set,
  name,
  size = 20,
  muted,
}: CraftIconProps & { set: CraftSet; name: string }) {
  const markup = set.icons[name]
  const style = {
    // 1em = one stroke width in design units (the generated markup strokes are in em)
    fontSize: `${(paintedStroke(size) * set.grid) / size}px`,
    '--craft-accent': 'var(--accent)',
    '--craft-paper': 'var(--surface)',
    ...(muted ? { filter: 'grayscale(1)', opacity: 0.55 } : {}),
  } as CSSProperties
  return (
    <svg
      className="craft-icon"
      width={size}
      height={size}
      viewBox={`0 0 ${set.grid} ${set.grid}`}
      style={style}
      aria-hidden
      dangerouslySetInnerHTML={stableHtml(markup ?? '')}
    />
  )
}
