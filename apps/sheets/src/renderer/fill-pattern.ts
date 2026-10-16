/// Format Cells → Fill → Pattern Color / Pattern Style. A pattern fill is two native colors plus
/// a pattern type (`<patternFill patternType fg bg>`). Univer paints one flat color per cell, so
/// the cell keeps the blended color for display and carries the editable parts in its `custom`
/// bag; the same mutation that sets the background writes the marker, so Undo/Redo move them
/// together, and the journal turns the marker into the neutral `fill` spec the XLSX writer emits.

import type { PatternType } from '@genoffice/xlsx-gateway/domain/style-color'

export const FILL_PATTERN_KEY = 'genofficeFillPattern'

export interface CellFillPattern {
  readonly pattern: Exclude<PatternType, 'solid'>
  /// Pattern color (`fgColor`), #RRGGBB.
  readonly fg: string
  /// Background color (`bgColor`), #RRGGBB.
  readonly bg: string
}

/// The 18 patterns of Excel's Pattern Style gallery in its row order (solid is the plain fill).
export const FILL_PATTERN_TYPES: readonly Exclude<PatternType, 'solid'>[] = [
  'darkHorizontal',
  'darkVertical',
  'darkDown',
  'darkUp',
  'darkGrid',
  'darkTrellis',
  'lightHorizontal',
  'lightVertical',
  'lightDown',
  'lightUp',
  'lightGrid',
  'lightTrellis',
  'gray125',
  'gray0625',
  'lightGray',
  'mediumGray',
  'darkGray',
]

const HEX = /^#?([0-9a-fA-F]{6})(?:[0-9a-fA-F]{2})?$/

export function toHex(color: string | null | undefined): string | null {
  const match = color ? HEX.exec(color.trim()) : null
  return match ? `#${match[1]!.toUpperCase()}` : null
}

/// Per-channel average, the same blend the file importer uses for the flat display color.
export function blendPatternColors(fg: string, bg: string): string {
  const first = toHex(fg) ?? '#000000'
  const second = toHex(bg) ?? '#FFFFFF'
  let out = '#'
  for (let at = 1; at < 7; at += 2) {
    const a = Number.parseInt(first.slice(at, at + 2), 16)
    const b = Number.parseInt(second.slice(at, at + 2), 16)
    out += Math.floor((a + b) / 2)
      .toString(16)
      .padStart(2, '0')
      .toUpperCase()
  }
  return out
}

export function isFillPattern(value: unknown): value is CellFillPattern {
  if (typeof value !== 'object' || value === null) return false
  const { pattern, fg, bg } = value as Record<string, unknown>
  return (
    typeof pattern === 'string' &&
    (FILL_PATTERN_TYPES as readonly string[]).includes(pattern) &&
    toHex(typeof fg === 'string' ? fg : null) !== null &&
    toHex(typeof bg === 'string' ? bg : null) !== null
  )
}

/// A marker is trusted only while the cell still shows its blended color; a later plain fill
/// (which leaves the marker behind) invalidates it.
export function fillPatternOf(
  custom: Record<string, unknown> | null | undefined | void,
  displayedBackground: string | null | undefined,
): CellFillPattern | null {
  const marker = custom ? custom[FILL_PATTERN_KEY] : undefined
  if (!isFillPattern(marker)) return null
  const shown = toHex(displayedBackground)
  return shown !== null && shown === blendPatternColors(marker.fg, marker.bg) ? marker : null
}

/// The sidecar's `fillPattern` for an imported cell xf → the marker (theme-resolved colors only).
export function markerFromFile(
  fill:
    | { patternType: string; foreground?: string | undefined; background?: string | undefined }
    | undefined,
): CellFillPattern | null {
  if (!fill) return null
  const candidate = {
    pattern: fill.patternType,
    fg: fill.foreground ?? null,
    // without a bgColor the importer shows the pattern color alone, so there is nothing to
    // reproduce: such a fill keeps its native form and is not offered for editing
    bg: fill.background ?? null,
  }
  return isFillPattern(candidate)
    ? { pattern: candidate.pattern, fg: toHex(candidate.fg)!, bg: toHex(candidate.bg)! }
    : null
}

/// Neutral style spec the XLSX writer consumes (`WorkbookStyleEdit.fill`).
export function fillSpecOf(fill: CellFillPattern): {
  pattern: PatternType
  fg: string
  bg: string
} {
  return { pattern: fill.pattern, fg: fill.fg, bg: fill.bg }
}

/// CSS that approximates a pattern for the dialog's swatches and the sample (display only).
export function patternPreviewCss(fill: CellFillPattern): {
  backgroundColor: string
  backgroundImage: string
  backgroundSize: string
} {
  const ink = fill.fg
  const dense = fill.pattern.startsWith('dark')
  const width = dense ? 2 : 1
  const gap = dense ? 4 : 6
  const line = (angle: number): string =>
    `repeating-linear-gradient(${angle}deg, ${ink} 0 ${width}px, transparent ${width}px ${gap}px)`
  const dots = (alpha: number): string =>
    `radial-gradient(${ink} ${alpha}px, transparent ${alpha + 0.5}px)`
  const image: Record<CellFillPattern['pattern'], string> = {
    darkHorizontal: line(0),
    lightHorizontal: line(0),
    darkVertical: line(90),
    lightVertical: line(90),
    darkDown: line(45),
    lightDown: line(45),
    darkUp: line(135),
    lightUp: line(135),
    darkGrid: `${line(0)}, ${line(90)}`,
    lightGrid: `${line(0)}, ${line(90)}`,
    darkTrellis: `${line(45)}, ${line(135)}`,
    lightTrellis: `${line(45)}, ${line(135)}`,
    gray125: dots(0.6),
    gray0625: dots(0.4),
    lightGray: dots(0.9),
    mediumGray: dots(1.3),
    darkGray: dots(1.7),
  }
  return {
    backgroundColor: fill.bg,
    backgroundImage: image[fill.pattern],
    backgroundSize: ['gray125', 'gray0625', 'lightGray', 'mediumGray', 'darkGray'].includes(
      fill.pattern,
    )
      ? '4px 4px'
      : 'auto',
  }
}

/// The flat color a pattern spec shows on the grid: the blend for a two-color pattern, else null
/// (callers fall back to the foreground / first gradient stop).
export function blendedDisplayOf(fill: unknown): string | null {
  if (typeof fill !== 'object' || fill === null || !('pattern' in fill)) return null
  const { pattern, fg, bg } = fill as { pattern: unknown; fg: unknown; bg?: unknown }
  if (pattern === 'solid' || typeof fg !== 'string' || typeof bg !== 'string') return null
  return blendPatternColors(fg, bg)
}
