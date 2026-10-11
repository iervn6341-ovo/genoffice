import type { ReactElement, ReactNode } from 'react'
import { CraftIcon, type CraftSet } from '@genoffice/ui'
import { GRID_ICONS } from '@genoffice/ui/craft-icons/grid'

/// Glyph icons drawn like the GridCraft icon set (packages/ui/src/craft-icons/NOTICE-craft-icons.md)
const CRAFT: CraftSet = { grid: 16, icons: GRID_ICONS }

/// Ribbon icon set drawn to the shared icon standard: 24×24 canvas with
/// 1.5-unit strokes, round caps and round joins. The rendered size comes
/// from CSS, which also pins the PAINTED stroke to ~1.5px (20px+ glyphs) /
/// ~1.25px (13-16px glyphs) via stroke-width overrides in styles.css —
/// proportional scaling would leave small glyphs hairline and big ones fat.
function Icon({ children }: { readonly children: ReactNode }): ReactElement {
  return (
    <svg
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth={1.5}
      strokeLinecap="round"
      strokeLinejoin="round"
      aria-hidden="true"
    >
      {children}
    </svg>
  )
}

/// Thin dropdown chevron, copied from the slides ribbon's
/// RbCaret (replaces the ▾ text glyph).
export function CaretIcon(): ReactElement {
  return (
    <svg className="chev" width="10" height="10" viewBox="0 0 24 24" fill="none" aria-hidden="true">
      <path
        d="M5.5 9.25 12 15.75l6.5-6.5"
        stroke="currentColor"
        strokeWidth="1.5"
        strokeLinecap="round"
        strokeLinejoin="round"
      />
    </svg>
  )
}

/// Border-menu glyphs (16-canvas, docs-ribbon conventions): a faint dashed
/// frame marks the cell block, solid strokes mark the edges the entry sets.
function BorderSvg({ children }: { children: ReactNode }): ReactElement {
  return (
    <svg
      className="border-icon"
      width="16"
      height="16"
      viewBox="0 0 16 16"
      fill="none"
      stroke="currentColor"
      strokeWidth="1.3"
      strokeLinecap="round"
      strokeLinejoin="round"
      aria-hidden="true"
    >
      {children}
    </svg>
  )
}

function BorderDashedFrame(): ReactElement {
  return (
    <rect
      x="3.02"
      y="3.02"
      width="9.96"
      height="9.96"
      rx="0.42"
      strokeWidth="1"
      strokeDasharray="1.5 1.7"
      opacity="0.55"
    />
  )
}

export function BorderAllIcon(): ReactElement {
  return (
    <BorderSvg>
      <rect x="3.02" y="3.02" width="9.96" height="9.96" rx="0.42" />
      <path d="M 3.02 8 h 9.96 M 8 3.02 v 9.96" />
    </BorderSvg>
  )
}

export function BorderOuterIcon(): ReactElement {
  return (
    <BorderSvg>
      <rect x="3.02" y="3.02" width="9.96" height="9.96" rx="0.42" />
      <path
        d="M 3.02 8 h 9.96 M 8 3.02 v 9.96"
        strokeWidth="1"
        strokeDasharray="1.5 1.7"
        opacity="0.55"
      />
    </BorderSvg>
  )
}

export function BorderThickOuterIcon(): ReactElement {
  return (
    <BorderSvg>
      <rect x="3.22" y="3.22" width="9.56" height="9.56" rx="0.42" strokeWidth="2.1" />
      <path
        d="M 3.02 8 h 9.96 M 8 3.02 v 9.96"
        strokeWidth="1"
        strokeDasharray="1.5 1.7"
        opacity="0.55"
      />
    </BorderSvg>
  )
}

export function BorderTopIcon(): ReactElement {
  return (
    <BorderSvg>
      <BorderDashedFrame />
      <path d="M 3.02 3.02 h 9.96" />
    </BorderSvg>
  )
}

export function BorderBottomIcon(): ReactElement {
  return (
    <BorderSvg>
      <BorderDashedFrame />
      <path d="M 3.02 12.98 h 9.96" />
    </BorderSvg>
  )
}

export function BorderLeftIcon(): ReactElement {
  return (
    <BorderSvg>
      <BorderDashedFrame />
      <path d="M 3.02 3.02 v 9.96" />
    </BorderSvg>
  )
}

export function BorderRightIcon(): ReactElement {
  return (
    <BorderSvg>
      <BorderDashedFrame />
      <path d="M 12.98 3.02 v 9.96" />
    </BorderSvg>
  )
}

export function BorderNoneIcon(): ReactElement {
  return (
    <BorderSvg>
      <BorderDashedFrame />
      <path
        d="M 3.02 8 h 9.96 M 8 3.02 v 9.96"
        strokeWidth="1"
        strokeDasharray="1.5 1.7"
        opacity="0.55"
      />
    </BorderSvg>
  )
}

/// Quick-access save button — standard floppy glyph shared across all apps.
export function SaveIcon(): ReactElement {
  return (
    <Icon>
      <path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z" />
      <path d="M17 21v-8H7v8" />
      <path d="M7 3v5h8V3" />
    </Icon>
  )
}

/// Save As keeps the floppy silhouette and adds the familiar pencil overlay.
export function SaveAsIcon(): ReactElement {
  return (
    <Icon>
      <path d="M12.25 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v1.25" />
      <path d="M7 21v-8h5" />
      <path d="M7 3v5h8V3" />
      <path d="m13.25 19.5.75-3.25 5.75-5.75 2.5 2.5-5.75 5.75-3.25.75Z" />
      <path d="m18.75 11.5 2.5 2.5" />
    </Icon>
  )
}

/// Quick-access undo/redo — design-supplied geometry shared with the docs
/// and slides ribbons, so the three apps' tab-row arrows read identically.
export function UndoIcon(): ReactElement {
  return (
    <Icon>
      <path d="M5.91026 4L2.5 7.14791L5.91026 10.8205" />
      <path d="M3.96154 7.41028H15.1636C18.5169 7.41028 21.3646 10.1484 21.4953 13.5C21.6334 17.0416 18.707 20.0769 15.1636 20.0769H6.88384" />
    </Icon>
  )
}

export function RedoIcon(): ReactElement {
  return (
    <Icon>
      <path d="M18.0897 4L21.5 7.14791L18.0897 10.8205" />
      <path d="M20.0385 7.41028H8.83636C5.4831 7.41028 2.63537 10.1484 2.5047 13.5C2.36657 17.0416 5.29296 20.0769 8.83636 20.0769H17.1162" />
    </Icon>
  )
}

// drawings shared by several glyph keys
const SPARKLE_ICON = (
  <Icon>
    <path d="M12 4.75 13.9 10.1 19.25 12l-5.35 1.9L12 19.25l-1.9-5.35L4.75 12l5.35-1.9L12 4.75Z" />
  </Icon>
)
const CLOCK_ICON = (
  <Icon>
    <circle cx="12" cy="12" r="7.5" />
    <path d="M12 7.75V12l3 2.25" />
  </Icon>
)
const REFRESH_ICON = (
  <Icon>
    <path d="M18.74 12A6.74 6.74 0 1 1 16.65 7.12" />
    <path d="M16.88 4.1v3.25h-3.25" />
  </Icon>
)

/// Icons keyed by the legacy glyph strings, so ribbon call sites stay
/// unchanged; symbols without an entry render as plain text (letterforms
/// such as $, ?, θ, ƒx are typography, not drawn icons).
export const RIBBON_GLYPH_ICONS: Record<string, ReactElement> = {
  // ---- alignment / rows / columns ----
  '≡': <CraftIcon set={CRAFT} name="AlignLeft" size={20} />,
  '≣': <CraftIcon set={CRAFT} name="AlignCenter" size={20} />,
  '☰': <CraftIcon set={CRAFT} name="AlignRight" size={20} />,
  '⤒': (
    <Icon>
      <path d="M5 4.75h14" />
      <path d="M12 19.25V9.5M8.5 13 12 9.5l3.5 3.5" />
    </Icon>
  ),
  '⤓': (
    <Icon>
      <path d="M5 19.25h14" />
      <path d="M12 4.75v9.75M8.5 11l3.5 3.5 3.5-3.5" />
    </Icon>
  ),
  '↩': <CraftIcon set={CRAFT} name="Wrap" size={20} />,
  '⇤': (
    <Icon>
      <path d="M5 5v14" />
      <path d="M19 12H9.5M13 8.5 9.5 12l3.5 3.5" />
    </Icon>
  ),
  '⇥': (
    <Icon>
      <path d="M19 5v14" />
      <path d="M5 12h9.5M11 8.5 14.5 12 11 15.5" />
    </Icon>
  ),
  // ---- mini chart-type grid ----
  '▮▬': <CraftIcon set={CRAFT} name="ChartBar" size={20} />,
  '📈': <CraftIcon set={CRAFT} name="ChartLine" size={20} />,
  '◪': (
    <Icon>
      <path d="M4.75 19.25V13.5l5.5-6 4 5.5 5-4.25v10.5H4.75Z" />
    </Icon>
  ),
  '◔': <CraftIcon set={CRAFT} name="ChartPie" size={20} />,
  '∴': <CraftIcon set={CRAFT} name="ChartScatter" size={20} />,
  '✳': (
    <Icon>
      <path d="M12 4.25 18.75 8.1v7.8L12 19.75 5.25 15.9V8.1L12 4.25Z" />
      <path d="M12 12V4.25M12 12l6.75 3.9M12 12l-6.75 3.9" />
    </Icon>
  ),
  '◍': (
    <Icon>
      <circle cx="12" cy="12" r="7.25" />
      <circle cx="12" cy="12" r="3" />
    </Icon>
  ),
  '𝄜': (
    <Icon>
      <path d="M4.5 19.5h15" />
      <path d="M7 19.5v-5.5M11 19.5V10" />
      <path d="m5 8.5 4.5-3.25L14 8l5-3.75" />
    </Icon>
  ),
  // ---- clipboard ----
  // paste shares the docs ribbon's clipboard geometry (docs 16-canvas × 1.5);
  // paste special swaps the text lines for option sliders on the same board
  '📋': <CraftIcon set={CRAFT} name="Paste" size={20} />,
  '📑': (
    <Icon>
      <path d="M 14.99 5.02 H 16.98 C 18.09 5.02 18.98 5.92 18.98 7.04 L 18.88 19.02 C 18.87 20.11 17.98 21 16.88 21 H 6.88 C 5.77 21 4.87 20.09 4.88 18.98 L 4.98 7 C 4.98 5.91 5.88 5.02 6.98 5.02 H 8.99" />
      <rect x="9.16" y="3.5" width="5.68" height="2.84" rx="0.95" />
      <path d="M 8 10 H 16 M 8 13.5 H 16 M 8 17 H 13" />
    </Icon>
  ),
  // cut / copy / format painter share one 24-canvas design across all three
  // apps (the docs ribbon carries the same geometry × 2/3 on its 16 canvas)
  '✂': <CraftIcon set={CRAFT} name="Cut" size={20} />,
  '⧉': (
    <Icon>
      <rect x="7" y="7" width="14" height="14" rx="3" />
      <path d="M 14.5 4 H 7 C 5.34 4 4 5.34 4 7 V 14.5" />
    </Icon>
  ),
  '🖌': <CraftIcon set={CRAFT} name="Brush" size={20} />,
  // ---- find / filter ----
  '🔍': (
    <Icon>
      <circle cx="10.5" cy="10.5" r="6" />
      <path d="m15 15 4.5 4.5" />
    </Icon>
  ),
  '▽': <CraftIcon set={CRAFT} name="Filter" size={20} />,
  '⊘': (
    <Icon>
      <path d="M4 5.25h11.5l-4.35 5v5.4l-2.8-1.75v-3.65L4 5.25Z" />
      <path d="m15.75 14.75 4.5 4.5m0-4.5-4.5 4.5" />
    </Icon>
  ),
  '⌖': (
    <Icon>
      <circle cx="12" cy="12" r="5.75" />
      <path d="M12 4.5v2.5M12 17v2.5M4.5 12H7M17 12h2.5" />
    </Icon>
  ),
  // ---- comments / review ----
  '🗨': <CraftIcon set={CRAFT} name="Comment" size={20} />,
  '🗑': (
    <Icon>
      <path d="M 5.48 7.05 h 13.05" />
      <path d="M 9.75 7.05 v -0.9 a 1.35 1.35 0 0 1 1.35 -1.35 h 1.8 a 1.35 1.35 0 0 1 1.35 1.35 v 0.9" />
      <path d="m 6.82 7.05 0.81 11.03 a 1.8 1.8 0 0 0 1.8 1.67 h 5.13 a 1.8 1.8 0 0 0 1.8 -1.66 l 0.81 -11.02" />
      <path d="M 10.2 10.65 v 5.4 M 13.8 10.65 v 5.4" />
    </Icon>
  ),
  '✓': (
    <Icon>
      <path d="m4.75 13.75 4.75 5L19.25 5.75" />
    </Icon>
  ),
  '☑': (
    <Icon>
      <rect x="4.75" y="4.75" width="14.5" height="14.5" rx="2" />
      <path d="m8.5 12.5 2.5 2.5 4.75-5.5" />
    </Icon>
  ),
  '⚠': (
    <Icon>
      <path d="M10.7 5.75 4.6 16.5a1.5 1.5 0 0 0 1.3 2.25h12.2a1.5 1.5 0 0 0 1.3-2.25L13.3 5.75a1.5 1.5 0 0 0-2.6 0Z" />
      <path d="M12 9.75v3.75" />
      <circle cx="12" cy="16.15" r="0.4" fill="currentColor" stroke="none" />
    </Icon>
  ),
  '❔': (
    <Icon>
      <path d="M 9.36 9.27 A 2.64 2.64 0 0 1 12 6.72 c 1.46 0 2.64 1.1 2.64 2.55 0 0.97 -0.53 1.58 -1.32 2.2 -0.79 0.62 -1.32 1.1 -1.32 2.11 v 0.4" />
      <circle cx="12" cy="16.84" r="0.44" fill="currentColor" stroke="none" />
      <circle cx="12" cy="12" r="7.26" />
    </Icon>
  ),
  '💡': (
    <Icon>
      <path d="M12 4.25a5.75 5.75 0 0 0-3.25 10.5c.62.43 1 1.05 1 1.78v.72h4.5v-.72c0-.73.38-1.35 1-1.78A5.75 5.75 0 0 0 12 4.25Z" />
      <path d="M10 19.75h4" />
    </Icon>
  ),
  '👓': (
    <Icon>
      <circle cx="7.25" cy="14.25" r="3.25" />
      <circle cx="16.75" cy="14.25" r="3.25" />
      <path d="M10.5 13.75a1.5 1.5 0 0 1 3 0" />
      <path d="M4 14.25c0-2.1.5-4.5 1.5-6.5M20 14.25c0-2.1-.5-4.5-1.5-6.5" />
    </Icon>
  ),
  '✒': (
    <Icon>
      <path d="M13.75 5.5 18.5 10.25 9.5 19.25H4.75V14.5l9-9Z" />
      <path d="m12 7.25 4.75 4.75" />
    </Icon>
  ),
  // ---- charts / analysis ----
  '📊': <CraftIcon set={CRAFT} name="Chart" size={20} />,
  '🗠': (
    <Icon>
      <path d="M4.5 4.5v15h15" />
      <path d="m7.5 15.5 3.5-4 2.7 2.5 4.3-5.5" />
    </Icon>
  ),
  '〜': <CraftIcon set={CRAFT} name="Sparkline" size={20} />,
  '✦': SPARKLE_ICON,
  '✧': SPARKLE_ICON,
  '∑': <CraftIcon set={CRAFT} name="Subtotal" size={20} />,
  Σ: <CraftIcon set={CRAFT} name="Sum" size={20} />,
  '🧮': <CraftIcon set={CRAFT} name="Calc" size={20} />,
  // ---- grids / tables / layout ----
  '▦': (
    <Icon>
      <rect x="4.5" y="4.5" width="15" height="15" rx="1.5" />
      <path d="M4.5 9.5h15M4.5 14.5h15M9.5 4.5v15M14.5 4.5v15" />
    </Icon>
  ),
  '⊞': <CraftIcon set={CRAFT} name="PivotTable" size={20} />,
  '▤': (
    <Icon>
      <rect x="4.5" y="4.5" width="15" height="15" rx="1.5" />
      <path d="M4.5 9.5h15M4.5 14.5h15" />
    </Icon>
  ),
  '⊟': <CraftIcon set={CRAFT} name="Group" size={20} />,
  '⊕': (
    <Icon>
      <circle cx="12" cy="12" r="7.25" />
      <path d="M12 8.5v7M8.5 12h7" />
    </Icon>
  ),
  '⊙': (
    <Icon>
      <circle cx="12" cy="12" r="7.25" />
      <circle cx="12" cy="12" r="1.9" fill="currentColor" stroke="none" />
    </Icon>
  ),
  '⿴': <CraftIcon set={CRAFT} name="Margins" size={20} />,
  '⤢': <CraftIcon set={CRAFT} name="Orient" size={20} />,
  '▭': <CraftIcon set={CRAFT} name="Paper" size={20} />,
  '⬚': (
    <Icon>
      <rect x="5" y="5" width="14" height="14" rx="1" strokeDasharray="3 2.7" />
    </Icon>
  ),
  '┆': <CraftIcon set={CRAFT} name="PageBreak" size={20} />,
  '❄': <CraftIcon set={CRAFT} name="Freeze" size={20} />,
  // ---- media / objects ----
  '🎨': (
    <Icon>
      <path d="M12 4.5c-4.7 0-8.5 3.4-8.5 7.5s3.8 7.5 8.5 7.5c1.05 0 1.9-.85 1.9-1.9 0-.5-.2-.95-.5-1.3-.3-.35-.5-.8-.5-1.25 0-1.05.85-1.9 1.9-1.9h2.2c2 0 3.5-1.5 3.5-3.4C20.5 7.2 16.7 4.5 12 4.5Z" />
      <circle cx="8.1" cy="10.1" r="0.5" fill="currentColor" stroke="none" />
      <circle cx="11.6" cy="8.1" r="0.5" fill="currentColor" stroke="none" />
      <circle cx="15.4" cy="9.3" r="0.5" fill="currentColor" stroke="none" />
    </Icon>
  ),
  '🖼': <CraftIcon set={CRAFT} name="Picture" size={20} />,
  '◇': <CraftIcon set={CRAFT} name="Shapes" size={20} />,
  '🔗': <CraftIcon set={CRAFT} name="Link" size={20} />,
  '🗎': (
    <Icon>
      <path d="M 13.37 4.49 H 7.68 c -0.76 0 -1.36 0.61 -1.36 1.37 v 12.29 c 0 0.76 0.61 1.37 1.37 1.37 h 8.65 c 0.76 0 1.37 -0.61 1.37 -1.36 V 8.82 L 13.37 4.49 Z" />
      <path d="M 13.37 4.49 V 8.82 h 4.32" />
    </Icon>
  ),
  '🗒': (
    <Icon>
      <path d="M 13.37 4.49 H 7.68 c -0.76 0 -1.36 0.61 -1.36 1.37 v 12.29 c 0 0.76 0.61 1.37 1.37 1.37 h 8.65 c 0.76 0 1.37 -0.61 1.37 -1.36 V 8.82 L 13.37 4.49 Z" />
      <path d="M 13.37 4.49 V 8.82 h 4.32" />
      <path d="M 9.27 12.46 h 5.46 M 9.27 15.19 h 3.64" />
    </Icon>
  ),
  '🕮': (
    <Icon>
      <path d="M12 6.25c-1.8-1.3-4-1.9-7.25-1.75V17.75c3.25-.15 5.45.45 7.25 1.75 1.8-1.3 4-1.9 7.25-1.75V4.5C16 4.35 13.8 4.95 12 6.25Z" />
      <path d="M12 6.25V19.5" />
    </Icon>
  ),
  '🛢': (
    <Icon>
      <ellipse cx="12" cy="6.5" rx="7" ry="2.5" />
      <path d="M 5.49 6.89 v 10.23 c 0 1.3 2.91 2.33 6.51 2.33 s 6.51 -1.02 6.51 -2.33 v -10.23" />
      <path d="M 5.49 12 c 0 1.3 2.91 2.33 6.51 2.33 s 6.51 -1.02 6.51 -2.33" />
    </Icon>
  ),
  '🌐': (
    <Icon>
      <circle cx="12" cy="12" r="7.5" />
      <path d="M12 4.5c2.5 2 3.75 4.5 3.75 7.5S14.5 17.5 12 19.5c-2.5-2-3.75-4.5-3.75-7.5S9.5 6.5 12 4.5Z" />
      <path d="M4.75 12h14.5" />
    </Icon>
  ),
  '🏷': <CraftIcon set={CRAFT} name="Name" size={20} />,
  '🔒': <CraftIcon set={CRAFT} name="Lock" size={20} />,
  '🔐': <CraftIcon set={CRAFT} name="Lock" size={20} />,
  '▶': (
    <Icon>
      <path d="M 7.8 4.5 v 15 L 20.1 12 7.8 4.5 Z" />
    </Icon>
  ),
  '⚡': (
    <Icon>
      <path d="M13 4.25 6.5 13.25h4.5L11 19.75l6.5-9h-4.5L13 4.25Z" />
    </Icon>
  ),
  '♿': (
    <Icon>
      <circle cx="12" cy="6.15" r="1.8" />
      <path d="M 5.48 9.75 c 4.32 1.17 8.73 1.17 13.05 0" />
      <path d="M 12 10.65 v 3.38 l -3.06 4.95 M 12 14.03 l 3.06 4.95" />
    </Icon>
  ),
  '🕘': CLOCK_ICON,
  '🕐': CLOCK_ICON,
  '🕒': CLOCK_ICON,
  '⟳': REFRESH_ICON,
  '↻': REFRESH_ICON,
  '✕': (
    <Icon>
      <path d="m6 6 12 12M18 6 6 18" />
    </Icon>
  ),
  // ---- sort / move ----
  '↓': <CraftIcon set={CRAFT} name="FillDown" size={20} />,
  '↑': (
    <Icon>
      <path d="M8 19V5.5M4.75 8.75 8 5.5l3.25 3.25" />
      <path d="M13.5 9.5h3M13.5 13.5h4.5M13.5 17.5h6" />
    </Icon>
  ),
  '⇄': (
    <Icon>
      <path d="M4.75 8.25h14.5M16 5l3.25 3.25L16 11.5" />
      <path d="M19.25 15.75H4.75M8 12.5l-3.25 3.25L8 19" />
    </Icon>
  ),
  '⇅': <CraftIcon set={CRAFT} name="SortFilter" size={20} />,
  '⇶': <CraftIcon set={CRAFT} name="TextColumns" size={20} />,
  // ---- text ----
  A: (
    <Icon>
      <path d="M5.75 7V4.75h12.5V7M12 4.75v14.5M9.5 19.25h5" />
    </Icon>
  ),
  // ---- windows / panes ----
  '🗔': (
    <Icon>
      <rect x="5.17" y="5.63" width="13.65" height="12.74" rx="1.37" />
      <path d="M 5.17 9.27 h 13.65" />
    </Icon>
  ),
  '◫': (
    <Icon>
      <rect x="4.5" y="5" width="15" height="14" rx="1.5" />
      <path d="M12 5v14" />
    </Icon>
  ),
  '▥': (
    <Icon>
      <rect x="4.5" y="4.5" width="15" height="15" rx="1.5" />
      <path d="M9.5 4.5v15M14.5 4.5v15" />
    </Icon>
  ),
  '⬡': (
    <Icon>
      <path d="M12 4.25 18.75 8.1v7.8L12 19.75 5.25 15.9V8.1L12 4.25Z" />
    </Icon>
  ),
  // ---- editing arrows ----
  '⌫': (
    <Icon>
      <path d="M9.25 5.5h9a1.5 1.5 0 0 1 1.5 1.5v10a1.5 1.5 0 0 1-1.5 1.5h-9L4.25 12l5-6.5Z" />
      <path d="m11.75 9.75 4.5 4.5m0-4.5-4.5 4.5" />
    </Icon>
  ),
  '⇢': (
    <Icon>
      <path d="M4.75 12h11.5" strokeDasharray="2.5 2.5" />
      <path d="M15.5 8.25 19.25 12l-3.75 3.75" />
    </Icon>
  ),
  '⇠': (
    <Icon>
      <path d="M19.25 12H7.75" strokeDasharray="2.5 2.5" />
      <path d="M8.5 8.25 4.75 12l3.75 3.75" />
    </Icon>
  ),
  '→': (
    <Icon>
      <path d="M4.75 12h14.5M15.5 8.25 19.25 12l-3.75 3.75" />
    </Icon>
  ),
  '←': (
    <Icon>
      <path d="M19.25 12H4.75M8.5 8.25 4.75 12l3.75 3.75" />
    </Icon>
  ),
  '↔': (
    <Icon>
      <path d="M4.75 12h14.5M8 8.75 4.75 12 8 15.25M16 8.75 19.25 12 16 15.25" />
    </Icon>
  ),
  '↕': (
    <Icon>
      <path d="M12 4.75v14.5M8.75 8 12 4.75 15.25 8M8.75 16 12 19.25 15.25 16" />
    </Icon>
  ),
  'A↑': (
    <Icon>
      <path d="M3.25 19 8.5 5.25 13.75 19M5.1 14.25h6.8" />
      <path d="M18 17.5V6.75M14.9 9.85 18 6.75l3.1 3.1" />
    </Icon>
  ),
  'A↓': (
    <Icon>
      <path d="M3.25 19 8.5 5.25 13.75 19M5.1 14.25h6.8" />
      <path d="M18 6.75V17.5M14.9 14.4l3.1 3.1 3.1-3.1" />
    </Icon>
  ),
  // ---- home formatting (borders / fill / merge / orientation) ----
  '⊡': <CraftIcon set={CRAFT} name="Borders" size={20} />,
  '◧': <CraftIcon set={CRAFT} name="Fill" size={20} />,
  '∅': (
    <Icon>
      <rect x="4.75" y="4.75" width="14.5" height="14.5" rx="2" />
      <path d="m6.5 17.5 11-11" />
    </Icon>
  ),
  '⇔': <CraftIcon set={CRAFT} name="Merge" size={20} />,
  '⤴': <CraftIcon set={CRAFT} name="Orientation" size={20} />,
}

/** Provider-independent assistant icon. */
export function AssistantMark({ size = 18 }: { size?: number }) {
  return (
    <svg
      width={size}
      height={size}
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth="1.7"
      strokeLinecap="round"
      strokeLinejoin="round"
      aria-hidden="true"
    >
      <path d="M5 4h14a2 2 0 0 1 2 2v10a2 2 0 0 1-2 2h-8l-5 4v-4H5a2 2 0 0 1-2-2V6a2 2 0 0 1 2-2Z" />
      <path d="m8 12 3-5 3 5M9 10h4M17 7v5" />
    </svg>
  )
}

export function BorderBottomDoubleIcon(): ReactElement {
  return (
    <BorderSvg>
      <BorderDashedFrame />
      <path d="M 3.02 11.6 h 9.96 M 3.02 13.6 h 9.96" strokeWidth="1" />
    </BorderSvg>
  )
}

export function BorderBottomThickIcon(): ReactElement {
  return (
    <BorderSvg>
      <BorderDashedFrame />
      <path d="M 3.02 12.98 h 9.96" strokeWidth="2.3" />
    </BorderSvg>
  )
}

export function BorderTopBottomIcon(): ReactElement {
  return (
    <BorderSvg>
      <BorderDashedFrame />
      <path d="M 3.02 3.02 h 9.96 M 3.02 12.98 h 9.96" />
    </BorderSvg>
  )
}

export function BorderTopThickBottomIcon(): ReactElement {
  return (
    <BorderSvg>
      <BorderDashedFrame />
      <path d="M 3.02 3.02 h 9.96" />
      <path d="M 3.02 12.98 h 9.96" strokeWidth="2.3" />
    </BorderSvg>
  )
}

export function BorderTopDoubleBottomIcon(): ReactElement {
  return (
    <BorderSvg>
      <BorderDashedFrame />
      <path d="M 3.02 3.02 h 9.96" />
      <path d="M 3.02 11.6 h 9.96 M 3.02 13.6 h 9.96" strokeWidth="1" />
    </BorderSvg>
  )
}

export function BorderInsideIcon(): ReactElement {
  return (
    <BorderSvg>
      <BorderDashedFrame />
      <path d="M 3.02 8 h 9.96 M 8 3.02 v 9.96" />
    </BorderSvg>
  )
}

export function BorderInsideHorizontalIcon(): ReactElement {
  return (
    <BorderSvg>
      <BorderDashedFrame />
      <path d="M 3.02 8 h 9.96" />
    </BorderSvg>
  )
}

export function BorderInsideVerticalIcon(): ReactElement {
  return (
    <BorderSvg>
      <BorderDashedFrame />
      <path d="M 8 3.02 v 9.96" />
    </BorderSvg>
  )
}
