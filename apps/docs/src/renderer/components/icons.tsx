/** Small monochrome SVG icons approximating Word's ribbon glyphs. */

import type { ReactNode } from 'react'
import { CraftIcon, type CraftSet } from '@genoffice/ui'
import { WORD_ICONS } from '@genoffice/ui/craft-icons/word'

/** Toolbar glyphs drawn like the WordCraft icon set (see packages/ui/src/craft-icons/NOTICE-craft-icons.md) */
const CRAFT: CraftSet = { grid: 20, icons: WORD_ICONS }

interface IconProps {
  size?: number
}

/** Constant painted stroke instead of proportional scaling: ~1.5px lines on
 *  20px+ glyphs, ~1.25px on the 13-19px ones, ~1.1px below (a proportional
 *  1-unit stroke would paint 1.75px at 28px and hairlines at small sizes).
 *  stroke-width is in 16-canvas units: units = painted-px × 16 / rendered-px.
 *  `paint` overrides the painted px: diagonal-heavy letterform icons pass 1.4
 *  as optical compensation — a slanted stroke's anti-aliasing spreads its ink
 *  over a wider footprint, so at an equal nominal width it reads a touch
 *  fatter than the axis-aligned line icons (pixel-measured on the ribbon). */
function pinnedStroke(size: number, paint?: number): number {
  const painted = paint ?? (size >= 20 ? 1.5 : size >= 13 ? 1.25 : 1.1)
  return (painted * 16) / size
}

function Svg({ size = 20, paint, children }: IconProps & { paint?: number; children: ReactNode }) {
  return (
    <svg
      width={size}
      height={size}
      viewBox="0 0 16 16"
      fill="none"
      stroke="currentColor"
      strokeWidth={pinnedStroke(size, paint)}
      strokeLinecap="round"
      strokeLinejoin="round"
      aria-hidden
    >
      {children}
    </svg>
  )
}

export function IconBullets(props: IconProps) {
  return <CraftIcon set={CRAFT} name="bullets" size={props.size ?? 20} />
}

export function IconNumbered(props: IconProps) {
  return <CraftIcon set={CRAFT} name="numbering" size={props.size ?? 20} />
}

export function IconMultilevel(props: IconProps) {
  return <CraftIcon set={CRAFT} name="multilevel" size={props.size ?? 20} />
}

export function IconIndentDec(props: IconProps) {
  return <CraftIcon set={CRAFT} name="outdent" size={props.size ?? 20} />
}

export function IconIndentInc(props: IconProps) {
  return <CraftIcon set={CRAFT} name="indent" size={props.size ?? 20} />
}

/* the whole "lines" family (align/indent/spacing/lists) shares the ink band
   y 3.44→12.98 so the row reads as one height */
export function IconAlignLeft(props: IconProps) {
  return <CraftIcon set={CRAFT} name="alignLeft" size={props.size ?? 20} />
}

export function IconAlignCenter(props: IconProps) {
  return <CraftIcon set={CRAFT} name="alignCenter" size={props.size ?? 20} />
}

export function IconAlignRight(props: IconProps) {
  return <CraftIcon set={CRAFT} name="alignRight" size={props.size ?? 20} />
}

export function IconAlignJustify(props: IconProps) {
  return <CraftIcon set={CRAFT} name="justify" size={props.size ?? 20} />
}

export function IconDirLtr(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M 3.02 3.85 h 9.96 M 3.02 6.34 h 6.64 M 3.02 11.32 h 7.1" />
      <path d="M 9.8 9.4 l 2.9 1.92 -2.9 1.92 z" fill="currentColor" stroke="none" />
    </Svg>
  )
}

export function IconDirRtl(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M 3.02 3.85 h 9.96 M 6.34 6.34 h 6.64 M 5.88 11.32 h 7.1" />
      <path d="M 6.2 9.4 l -2.9 1.92 2.9 1.92 z" fill="currentColor" stroke="none" />
    </Svg>
  )
}

export function IconLineSpacing(props: IconProps) {
  return <CraftIcon set={CRAFT} name="lineSpacing" size={props.size ?? 20} />
}

export function IconClearFormat(props: IconProps) {
  return <CraftIcon set={CRAFT} name="clear" size={props.size ?? 20} />
}

/* THE letter A — one canonical path shared by grow/shrink font, clear
   formatting, change case and font color, so every A in the ribbon is
   literally the same glyph (apex y4 → baseline y12.5, cap ≈ the 15px letter
   glyphs' cap height). dx slides it horizontally to make room for the
   companion element (arrow, eraser, lowercase a, …). */
function LetterA({ dx = 0 }: { dx?: number }) {
  return <path d={`M${2.43 + dx} 12.5 ${5.67 + dx} 4l3.25 8.5M${3.55 + dx} 9.56h4.24`} />
}

export function IconGrowFont(props: IconProps) {
  return <CraftIcon set={CRAFT} name="grow" size={props.size ?? 20} />
}

export function IconShrinkFont(props: IconProps) {
  return <CraftIcon set={CRAFT} name="shrink" size={props.size ?? 20} />
}

/** change case (Aa): the shared A + a stroke-drawn lowercase a on the same baseline */
export function IconChangeCase(props: IconProps) {
  return <CraftIcon set={CRAFT} name="case" size={props.size ?? 20} />
}

/** font color: the shared A alone, centered; the color bar is rendered by the button */
export function IconFontColorA(props: IconProps) {
  return <CraftIcon set={CRAFT} name="fontcolor" size={props.size ?? 20} />
}

/* Fluent-style sub/superscript: lowercase-x strokes + a stroked digit 2 in the
   corner (replaces the old HTML x<sub>2</sub> text glyphs that rendered smaller
   and thinner than the neighbouring 15px B/I/U letterforms) */
export function IconSuperscript(props: IconProps) {
  return <CraftIcon set={CRAFT} name="superscript" size={props.size ?? 20} />
}

export function IconSubscript(props: IconProps) {
  return <CraftIcon set={CRAFT} name="subscript" size={props.size ?? 20} />
}

export function IconHighlight(props: IconProps) {
  return <CraftIcon set={CRAFT} name="highlight" size={props.size ?? 20} />
}

/* ---------- shared shapes ---------- */

/** page outline used by many icons */
const PAGE = <path d="M4.92 3h4.62l1.93 1.93v8.09h-6.55z" />

function TextGlyph({
  x,
  y,
  s,
  children,
  bold,
}: {
  x: number
  y: number
  s: number
  children: string
  bold?: boolean
}) {
  return (
    <text
      x={x}
      y={y}
      fontSize={s}
      fill="currentColor"
      stroke="none"
      fontFamily="Segoe UI, sans-serif"
      fontWeight={bold ? 700 : 400}
    >
      {children}
    </text>
  )
}

/* ---------- clipboard (Home) ---------- */

export function IconPaste(props: IconProps) {
  return <CraftIcon set={CRAFT} name="paste" size={props.size ?? 20} />
}

export function IconCut(props: IconProps) {
  return <CraftIcon set={CRAFT} name="cut" size={props.size ?? 20} />
}

export function IconCopy(props: IconProps) {
  return <CraftIcon set={CRAFT} name="copy" size={props.size ?? 20} />
}

export function IconFormatPainter(props: IconProps) {
  return <CraftIcon set={CRAFT} name="painter" size={props.size ?? 20} />
}

/* ---------- Insert ---------- */

export function IconTable(props: IconProps) {
  return <CraftIcon set={CRAFT} name="table" size={props.size ?? 20} />
}

export function IconPicture(props: IconProps) {
  return <CraftIcon set={CRAFT} name="picture" size={props.size ?? 20} />
}

export function IconRemoveBg(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="3.02" y="3.85" width="9.96" height="8.3" rx="0.66" strokeDasharray="2.2 1.6" />
      <circle cx="8" cy="6.92" r="1.41" />
      <path d="M 5.43 12.15 c 0.33 -1.91 1.41 -2.9 2.57 -2.9 s 2.24 1 2.57 2.91" />
    </Svg>
  )
}

export function IconCrop(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M 5.33 3.17 v 7.5 h 7.5" />
      <path d="M 3.17 5.33 h 7.5 v 7.5" />
    </Svg>
  )
}

export function IconRotateRight(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M 12.4 6.2 a 4.6 4.6 0 1 0 0.6 3.3" />
      <path d="M 12.7 3.2 v 3 h -3" />
    </Svg>
  )
}

export function IconRotateLeft(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M 3.6 6.2 a 4.6 4.6 0 1 1 -0.6 3.3" />
      <path d="M 3.3 3.2 v 3 h 3" />
    </Svg>
  )
}

export function IconFlipH(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M 8 2.6 v 10.8" strokeDasharray="1.7 1.5" />
      <path d="M 6 5.2 L 2.6 8 L 6 10.8 Z" />
      <path d="M 10 5.2 L 13.4 8 L 10 10.8 Z" fill="currentColor" />
    </Svg>
  )
}

export function IconFlipV(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M 2.6 8 h 10.8" strokeDasharray="1.7 1.5" />
      <path d="M 5.2 6 L 8 2.6 L 10.8 6 Z" />
      <path d="M 5.2 10 L 8 13.4 L 10.8 10 Z" fill="currentColor" />
    </Svg>
  )
}

export function IconReplacePicture(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="2.87" y="6.03" width="7.11" height="6.32" rx="0.63" />
      <circle cx="4.92" cy="8" r="0.71" />
      <path d="M 3.26 11.79 l 2.13 -2.13 1.5 1.5 1.11 -1.11 1.42 1.42" />
      <path d="M 9.19 3.73 h 3.63 m 0 0 -1.34 -1.26 m 1.34 1.26 -1.34 1.26" />
    </Svg>
  )
}

export function IconChart(props: IconProps) {
  return <CraftIcon set={CRAFT} name="chart" size={props.size ?? 20} />
}

export function IconShapes(props: IconProps) {
  return <CraftIcon set={CRAFT} name="shapes" size={props.size ?? 20} />
}

export function IconSearch(props: IconProps) {
  return <CraftIcon set={CRAFT} name="find" size={props.size ?? 20} />
}

export function IconReplace(props: IconProps) {
  return <CraftIcon set={CRAFT} name="replace" size={props.size ?? 20} />
}

export function IconLink(props: IconProps) {
  return <CraftIcon set={CRAFT} name="link" size={props.size ?? 20} />
}

export function IconComment(props: IconProps) {
  return <CraftIcon set={CRAFT} name="comment" size={props.size ?? 20} />
}

export function IconComments(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M 5.75 3.45 h 7.28 v 5.46 h -1.68" />
      <path d="M 2.99 5.75 h 8.19 v 5.46 h -4.1 L 4.81 13.5 v -2.29 h -1.82 z" />
      <path d="M 5.27 8.2 h 3.9" />
    </Svg>
  )
}

export function IconPageBreak(props: IconProps) {
  return <CraftIcon set={CRAFT} name="pageBreak" size={props.size ?? 20} />
}

export function IconHeader(props: IconProps) {
  return <CraftIcon set={CRAFT} name="header" size={props.size ?? 20} />
}

export function IconFooter(props: IconProps) {
  return <CraftIcon set={CRAFT} name="footer" size={props.size ?? 20} />
}

export function IconPageNumber(props: IconProps) {
  return <CraftIcon set={CRAFT} name="pageNumber" size={props.size ?? 20} />
}

export function IconSymbol(props: IconProps) {
  return <CraftIcon set={CRAFT} name="symbol" size={props.size ?? 20} />
}

export function IconEquation(props: IconProps) {
  return <CraftIcon set={CRAFT} name="equation" size={props.size ?? 20} />
}

/* ---------- Table Design / Layout ---------- */

export function IconTableDelete(props: IconProps) {
  return <CraftIcon set={CRAFT} name="deleteTable" size={props.size ?? 20} />
}

export function IconAutoFit(props: IconProps) {
  return <CraftIcon set={CRAFT} name="autofit" size={props.size ?? 20} />
}

export function IconRepeatHeader(props: IconProps) {
  return <CraftIcon set={CRAFT} name="repeatHeader" size={props.size ?? 20} />
}

export function IconTableProperties(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="2.85" y="3.2" width="7.7" height="9.6" rx="0.65" />
      <path d="M 2.85 6.4 h 7.7 M 2.85 9.6 h 7.7 M 6.7 3.2 v 9.6" strokeWidth="1" />
      <path d="M 11.7 5.15 h 2.15 M 11.7 8 h 2.15 M 11.7 10.85 h 2.15" />
      <circle cx="12.35" cy="5.15" r="0.55" fill="currentColor" stroke="none" />
      <circle cx="13.15" cy="8" r="0.55" fill="currentColor" stroke="none" />
      <circle cx="12.65" cy="10.85" r="0.55" fill="currentColor" stroke="none" />
    </Svg>
  )
}

export function IconRowInsertAbove(props: IconProps) {
  return <CraftIcon set={CRAFT} name="insertAbove" size={props.size ?? 20} />
}

export function IconRowInsertBelow(props: IconProps) {
  return <CraftIcon set={CRAFT} name="insertBelow" size={props.size ?? 20} />
}

export function IconColInsertLeft(props: IconProps) {
  return <CraftIcon set={CRAFT} name="insertLeft" size={props.size ?? 20} />
}

export function IconColInsertRight(props: IconProps) {
  return <CraftIcon set={CRAFT} name="insertRight" size={props.size ?? 20} />
}

export function IconMergeCells(props: IconProps) {
  return <CraftIcon set={CRAFT} name="merge" size={props.size ?? 20} />
}

export function IconSplitCells(props: IconProps) {
  return <CraftIcon set={CRAFT} name="splitCells" size={props.size ?? 20} />
}

export function IconRowDelete(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="3.02" y="3.44" width="9.96" height="9.13" rx="0.66" />
      <path d="M 3.02 6.51 h 9.96 M 3.02 9.49 h 9.96" strokeWidth="1" />
      <path d="m 6.01 6.92 3.98 2.16 M 9.99 6.92 6.01 9.08" />
    </Svg>
  )
}

export function IconColDelete(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="3.44" y="3.02" width="9.13" height="9.96" rx="0.66" />
      <path d="M 6.51 3.02 v 9.96 M 9.49 3.02 v 9.96" strokeWidth="1" />
      <path d="m 6.92 6.01 2.16 3.98 M 9.08 6.01 6.92 9.99" />
    </Svg>
  )
}

export function IconCellAlignTop(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="3.02" y="3.44" width="9.96" height="9.13" rx="0.66" />
      <path d="M 5.1 5.68 h 5.81 M 5.1 7.5 h 3.74" />
    </Svg>
  )
}

export function IconCellAlignMiddle(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="3.02" y="3.44" width="9.96" height="9.13" rx="0.66" />
      <path d="M 5.1 7.09 h 5.81 M 5.1 8.91 h 3.74" />
    </Svg>
  )
}

export function IconCellAlignBottom(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="3.02" y="3.44" width="9.96" height="9.13" rx="0.66" />
      <path d="M 5.1 8.5 h 5.81 M 5.1 10.32 h 3.74" />
    </Svg>
  )
}

export function IconBorderAll(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="3.02" y="3.02" width="9.96" height="9.96" rx="0.42" />
      <path d="M 3.02 8 h 9.96 M 8 3.02 v 9.96" />
    </Svg>
  )
}

export function IconBorderOuter(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="3.02" y="3.02" width="9.96" height="9.96" rx="0.42" />
      <path
        d="M 3.02 8 h 9.96 M 8 3.02 v 9.96"
        strokeWidth="1"
        strokeDasharray="1.5 1.7"
        opacity="0.55"
      />
    </Svg>
  )
}

export function IconBorderInner(props: IconProps) {
  return (
    <Svg {...props}>
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
      <path d="M 3.02 8 h 9.96 M 8 3.02 v 9.96" />
    </Svg>
  )
}

export function IconBorderNone(props: IconProps) {
  return (
    <Svg {...props}>
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
      <path
        d="M 3.02 8 h 9.96 M 8 3.02 v 9.96"
        strokeWidth="1"
        strokeDasharray="1.5 1.7"
        opacity="0.55"
      />
    </Svg>
  )
}

function IconBorderDashedFrame() {
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

export function IconBorderTop(props: IconProps) {
  return (
    <Svg {...props}>
      <IconBorderDashedFrame />
      <path d="M 3.02 3.02 h 9.96" />
    </Svg>
  )
}

export function IconBorderBottom(props: IconProps) {
  return (
    <Svg {...props}>
      <IconBorderDashedFrame />
      <path d="M 3.02 12.98 h 9.96" />
    </Svg>
  )
}

export function IconBorderLeft(props: IconProps) {
  return (
    <Svg {...props}>
      <IconBorderDashedFrame />
      <path d="M 3.02 3.02 v 9.96" />
    </Svg>
  )
}

export function IconBorderRight(props: IconProps) {
  return (
    <Svg {...props}>
      <IconBorderDashedFrame />
      <path d="M 12.98 3.02 v 9.96" />
    </Svg>
  )
}

export function IconBorderInsideH(props: IconProps) {
  return (
    <Svg {...props}>
      <IconBorderDashedFrame />
      <path d="M 3.02 8 h 9.96" />
    </Svg>
  )
}

export function IconBorderInsideV(props: IconProps) {
  return (
    <Svg {...props}>
      <IconBorderDashedFrame />
      <path d="M 8 3.02 v 9.96" />
    </Svg>
  )
}

/* ---------- Design ---------- */

export function IconTheme(props: IconProps) {
  return <CraftIcon set={CRAFT} name="themes" size={props.size ?? 20} />
}

export function IconThemeFonts(props: IconProps) {
  return <CraftIcon set={CRAFT} name="fonts" size={props.size ?? 20} />
}

export function IconThemeColors(props: IconProps) {
  return <CraftIcon set={CRAFT} name="colors" size={props.size ?? 20} />
}

export function IconPageColor(props: IconProps) {
  return <CraftIcon set={CRAFT} name="pageColor" size={props.size ?? 20} />
}

export function IconWatermark(props: IconProps) {
  return <CraftIcon set={CRAFT} name="watermark" size={props.size ?? 20} />
}

export function IconPageBorders(props: IconProps) {
  return <CraftIcon set={CRAFT} name="pageBorders" size={props.size ?? 20} />
}

/* ---------- Layout ---------- */

export function IconMargins(props: IconProps) {
  return <CraftIcon set={CRAFT} name="margins" size={props.size ?? 20} />
}

export function IconOrientation(props: IconProps) {
  return <CraftIcon set={CRAFT} name="orientation" size={props.size ?? 20} />
}

export function IconPageSize(props: IconProps) {
  return <CraftIcon set={CRAFT} name="size" size={props.size ?? 20} />
}

export function IconColumns(props: IconProps) {
  return <CraftIcon set={CRAFT} name="columns" size={props.size ?? 20} />
}

/* ---------- References ---------- */

export function IconToc(props: IconProps) {
  return <CraftIcon set={CRAFT} name="toc" size={props.size ?? 20} />
}

export function IconRefresh(props: IconProps) {
  return <CraftIcon set={CRAFT} name="update" size={props.size ?? 20} />
}

export function IconFootnote(props: IconProps) {
  return <CraftIcon set={CRAFT} name="footnote" size={props.size ?? 20} />
}

export function IconEndnote(props: IconProps) {
  return <CraftIcon set={CRAFT} name="endnote" size={props.size ?? 20} />
}

export function IconCitation(props: IconProps) {
  return <CraftIcon set={CRAFT} name="citation" size={props.size ?? 20} />
}

export function IconBook(props: IconProps) {
  return <CraftIcon set={CRAFT} name="bibliography" size={props.size ?? 20} />
}

export function IconCaption(props: IconProps) {
  return <CraftIcon set={CRAFT} name="caption" size={props.size ?? 20} />
}

export function IconIndex(props: IconProps) {
  return <CraftIcon set={CRAFT} name="index" size={props.size ?? 20} />
}

/* ---------- Review ---------- */

export function IconWordCount(props: IconProps) {
  return <CraftIcon set={CRAFT} name="wordCount" size={props.size ?? 20} />
}

export function IconSpellcheck(props: IconProps) {
  return <CraftIcon set={CRAFT} name="spelling" size={props.size ?? 20} />
}

export function IconSparkle(props: IconProps) {
  return (
    <Svg {...props}>
      <path
        d="M 8 3.03 C 8 5.78 10.22 8 12.97 8 C 10.22 8 8 10.22 8 12.97 C 8 10.22 5.78 8 3.03 8 C 5.78 8 8 5.78 8 3.03 Z"
        fill="currentColor"
        stroke="none"
      />
    </Svg>
  )
}

export function IconWand(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M 3.6 12.4 9.6 6.4" strokeWidth="1" />
      <path
        d="M 11.04 2.8 l 0.56 1.52 1.52 0.56 -1.52 0.56 -0.56 1.52 -0.56 -1.52 -1.52 -0.56 1.52 -0.56 z"
        fill="currentColor"
        stroke="none"
      />
      <path
        d="M 12.4 8.4 l 0.32 0.88 0.88 0.32 -0.88 0.32 -0.32 0.88 -0.32 -0.88 -0.88 -0.32 0.88 -0.32 z"
        fill="currentColor"
        stroke="none"
      />
    </Svg>
  )
}

export function IconTranslate(props: IconProps) {
  return <CraftIcon set={CRAFT} name="translate" size={props.size ?? 20} />
}

export function IconTrackChanges(props: IconProps) {
  return <CraftIcon set={CRAFT} name="trackChanges" size={props.size ?? 20} />
}

export function IconAccept(props: IconProps) {
  return <CraftIcon set={CRAFT} name="accept" size={props.size ?? 20} />
}

export function IconReject(props: IconProps) {
  return <CraftIcon set={CRAFT} name="reject" size={props.size ?? 20} />
}

export function IconCompare(props: IconProps) {
  return <CraftIcon set={CRAFT} name="compare" size={props.size ?? 20} />
}

export function IconLock(props: IconProps) {
  return <CraftIcon set={CRAFT} name="protect" size={props.size ?? 20} />
}

export function IconKey(props: IconProps) {
  return (
    <Svg {...props}>
      <circle cx="5.4" cy="10.6" r="2.5" />
      <path d="M 7.3 8.7 L 12.9 3.1 M 10.4 5.6 l 1.7 1.7 M 12.1 3.9 l 1.4 1.4" />
    </Svg>
  )
}

export function IconEye(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M 1.7 8 C 3.1 5.15 5.35 3.6 8 3.6 s 4.9 1.55 6.3 4.4 C 12.9 10.85 10.65 12.4 8 12.4 S 3.1 10.85 1.7 8 Z" />
      <circle cx="8" cy="8" r="1.95" />
    </Svg>
  )
}

export function IconEyeOff(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M 3.5 4.65 C 2.78 5.5 2.18 6.62 1.7 8 c 1.4 2.85 3.65 4.4 6.3 4.4 c 1.02 0 1.97 -0.23 2.84 -0.68 M 6.4 3.82 C 6.91 3.67 7.44 3.6 8 3.6 c 2.65 0 4.9 1.55 6.3 4.4 c -0.4 0.82 -0.87 1.54 -1.4 2.15" />
      <path d="M 6.62 6.62 a 1.95 1.95 0 0 0 2.76 2.76" />
      <path d="M 2.7 2.7 l 10.6 10.6" />
    </Svg>
  )
}

export function IconAlert(props: IconProps) {
  return (
    <Svg {...props}>
      <circle cx="8" cy="8" r="6.2" />
      <path d="M 8 4.9 v 3.5" />
      <circle cx="8" cy="10.9" r="0.75" fill="currentColor" stroke="none" />
    </Svg>
  )
}

/* ---------- View ---------- */

function Magnifier({ children }: { children?: ReactNode }) {
  return (
    <>
      <circle cx="7.15" cy="7.15" r="4.08" />
      <path d="M 10.21 10.21 13.1 13.1" strokeWidth="1" />
      {children}
    </>
  )
}

export function IconZoomOut(props: IconProps) {
  return (
    <Svg {...props}>
      <Magnifier>
        <path d="M 5.28 7.15 h 3.74" />
      </Magnifier>
    </Svg>
  )
}

export function IconZoomIn(props: IconProps) {
  return (
    <Svg {...props}>
      <Magnifier>
        <path d="M 5.28 7.15 h 3.74 M 7.15 5.28 v 3.74" />
      </Magnifier>
    </Svg>
  )
}

export function IconZoom100(props: IconProps) {
  return <CraftIcon set={CRAFT} name="zoom100" size={props.size ?? 20} />
}

export function IconPageWidth(props: IconProps) {
  return <CraftIcon set={CRAFT} name="pageWidth" size={props.size ?? 20} />
}

export function IconWholePage(props: IconProps) {
  return <CraftIcon set={CRAFT} name="onePage" size={props.size ?? 20} />
}

export function IconAiPanel(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="3" y="3.76" width="10.01" height="8.47" rx="0.62" />
      <path d="M 9.39 3.76 v 8.47" />
      <path
        d="M 10.31 6.61 l 0.39 1 1 0.39 -1 0.39 -0.38 1 -0.38 -1 -1 -0.38 1 -0.38 z"
        fill="currentColor"
        stroke="none"
      />
    </Svg>
  )
}

export function IconMoon(props: IconProps) {
  return <CraftIcon set={CRAFT} name="darkMode" size={props.size ?? 20} />
}

export function IconReadMode(props: IconProps) {
  return <CraftIcon set={CRAFT} name="readMode" size={props.size ?? 20} />
}

export function IconOutlineView(props: IconProps) {
  return <CraftIcon set={CRAFT} name="outline" size={props.size ?? 20} />
}

export function IconRuler(props: IconProps) {
  return <CraftIcon set={CRAFT} name="ruler" size={props.size ?? 20} />
}

export function IconNavPane(props: IconProps) {
  return <CraftIcon set={CRAFT} name="navPane" size={props.size ?? 20} />
}

export function IconSplit(props: IconProps) {
  return <CraftIcon set={CRAFT} name="split" size={props.size ?? 20} />
}

export function IconPrintLayout(props: IconProps) {
  return <CraftIcon set={CRAFT} name="printLayout" size={props.size ?? 20} />
}

export function IconWebLayout(props: IconProps) {
  return <CraftIcon set={CRAFT} name="webLayout" size={props.size ?? 20} />
}

export function IconGridlines(props: IconProps) {
  return <CraftIcon set={CRAFT} name="gridlines" size={props.size ?? 20} />
}

export function IconNewWindow(props: IconProps) {
  return <CraftIcon set={CRAFT} name="newWindow" size={props.size ?? 20} />
}

export function IconArrangeAll(props: IconProps) {
  return <CraftIcon set={CRAFT} name="arrangeAll" size={props.size ?? 20} />
}

export function IconSwitchWindows(props: IconProps) {
  return <CraftIcon set={CRAFT} name="switchWindows" size={props.size ?? 20} />
}

export function IconPosition(props: IconProps) {
  return <CraftIcon set={CRAFT} name="position" size={props.size ?? 20} />
}

export function IconWrapText(props: IconProps) {
  return <CraftIcon set={CRAFT} name="wrapText" size={props.size ?? 20} />
}

export function IconDoc(props: IconProps) {
  return (
    <Svg {...props}>
      {PAGE}
      <path d="M 9.54 3 V 4.92 h 1.93" />
      <path d="M 6.08 6.84 h 3.85 M 6.08 8.77 h 3.85 M 6.08 10.7 h 2.7" />
    </Svg>
  )
}

/* ---------- AI panel ---------- */

export function IconSend(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M 3.01 8 12.99 3.36 10.58 12.64 7.66 9.38 z" strokeLinejoin="round" />
      <path d="M 7.66 9.38 12.99 3.36" />
    </Svg>
  )
}

export function IconStop(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="3" y="3" width="10" height="10" rx="1.88" fill="currentColor" stroke="none" />
    </Svg>
  )
}

export function IconGear(props: IconProps) {
  return (
    <Svg {...props}>
      <circle cx="8" cy="8" r="1.78" />
      <path d="M 8 2.98 v 1.62 M 8 11.4 v 1.62 M 13.02 8 h -1.62 M 4.6 8 h -1.62 M 11.56 4.44 l -1.13 1.13 M 5.57 10.43 l -1.13 1.13 M 11.56 11.56 10.43 10.43 M 5.57 5.57 4.44 4.44" />
    </Svg>
  )
}

/** collapse the left-docked AI panel: sheets-parity glyph (16-canvas, 1.2/1.3 stroke),
 *  self-contained so the shared Svg wrapper's pinned stroke doesn't alter its weight */
export function IconSidebarCollapse({ size = 20 }: IconProps) {
  return (
    <svg
      width={size}
      height={size}
      viewBox="0 0 16 16"
      fill="none"
      stroke="currentColor"
      strokeWidth="1.2"
      strokeLinecap="round"
      aria-hidden
    >
      <rect x="1.5" y="2.5" width="13" height="11" rx="1" />
      <path d="M5.5 2.5v11" />
      <path d="M12.5 8H8.1M9.8 5.9 7.7 8l2.1 2.1" strokeWidth="1.3" strokeLinejoin="round" />
    </svg>
  )
}

export function IconClock(props: IconProps) {
  return (
    <Svg {...props}>
      <circle cx="8" cy="8" r="4.98" />
      <path d="M 8 5.34 V 8 l 1.91 1.33" />
    </Svg>
  )
}

export function IconPaperclip(props: IconProps) {
  return (
    <Svg {...props}>
      <path
        d="M 12.5 7.28 8.18 11.6 a 3.06 3.06 0 0 1 -4.32 -4.32 l 4.5 -4.5 a 2.07 2.07 0 0 1 2.88 2.88 l -4.5 4.5 a 0.99 0.99 0 0 1 -1.44 -1.44 l 4.14 -4.14"
        strokeLinejoin="round"
      />
    </Svg>
  )
}

export function IconNewChat(props: IconProps) {
  return (
    <Svg {...props}>
      <path
        d="M 12.68 7.32 v -2.55 A 1.44 1.44 0 0 0 11.23 3.33 H 4.77 a 1.44 1.44 0 0 0 -1.44 1.44 v 5.19 a 1.44 1.44 0 0 0 1.44 1.44 h 0.94 v 1.7 l 2.21 -1.7 h 1.11"
        strokeLinejoin="round"
      />
      <path d="M 11.57 9.19 v 3.4 M 9.87 10.89 h 3.4" />
    </Svg>
  )
}

/* ---------- titlebar quick access ---------- */

/** Design-supplied glyphs on the 1:16 stroke:canvas ratio (24-canvas / 1.5 stroke):
 *  the stroke scales proportionally with size instead of the pinnedStroke policy. */
function SvgRatio({ size = 24, children }: IconProps & { children: ReactNode }) {
  return (
    <svg
      width={size}
      height={size}
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth={1.5}
      strokeLinecap="round"
      strokeLinejoin="round"
      aria-hidden
    >
      {children}
    </svg>
  )
}

export function IconSave(props: IconProps) {
  return <CraftIcon set={CRAFT} name="save" size={props.size ?? 20} />
}

export function IconUndo(props: IconProps) {
  return <CraftIcon set={CRAFT} name="undo" size={props.size ?? 20} />
}

export function IconRedo(props: IconProps) {
  return <CraftIcon set={CRAFT} name="redo" size={props.size ?? 20} />
}

export function IconCursor(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M 3.98 3.01 12.02 8.53 l -3.41 0.81 L 6.99 12.95 3.98 3.01 Z" />
    </Svg>
  )
}

export function IconPen(props: IconProps) {
  return <CraftIcon set={CRAFT} name="pen" size={props.size ?? 20} />
}

export function IconHighlighterPen(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M 5.68 9.4 10.6 4.47 a 1.21 1.21 0 0 1 1.77 0 l -0.84 -0.84 0.84 0.84 a 1.21 1.21 0 0 1 0 1.77 L 7.44 11.16 l -2.42 0.65 0.65 -2.42 Z" />
      <path d="M 3.35 13.58 h 9.3" strokeWidth="1" opacity="0.5" />
    </Svg>
  )
}

export function IconEraser(props: IconProps) {
  return <CraftIcon set={CRAFT} name="eraser" size={props.size ?? 20} />
}

export function IconTextBox(props: IconProps) {
  return <CraftIcon set={CRAFT} name="textBox" size={props.size ?? 20} />
}

export function IconWordArt(props: IconProps) {
  return <CraftIcon set={CRAFT} name="wordArt" size={props.size ?? 20} />
}

export function IconPencil(props: IconProps) {
  return <CraftIcon set={CRAFT} name="pencil" size={props.size ?? 20} />
}

export function IconTrash(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M3.11 4.89h9.79M6.4 4.89V3.73a.62.62 0 0 1 .62-.62h1.96a.62.62 0 0 1 .62.62v1.16" />
      <path d="M4.44 4.89l.62 7.39a.89.89 0 0 0 .89.8h4.09a.89.89 0 0 0 .89-.8l.62-7.39" />
      <path d="M6.75 7.11v3.56M9.25 7.11v3.56" />
    </Svg>
  )
}

/** Thin dropdown chevron replacing the ▾ text glyph (same path as the slides ribbon's RbCaret);
 *  1.5 stroke on a 24 viewBox keeps the 1 : 16 stroke : canvas ratio. */
export function IconCaret({ size = 10 }: IconProps) {
  return (
    <svg
      className="rb-caret-svg"
      width={size}
      height={size}
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth="2.6"
      strokeLinecap="round"
      strokeLinejoin="round"
      aria-hidden
    >
      <path d="M5.5 9.25 12 15.75l6.5-6.5" />
    </svg>
  )
}

export function IconPalette(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M8 12.98a4.98 4.98 0 1 1 4.98-4.98c0 2.44-1.74 2.49-2.74 2.49-.8 0-1.25.5-1.25 1.25 0 .7-.45 1.25-1 1.25Z" />
      <circle cx="8.83" cy="4.93" r="0.71" fill="currentColor" stroke="none" />
      <circle cx="11.07" cy="6.71" r="0.71" fill="currentColor" stroke="none" />
      <circle cx="6.09" cy="5.51" r="0.71" fill="currentColor" stroke="none" />
      <circle cx="4.93" cy="8.25" r="0.71" fill="currentColor" stroke="none" />
    </Svg>
  )
}

export function IconSort(props: IconProps) {
  return <CraftIcon set={CRAFT} name="sort" size={props.size ?? 20} />
}

export function IconPilcrow(props: IconProps) {
  return <CraftIcon set={CRAFT} name="pilcrow" size={props.size ?? 20} />
}

export function IconShading(props: IconProps) {
  return <CraftIcon set={CRAFT} name="shading" size={props.size ?? 20} />
}

export function IconCheck(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M2.8 8.6 6.2 12l7-7.5" />
    </Svg>
  )
}

export function IconCheckbox(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="2.99" y="2.99" width="10.01" height="10.01" rx="1.37" />
    </Svg>
  )
}

export function IconCheckboxChecked(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="2.99" y="2.99" width="10.01" height="10.01" rx="1.37" />
      <path d="M 5.45 8.27 l 1.91 2 3.37 -4.19" />
    </Svg>
  )
}

export function IconClose(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M3.5 3.5l9 9M12.5 3.5l-9 9" />
    </Svg>
  )
}

/** Provider-independent assistant icon. */
export function AssistantMark({ size = 30 }: { size?: number }) {
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

export function IconBreaks(props: IconProps) {
  return <CraftIcon set={CRAFT} name="breaks" size={props.size ?? 20} />
}

export function IconLineNumbers(props: IconProps) {
  return <CraftIcon set={CRAFT} name="lineNumbers" size={props.size ?? 20} />
}

export function IconHyphenation(props: IconProps) {
  return <CraftIcon set={CRAFT} name="hyphenation" size={props.size ?? 20} />
}

export function IconBringForward(props: IconProps) {
  return <CraftIcon set={CRAFT} name="bringForward" size={props.size ?? 20} />
}

export function IconSendBackward(props: IconProps) {
  return <CraftIcon set={CRAFT} name="sendBackward" size={props.size ?? 20} />
}

export function IconSelectionPane(props: IconProps) {
  return <CraftIcon set={CRAFT} name="selectionPane" size={props.size ?? 20} />
}

export function IconAlignObjects(props: IconProps) {
  return <CraftIcon set={CRAFT} name="align" size={props.size ?? 20} />
}

export function IconGroupObjects(props: IconProps) {
  return <CraftIcon set={CRAFT} name="group" size={props.size ?? 20} />
}

export function IconRotateObject(props: IconProps) {
  return <CraftIcon set={CRAFT} name="rotate" size={props.size ?? 20} />
}

export function IconParaSpacing(props: IconProps) {
  return <CraftIcon set={CRAFT} name="paraSpacing" size={props.size ?? 20} />
}

export function IconThemeEffects(props: IconProps) {
  return <CraftIcon set={CRAFT} name="effects" size={props.size ?? 20} />
}

export function IconSetDefault(props: IconProps) {
  return <CraftIcon set={CRAFT} name="setDefault" size={props.size ?? 20} />
}

export function IconStylesPane(props: IconProps) {
  return <CraftIcon set={CRAFT} name="stylesPane" size={props.size ?? 20} />
}

export function IconSpaceBefore(props: IconProps) {
  return <CraftIcon set={CRAFT} name="spaceBefore" size={props.size ?? 20} />
}

export function IconSpaceAfter(props: IconProps) {
  return <CraftIcon set={CRAFT} name="spaceAfter" size={props.size ?? 20} />
}

export function IconMultiplePages(props: IconProps) {
  return <CraftIcon set={CRAFT} name="multiplePages" size={props.size ?? 20} />
}

export function IconBorders(props: IconProps) {
  return <CraftIcon set={CRAFT} name="borders" size={props.size ?? 20} />
}
