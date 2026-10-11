/** Small monochrome SVG icons approximating Word's ribbon glyphs. */

import type { ReactNode } from 'react'
import { CraftIcon, stableHtml, type CraftSet } from '@genoffice/ui'
import { DECK_ICONS } from '@genoffice/ui/craft-icons/deck'
import type { AnimEffectKind } from '../../shared/ipc'

/** Toolbar glyphs drawn like the DeckCraft icon set (see packages/ui/src/craft-icons/NOTICE-craft-icons.md) */
const CRAFT: CraftSet = { grid: 24, icons: DECK_ICONS }

interface IconProps {
  size?: number
}

/** Constant painted stroke instead of proportional scaling: ~1.5px lines on
 *  20px+ glyphs, ~1.25px on the 13-19px ones, ~1.1px below (a proportional
 *  1.5-unit stroke would paint 1.75px at 28px and hairlines at small sizes).
 *  stroke-width is in 24-canvas units: units = painted-px × 24 / rendered-px. */
function pinnedStroke(size: number): number {
  const painted = size >= 20 ? 1.5 : size >= 13 ? 1.25 : 1.1
  return (painted * 24) / size
}

function Svg({ size = 24, children }: IconProps & { children: ReactNode }) {
  return (
    <svg
      width={size}
      height={size}
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth={pinnedStroke(size)}
      strokeLinecap="round"
      strokeLinejoin="round"
      aria-hidden
    >
      {children}
    </svg>
  )
}

/** Renders gallery glyph markup (inner SVG for a 24×24 viewBox, stroke=currentColor). */
export function IconGlyph({ body, size = 18 }: { body: string; size?: number }) {
  return (
    <svg
      width={size}
      height={size}
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth={pinnedStroke(size)}
      strokeLinecap="round"
      strokeLinejoin="round"
      aria-hidden
      dangerouslySetInnerHTML={stableHtml(body)}
    />
  )
}

export function IconFind(props: IconProps) {
  return <CraftIcon set={CRAFT} name="search" size={props.size ?? 24} />
}

export function IconBullets(props: IconProps) {
  return <CraftIcon set={CRAFT} name="bullets" size={props.size ?? 24} />
}

export function IconNumbered(props: IconProps) {
  return <CraftIcon set={CRAFT} name="numbering" size={props.size ?? 24} />
}

export function IconIndentDec(props: IconProps) {
  return <CraftIcon set={CRAFT} name="decrease-indent" size={props.size ?? 24} />
}

export function IconIndentInc(props: IconProps) {
  return <CraftIcon set={CRAFT} name="increase-indent" size={props.size ?? 24} />
}

export function IconAlignLeft(props: IconProps) {
  return <CraftIcon set={CRAFT} name="align-left" size={props.size ?? 24} />
}

export function IconDirLtr(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M 4.53 5.78 h 14.94 M 4.53 9.51 h 9.96 M 4.53 16.98 h 10.65" />
      <path d="M 14.7 14.1 l 4.35 2.88 -4.35 2.88 z" fill="currentColor" stroke="none" />
    </Svg>
  )
}

export function IconDirRtl(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M 4.53 5.78 h 14.94 M 9.51 9.51 h 9.96 M 8.82 16.98 h 10.65" />
      <path d="M 9.3 14.1 l -4.35 2.88 4.35 2.88 z" fill="currentColor" stroke="none" />
    </Svg>
  )
}

export function IconAlignCenter(props: IconProps) {
  return <CraftIcon set={CRAFT} name="align-center" size={props.size ?? 24} />
}

export function IconAlignRight(props: IconProps) {
  return <CraftIcon set={CRAFT} name="align-right" size={props.size ?? 24} />
}

export function IconAlignJustify(props: IconProps) {
  return <CraftIcon set={CRAFT} name="justify" size={props.size ?? 24} />
}

export function IconLineSpacing(props: IconProps) {
  return <CraftIcon set={CRAFT} name="line-spacing" size={props.size ?? 24} />
}

export function IconClearFormat(props: IconProps) {
  return <CraftIcon set={CRAFT} name="clear-formatting" size={props.size ?? 24} />
}

export function IconGrowFont(props: IconProps) {
  return <CraftIcon set={CRAFT} name="font-grow" size={props.size ?? 24} />
}

export function IconShrinkFont(props: IconProps) {
  return <CraftIcon set={CRAFT} name="font-shrink" size={props.size ?? 24} />
}

/* Fluent-style sub/superscript: lowercase-x strokes + a stroked digit 2 in the
   corner (replaces the old HTML x<sub>2</sub> text glyphs) */
export function IconSuperscript(props: IconProps) {
  return <CraftIcon set={CRAFT} name="superscript" size={props.size ?? 24} />
}

export function IconSubscript(props: IconProps) {
  return <CraftIcon set={CRAFT} name="subscript" size={props.size ?? 24} />
}

/** Character spacing (MS-style): AV above a double-headed arrow */
export function IconCharSpacing(props: IconProps) {
  return <CraftIcon set={CRAFT} name="char-spacing" size={props.size ?? 24} />
}

/** Change Case (MS-style): "Aa" */
export function IconChangeCase(props: IconProps) {
  return <CraftIcon set={CRAFT} name="change-case" size={props.size ?? 24} />
}

/** Align Text (vertical alignment): text lines centred in a frame, arrows to top and bottom */
export function IconAlignText(props: IconProps) {
  return <CraftIcon set={CRAFT} name="align-text" size={props.size ?? 24} />
}

/** Text Direction: a rotated "A" beside a downward arrow */
export function IconTextDirection(props: IconProps) {
  return <CraftIcon set={CRAFT} name="text-direction" size={props.size ?? 24} />
}

/** MS-style text highlighter: marker nib only; the color bar is rendered by the button */
export function IconTextHighlight(props: IconProps) {
  return <CraftIcon set={CRAFT} name="highlight" size={props.size ?? 24} />
}

export function IconHighlight(props: IconProps) {
  return <CraftIcon set={CRAFT} name="highlight" size={props.size ?? 24} />
}

/* ---------- shared shapes ---------- */

/** page outline used by many icons */
const PAGE = <path d="M6 2.25 h9 l3.75 3.75 v15.75 h-12.75 z" />

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
  return <CraftIcon set={CRAFT} name="paste" size={props.size ?? 24} />
}

export function IconCut(props: IconProps) {
  return <CraftIcon set={CRAFT} name="cut" size={props.size ?? 24} />
}

export function IconCopy(props: IconProps) {
  return <CraftIcon set={CRAFT} name="copy" size={props.size ?? 24} />
}

/** Shape style gallery: two offset rounded rects (context-menu quick bar) */
export function IconShapeStyle(props: IconProps) {
  return <CraftIcon set={CRAFT} name="quick-styles" size={props.size ?? 24} />
}

/** Shape Effects: a rounded square casting an offset shadow */
export function IconShapeEffects(props: IconProps) {
  return (
    <Svg size={props.size ?? 24}>
      <rect x="3.5" y="3.5" width="12" height="12" rx="2.5" />
      <path d="M18.5 8.5v9a2.5 2.5 0 0 1-2.5 2.5H8" />
      <path d="M20.5 11.5v6a3.5 3.5 0 0 1-3.5 3.5h-6" opacity="0.5" />
    </Svg>
  )
}

/** Paint bucket + drop (context-menu quick bar fill entry) */
export function IconFillColor(props: IconProps) {
  return <CraftIcon set={CRAFT} name="shape-fill" size={props.size ?? 24} />
}

/** Concentric squares suggesting an outline ring (context-menu quick bar outline entry) */
export function IconOutlineColor(props: IconProps) {
  return <CraftIcon set={CRAFT} name="shape-outline" size={props.size ?? 24} />
}

export function IconFormatPainter(props: IconProps) {
  return <CraftIcon set={CRAFT} name="format-painter" size={props.size ?? 24} />
}

/** Slide layout: slide frame + title line + two content placeholders */
export function IconSlideLayout(props: IconProps) {
  return <CraftIcon set={CRAFT} name="layout" size={props.size ?? 24} />
}

/* ---------- Insert ---------- */

export function IconTable(props: IconProps) {
  return <CraftIcon set={CRAFT} name="table" size={props.size ?? 24} />
}

export function IconPicture(props: IconProps) {
  return <CraftIcon set={CRAFT} name="picture" size={props.size ?? 24} />
}

export function IconShapes(props: IconProps) {
  return <CraftIcon set={CRAFT} name="shapes" size={props.size ?? 24} />
}

export function IconLink(props: IconProps) {
  return <CraftIcon set={CRAFT} name="link" size={props.size ?? 24} />
}

export function IconComment(props: IconProps) {
  return <CraftIcon set={CRAFT} name="comment" size={props.size ?? 24} />
}

export function IconPageBreak(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M 7.38 4.49 h 9.24 v 5.2 M 7.38 4.49 v 5.2 M 7.38 19.51 h 9.24 v -5.2 M 7.38 19.51 v -5.2" />
      <path d="M 4.49 12 h 2.31 M 8.54 12 h 2.31 M 12.58 12 h 2.31 M 16.62 12 h 2.89" />
    </Svg>
  )
}

export function IconHeader(props: IconProps) {
  return <CraftIcon set={CRAFT} name="header-footer" size={props.size ?? 24} />
}

export function IconFooter(props: IconProps) {
  return <CraftIcon set={CRAFT} name="header-footer" size={props.size ?? 24} />
}

export function IconPageNumber(props: IconProps) {
  return <CraftIcon set={CRAFT} name="slide-number" size={props.size ?? 24} />
}

export function IconSymbol(props: IconProps) {
  return <CraftIcon set={CRAFT} name="symbol" size={props.size ?? 24} />
}

export function IconEquation(props: IconProps) {
  return <CraftIcon set={CRAFT} name="equation" size={props.size ?? 24} />
}

/* ---------- Design ---------- */

export function IconSlideMaster(props: IconProps) {
  return <CraftIcon set={CRAFT} name="slide-master" size={props.size ?? 24} />
}

export function IconTheme(props: IconProps) {
  return <CraftIcon set={CRAFT} name="themes" size={props.size ?? 24} />
}

export function IconThemeFonts(props: IconProps) {
  return <CraftIcon set={CRAFT} name="fonts" size={props.size ?? 24} />
}

export function IconThemeColors(props: IconProps) {
  return <CraftIcon set={CRAFT} name="colors" size={props.size ?? 24} />
}

export function IconPageColor(props: IconProps) {
  return <CraftIcon set={CRAFT} name="background-styles" size={props.size ?? 24} />
}

export function IconWatermark(props: IconProps) {
  return (
    <Svg {...props}>
      {PAGE}
      <path d="M7.8 17.25 16.2 7.5" opacity="0.45" />
    </Svg>
  )
}

export function IconPageBorders(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="4.53" y="4.53" width="14.94" height="14.94" rx="1" />
      <rect x="7.52" y="7.52" width="8.96" height="8.96" />
    </Svg>
  )
}

/* ---------- Layout ---------- */

export function IconMargins(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="6.23" y="4.49" width="11.55" height="15.02" rx="0.92" />
      <rect x="8.77" y="7.03" width="6.47" height="9.93" strokeDasharray="2.4 2.1" />
    </Svg>
  )
}

export function IconOrientation(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="4.8" y="7.2" width="9" height="12" rx="0.96" />
      <rect x="9" y="11.4" width="10.8" height="7.8" rx="0.96" fill="var(--surface, #fff)" />
      <path d="M 15.6 5.04 a 6 6 0 0 1 3.6 3.12 M 19.2 5.4 v 3 h -3" />
    </Svg>
  )
}

/* landscape slide with a diagonal resize arrow — PowerPoint's Slide Size glyph */
export function IconPageSize(props: IconProps) {
  return <CraftIcon set={CRAFT} name="slide-size" size={props.size ?? 24} />
}

/* stacked slides with a check on the front one — applied across the deck */
export function IconApplyAll(props: IconProps) {
  return <CraftIcon set={CRAFT} name="apply-to-all" size={props.size ?? 24} />
}

export function IconColumns(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M 4.47 5.15 h 6.16 M 4.47 8.58 h 6.16 M 4.47 12 h 6.16 M 4.47 15.42 h 6.16 M 4.47 18.85 h 6.16" />
      <path d="M 13.37 5.15 h 6.16 M 13.37 8.58 h 6.16 M 13.37 12 h 6.16 M 13.37 15.42 h 6.16 M 13.37 18.85 h 6.16" />
    </Svg>
  )
}

/* ---------- References ---------- */

export function IconToc(props: IconProps) {
  return (
    <Svg {...props}>
      {PAGE}
      <path d="M8.25 7.5 h7.5 M10.05 10.95 h5.7 M10.05 14.4 h5.7 M8.25 17.85 h7.5" />
    </Svg>
  )
}

export function IconRefresh(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M 19.02 9.98 a 7.29 7.29 0 0 0 -13.5 -1.62 M 4.98 14.03 a 7.29 7.29 0 0 0 13.5 1.62" />
      <path d="M 19.43 4.57 v 4.05 h -4.05 M 4.57 19.43 v -4.05 h 4.05" />
    </Svg>
  )
}

export function IconFootnote(props: IconProps) {
  return (
    <Svg {...props}>
      <TextGlyph x={2.25} y={18} s={13.5}>
        AB
      </TextGlyph>
      <TextGlyph x={17.25} y={12} s={10.5} bold>
        1
      </TextGlyph>
    </Svg>
  )
}

export function IconEndnote(props: IconProps) {
  return (
    <Svg {...props}>
      <TextGlyph x={2.25} y={18} s={13.5}>
        AB
      </TextGlyph>
      <TextGlyph x={16.95} y={12} s={10.5} bold>
        n
      </TextGlyph>
    </Svg>
  )
}

export function IconCitation(props: IconProps) {
  return (
    <Svg {...props}>
      <TextGlyph x={3} y={18.75} s={21} bold>
        “”
      </TextGlyph>
    </Svg>
  )
}

export function IconBook(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M 12 6.13 C 10.43 4.95 7.82 4.43 4.82 4.69 v 13.57 c 3 -0.26 5.61 0.26 7.18 1.44 1.57 -1.17 4.18 -1.7 7.18 -1.44 V 4.69 c -3 -0.26 -5.61 0.26 -7.18 1.44 z" />
      <path d="M 12 6.13 v 13.57" />
    </Svg>
  )
}

export function IconCaption(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M 4.53 7.64 h 10.58 L 19.47 12 l -4.36 4.36 H 4.53 z" />
      <circle cx="8.27" cy="12" r="1.12" fill="currentColor" stroke="none" />
    </Svg>
  )
}

export function IconIndex(props: IconProps) {
  return (
    <Svg {...props}>
      <TextGlyph x={2.7} y={9.75} s={9.75}>
        A
      </TextGlyph>
      <TextGlyph x={2.7} y={20.25} s={9.75}>
        B
      </TextGlyph>
      <path d="M12 6.75 h9 M12 12 h9 M12 17.25 h9" />
    </Svg>
  )
}

/* ---------- Review ---------- */

export function IconWordCount(props: IconProps) {
  return (
    <Svg {...props}>
      <TextGlyph x={2.4} y={12} s={12}>
        123
      </TextGlyph>
      <path d="M3 16.5 h18 M3 20.25 h12" />
    </Svg>
  )
}

export function IconSpellcheck(props: IconProps) {
  return <CraftIcon set={CRAFT} name="spelling" size={props.size ?? 24} />
}

export function IconSparkle(props: IconProps) {
  return (
    <Svg {...props}>
      <path
        d="M 12 4.55 C 12 8.67 15.33 12 19.45 12 C 15.33 12 12 15.33 12 19.45 C 12 15.33 8.67 12 4.55 12 C 8.67 12 12 8.67 12 4.55 Z"
        fill="currentColor"
        stroke="none"
      />
    </Svg>
  )
}

/** Remove Background: dashed marching-ants selection around a landscape photo
 * (sun + mountains), an AI sparkle in the top-right notch. */
export function IconRemoveBg(props: IconProps) {
  return <CraftIcon set={CRAFT} name="remove-background" size={props.size ?? 24} />
}

export function IconWand(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M 5.4 18.6 14.4 9.6" />
      <path
        d="M 16.56 4.2 l 0.84 2.28 2.28 0.84 -2.28 0.84 -0.84 2.28 -0.84 -2.28 -2.28 -0.84 2.28 -0.84 z"
        fill="currentColor"
        stroke="none"
      />
      <path
        d="M 18.6 12.6 l 0.48 1.32 1.32 0.48 -1.32 0.48 -0.48 1.32 -0.48 -1.32 -1.32 -0.48 1.32 -0.48 z"
        fill="currentColor"
        stroke="none"
      />
    </Svg>
  )
}

export function IconTranslate(props: IconProps) {
  return <CraftIcon set={CRAFT} name="translate" size={props.size ?? 24} />
}

export function IconTrackChanges(props: IconProps) {
  return (
    <Svg {...props}>
      {PAGE}
      <path d="M8.25 8.25 h7.5 M8.25 12 h4.5" />
      <path d="M12.75 19.8 20.4 12.15 l1.8 1.8 -7.65 7.65 -2.7 0.9 z" fill="var(--surface, #fff)" />
    </Svg>
  )
}

export function IconAccept(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M3.75 12.75 9 18 l11.25 -12" />
    </Svg>
  )
}

export function IconReject(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M3.75 3.75 20.25 20.25 M20.25 3.75 l-16.5 16.5" />
    </Svg>
  )
}

export function IconCompare(props: IconProps) {
  return <CraftIcon set={CRAFT} name="compare" size={props.size ?? 24} />
}

export function IconLock(props: IconProps) {
  return <CraftIcon set={CRAFT} name="lock" size={props.size ?? 24} />
}

/* ---------- View ---------- */

function Magnifier({ children }: { children?: ReactNode }) {
  return (
    <>
      <circle cx="10.5" cy="10.5" r="7.2" />
      <path d="M15.9 15.9 21 21" />
      {children}
    </>
  )
}

export function IconZoomOut(props: IconProps) {
  return <CraftIcon set={CRAFT} name="zoom-out" size={props.size ?? 24} />
}

export function IconZoomIn(props: IconProps) {
  return <CraftIcon set={CRAFT} name="zoom-in" size={props.size ?? 24} />
}

export function IconZoom100(props: IconProps) {
  return (
    <Svg {...props}>
      <Magnifier>
        <TextGlyph x={5.7} y={13.5} s={7.5} bold>
          1:1
        </TextGlyph>
      </Magnifier>
    </Svg>
  )
}

export function IconPageWidth(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="4.53" y="4.53" width="14.94" height="14.94" rx="1" />
      <path d="M 7.02 12 h 9.96 M 9.26 9.76 7.02 12 l 2.24 2.24 M 14.74 9.76 16.98 12 l -2.24 2.24" />
    </Svg>
  )
}

/** View tab fit-to-window, mirroring Fluent's zoom-fit: a small content rect
 *  framed by four edge chevrons. Deliberately arrow-free — inward/outward
 *  arrows read as fullscreen, and the original page-with-plus as "new page". */
export function IconFitWindow(props: IconProps) {
  return <CraftIcon set={CRAFT} name="fit-to-window" size={props.size ?? 24} />
}

export function IconAiPanel(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="4.49" y="5.65" width="15.02" height="12.71" rx="0.92" />
      <path d="M 14.08 5.65 v 12.71" />
      <path
        d="M 15.47 9.92 l 0.58 1.5 1.5 0.58 -1.5 0.58 -0.58 1.5 -0.58 -1.5 -1.5 -0.58 1.5 -0.58 z"
        fill="currentColor"
        stroke="none"
      />
    </Svg>
  )
}

export function IconMoon(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M 18.79 14.35 A 7.57 7.57 0 0 1 9.65 5.21 a 7.57 7.57 0 1 0 9.14 9.14 z" />
    </Svg>
  )
}

export function IconReadMode(props: IconProps) {
  return <CraftIcon set={CRAFT} name="view-reading" size={props.size ?? 24} />
}

export function IconOutlineView(props: IconProps) {
  return <CraftIcon set={CRAFT} name="view-outline" size={props.size ?? 24} />
}

export function IconRuler(props: IconProps) {
  return <CraftIcon set={CRAFT} name="ruler" size={props.size ?? 24} />
}

export function IconNavPane(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="4.49" y="5.65" width="15.02" height="12.71" rx="0.92" />
      <path d="M 9.69 5.65 v 12.71" />
      <path d="M 5.99 8.54 h 2.31 M 5.99 11.42 h 2.31 M 5.99 14.31 h 2.31" />
    </Svg>
  )
}

export function IconSplit(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="4.53" y="4.53" width="14.94" height="14.94" rx="1" />
      <path d="M 4.53 12 h 14.94" />
    </Svg>
  )
}

export function IconPrintLayout(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="6.8" y="4.49" width="10.4" height="15.02" rx="0.92" />
      <path d="M 9.11 7.96 h 5.78 M 9.11 10.85 h 5.78 M 9.11 13.73 h 5.78 M 9.11 16.62 h 3.47" />
    </Svg>
  )
}

export function IconWebLayout(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="4.49" y="5.65" width="15.02" height="12.71" rx="0.92" />
      <path d="M 4.49 8.54 h 15.02" />
      <path d="M 6.8 11.42 h 10.4 M 6.8 13.73 h 10.4 M 6.8 16.04 h 6.93" />
    </Svg>
  )
}

export function IconGridlines(props: IconProps) {
  return <CraftIcon set={CRAFT} name="gridlines" size={props.size ?? 24} />
}

export function IconNewWindow(props: IconProps) {
  return <CraftIcon set={CRAFT} name="new-window" size={props.size ?? 24} />
}

export function IconArrangeAll(props: IconProps) {
  return <CraftIcon set={CRAFT} name="arrange" size={props.size ?? 24} />
}

export function IconSwitchWindows(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="4.49" y="9.11" width="10.4" height="9.24" rx="0.92" />
      <path d="M 8.54 9.11 v -2.31 a 1.16 1.16 0 0 1 1.16 -1.15 h 8.66 a 1.16 1.16 0 0 1 1.16 1.16 v 8.09 a 1.16 1.16 0 0 1 -1.15 1.16 h -3.46" />
    </Svg>
  )
}

export function IconPosition(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="4.49" y="4.49" width="15.02" height="15.02" rx="1.16" />
      <rect x="8.54" y="8.54" width="6.93" height="6.93" />
    </Svg>
  )
}

export function IconWrapText(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="4.49" y="8.54" width="6.93" height="6.93" />
      <path d="M 13.73 5.07 h 5.78 M 13.73 8.54 h 5.78 M 13.73 12 h 5.78 M 13.73 15.47 h 5.78 M 4.49 18.93 h 15.02 M 4.49 5.07 h 6.93" />
    </Svg>
  )
}

export function IconDoc(props: IconProps) {
  return (
    <Svg {...props}>
      {PAGE}
      <path d="M15 2.25 V6 h3.75" />
      <path d="M8.25 9.75 h7.5 M8.25 13.5 h7.5 M8.25 17.25 h5.25" />
    </Svg>
  )
}

/* ---------- AI panel ---------- */

export function IconSend(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M 4.52 12 19.48 5.03 15.87 18.97 11.48 14.06 z" />
      <path d="M 11.48 14.06 19.48 5.03" />
    </Svg>
  )
}

export function IconStop(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="5" y="5" width="14" height="14" rx="2.625" fill="currentColor" stroke="none" />
    </Svg>
  )
}

export function IconGear(props: IconProps) {
  return (
    <Svg {...props}>
      <circle cx="12" cy="12" r="2.67" />
      <path d="M 12 4.47 v 2.43 M 12 17.1 v 2.43 M 19.53 12 h -2.43 M 6.9 12 h -2.43 M 17.35 6.65 l -1.7 1.7 M 8.36 15.65 l -1.7 1.7 M 17.35 17.35 15.65 15.65 M 8.36 8.36 6.65 6.65" />
    </Svg>
  )
}

/** collapse the right sidebar: panel outline + arrow pushing into it */
/** Collapse glyph for RIGHT-docked panes (Format/Animation/Comments) — exact mirror of
 *  IconSidebarCollapseLeft so both sides share the Sheets-parity look (16-canvas,
 *  1.2/1.3 stroke), self-contained for the same pinned-stroke reason. */
export function IconSidebarCollapse({ size = 24 }: IconProps) {
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
      <path d="M10.5 2.5v11" />
      <path d="M3.5 8h4.4M6.2 5.9 8.3 8l-2.1 2.1" strokeWidth="1.3" strokeLinejoin="round" />
    </svg>
  )
}

/** Mirror of IconSidebarCollapse for the LEFT-docked AI panel.
 *  Sheets-parity glyph (16-canvas, 1.2/1.3 stroke), self-contained so the shared
 *  Svg wrapper's 24-canvas pinned stroke doesn't alter its weight. */
export function IconSidebarCollapseLeft({ size = 24 }: IconProps) {
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
      <circle cx="12" cy="12" r="7.47" />
      <path d="M 12 8.02 V 12 l 2.86 1.99" />
    </Svg>
  )
}

export function IconPaperclip(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M 18.75 10.92 12.27 17.4 a 4.59 4.59 0 0 1 -6.48 -6.48 l 6.75 -6.75 a 3.11 3.11 0 0 1 4.32 4.32 l -6.75 6.75 a 1.49 1.49 0 0 1 -2.16 -2.16 l 6.21 -6.21" />
    </Svg>
  )
}

export function IconNewChat(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M 19.01 10.98 v -3.82 A 2.17 2.17 0 0 0 16.85 4.99 H 7.15 a 2.17 2.17 0 0 0 -2.17 2.17 v 7.78 a 2.17 2.17 0 0 0 2.17 2.17 h 1.4 v 2.55 l 3.32 -2.55 h 1.66" />
      <path d="M 17.36 13.79 v 5.1 M 14.81 16.34 h 5.1" />
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
  return <CraftIcon set={CRAFT} name="save" size={props.size ?? 24} />
}

export function IconUndo(props: IconProps) {
  return <CraftIcon set={CRAFT} name="undo" size={props.size ?? 24} />
}

export function IconRedo(props: IconProps) {
  return <CraftIcon set={CRAFT} name="redo" size={props.size ?? 24} />
}

export function IconCursor(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M 5.97 4.52 18.03 12.8 l -5.12 1.21 L 10.49 19.43 5.97 4.52 Z" />
    </Svg>
  )
}

export function IconPen(props: IconProps) {
  return <CraftIcon set={CRAFT} name="pen" size={props.size ?? 24} />
}

export function IconHighlighterPen(props: IconProps) {
  return <CraftIcon set={CRAFT} name="highlighter-pen" size={props.size ?? 24} />
}

export function IconEraser(props: IconProps) {
  return <CraftIcon set={CRAFT} name="eraser" size={props.size ?? 24} />
}

export function IconTextBox(props: IconProps) {
  return <CraftIcon set={CRAFT} name="text-box" size={props.size ?? 24} />
}

/** New slide (MS-style): slide frame with a title band and a split content area */
export function IconNewSlide(props: IconProps) {
  return <CraftIcon set={CRAFT} name="new-slide" size={props.size ?? 24} />
}

/** Section: divider + disclosure triangle, slide thumbnails grouped beneath */
export function IconSection(props: IconProps) {
  return <CraftIcon set={CRAFT} name="section" size={props.size ?? 24} />
}

export function IconRect(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="4.49" y="6.8" width="15.02" height="10.4" />
    </Svg>
  )
}

export function IconRoundRect(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="4.49" y="6.8" width="15.02" height="10.4" rx="2.89" />
    </Svg>
  )
}

export function IconEllipse(props: IconProps) {
  return (
    <Svg {...props}>
      <ellipse cx="12" cy="12" rx="7.51" ry="5.2" />
    </Svg>
  )
}

/** Slide show: from beginning (slide open at the bottom-right, its bottom edge
 * running out as an arrow into a play triangle — loop back to slide 1 and play) */
export function IconPlayFromStart(props: IconProps) {
  return <CraftIcon set={CRAFT} name="play-from-start" size={props.size ?? 24} />
}

/** Slide show: from current slide (screen + play triangle) */
export function IconPlayCurrent(props: IconProps) {
  return <CraftIcon set={CRAFT} name="play-from-current" size={props.size ?? 24} />
}

/** Status-bar play: solid triangle in a rounded square. The frame is nearly
 *  the full viewBox — the monitor-and-stand ribbon glyph reads too small at
 *  status-bar sizes */
export function IconPlayBoxed(props: IconProps) {
  return <CraftIcon set={CRAFT} name="play" size={props.size ?? 24} />
}

/** Status-bar notes toggle: a note page with text lines (near-full viewBox,
 *  same rationale as IconPlayBoxed at status-bar sizes) */
export function IconNotes(props: IconProps) {
  return <CraftIcon set={CRAFT} name="notes" size={props.size ?? 24} />
}

/** Presenter view: main screen + small presenter screen */
export function IconPresenterView(props: IconProps) {
  return <CraftIcon set={CRAFT} name="presenter-view" size={props.size ?? 24} />
}

/** Custom show: screen + gear dots */
export function IconCustomShow(props: IconProps) {
  return <CraftIcon set={CRAFT} name="custom-show" size={props.size ?? 24} />
}

/** Set up slide show: screen + wrench slash */
export function IconSetupShow(props: IconProps) {
  return <CraftIcon set={CRAFT} name="set-up-show" size={props.size ?? 24} />
}

/** Hide slide: slide + slash */
export function IconHideSlide(props: IconProps) {
  return <CraftIcon set={CRAFT} name="hide-slide" size={props.size ?? 24} />
}

/** Rehearse timings: stopwatch */
export function IconRehearse(props: IconProps) {
  return <CraftIcon set={CRAFT} name="rehearse-timings" size={props.size ?? 24} />
}

/** Record slide show: recording dot */
export function IconRecord(props: IconProps) {
  return <CraftIcon set={CRAFT} name="record" size={props.size ?? 24} />
}

/** Chart: bar chart */
export function IconChart(props: IconProps) {
  return <CraftIcon set={CRAFT} name="chart" size={props.size ?? 24} />
}

/** SmartArt: connected nodes */
export function IconSmartArt(props: IconProps) {
  return <CraftIcon set={CRAFT} name="smart-art" size={props.size ?? 24} />
}

/** WordArt: outlined block-letter A (matches docs IconWordArt, scaled 16→24 canvas) */
export function IconWordArt(props: IconProps) {
  return <CraftIcon set={CRAFT} name="word-art" size={props.size ?? 24} />
}

/** Icon gallery: smiley */
export function IconIconLib(props: IconProps) {
  return <CraftIcon set={CRAFT} name="icons" size={props.size ?? 24} />
}

/** Video: film play */
export function IconVideo(props: IconProps) {
  return <CraftIcon set={CRAFT} name="video" size={props.size ?? 24} />
}

/** Audio: speaker */
export function IconAudio(props: IconProps) {
  return <CraftIcon set={CRAFT} name="audio" size={props.size ?? 24} />
}

/** Screen recording: screen + record dot */
export function IconScreenRec(props: IconProps) {
  return <CraftIcon set={CRAFT} name="screen-recording" size={props.size ?? 24} />
}

/** 3D model: cube */
export function Icon3d(props: IconProps) {
  return <CraftIcon set={CRAFT} name="models-3d" size={props.size ?? 24} />
}

/** Zoom link: jump arrow + page */
export function IconZoomJump(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="4.53" y="4.53" width="9.96" height="7.47" rx="1" />
      <rect x="9.51" y="12" width="9.96" height="7.47" rx="1" />
      <path d="M 14.49 8.27 l 3.74 0 M 18.22 8.27 l -1.74 -1.74 M 18.22 8.27 l -1.74 1.74" />
    </Svg>
  )
}

/** Date-time: calendar + hour hand */
export function IconDateTime(props: IconProps) {
  return <CraftIcon set={CRAFT} name="date-time" size={props.size ?? 24} />
}

/** Object align/distribute (distinct from paragraph text alignment IconAlign*): color blocks + alignment baseline */
export function IconObjAlignLeft(props: IconProps) {
  return <CraftIcon set={CRAFT} name="align-obj-left" size={props.size ?? 24} />
}

export function IconObjAlignCenterH(props: IconProps) {
  return <CraftIcon set={CRAFT} name="align-obj-center" size={props.size ?? 24} />
}

export function IconObjAlignRight(props: IconProps) {
  return <CraftIcon set={CRAFT} name="align-obj-right" size={props.size ?? 24} />
}

export function IconObjAlignTop(props: IconProps) {
  return <CraftIcon set={CRAFT} name="align-obj-top" size={props.size ?? 24} />
}

export function IconObjAlignMiddle(props: IconProps) {
  return <CraftIcon set={CRAFT} name="align-obj-middle" size={props.size ?? 24} />
}

export function IconObjAlignBottom(props: IconProps) {
  return <CraftIcon set={CRAFT} name="align-obj-bottom" size={props.size ?? 24} />
}

export function IconObjDistributeH(props: IconProps) {
  return <CraftIcon set={CRAFT} name="distribute-h" size={props.size ?? 24} />
}

export function IconObjDistributeV(props: IconProps) {
  return <CraftIcon set={CRAFT} name="distribute-v" size={props.size ?? 24} />
}

export function IconObjFlipH(props: IconProps) {
  return <CraftIcon set={CRAFT} name="flip-h" size={props.size ?? 24} />
}

export function IconObjFlipV(props: IconProps) {
  return <CraftIcon set={CRAFT} name="flip-v" size={props.size ?? 24} />
}

export function IconSwitchRowCol(props: IconProps) {
  return <CraftIcon set={CRAFT} name="switch-row-col" size={props.size ?? 24} />
}

export function IconEditChartData(props: IconProps) {
  return <CraftIcon set={CRAFT} name="edit-data" size={props.size ?? 24} />
}

export function IconChangeChartType(props: IconProps) {
  return <CraftIcon set={CRAFT} name="change-chart-type" size={props.size ?? 24} />
}

// ── Animation effect icons (Animations tab; color comes from .rb-anim-* via currentColor) ──

const ANIM_EFFECT_BODIES: Record<AnimEffectKind, ReactNode> = {
  // entrance
  appear: (
    <>
      <rect x="4" y="5" width="16" height="14" rx="1.5" strokeDasharray="3 2.2" />
      <rect x="8.2" y="9.2" width="7.6" height="5.6" fill="currentColor" stroke="none" />
    </>
  ),
  fade: (
    <>
      <rect
        x="4.5"
        y="6.5"
        width="4.6"
        height="11"
        fill="currentColor"
        stroke="none"
        opacity="0.2"
      />
      <rect
        x="9.7"
        y="6.5"
        width="4.6"
        height="11"
        fill="currentColor"
        stroke="none"
        opacity="0.5"
      />
      <rect x="14.9" y="6.5" width="4.6" height="11" fill="currentColor" stroke="none" />
    </>
  ),
  flyIn: (
    <>
      <rect x="5.5" y="3.5" width="13" height="8" rx="1" />
      <path d="M12 21 V14.5 M8.8 17.4 L12 14.2 l3.2 3.2" />
    </>
  ),
  wipe: (
    <>
      <rect x="3.5" y="6" width="17" height="12" rx="1" />
      <rect x="3.5" y="6" width="8.2" height="12" fill="currentColor" stroke="none" />
      <path d="M13.3 12 h4.8 M15.9 9.8 l2.2 2.2 -2.2 2.2" />
    </>
  ),
  wipeDown: (
    <>
      <rect x="3.5" y="4.5" width="17" height="15" rx="1" />
      <rect x="3.5" y="4.5" width="17" height="6.4" fill="currentColor" stroke="none" />
      <path d="M12 12.6 V17.3 M9.9 15.3 L12 17.4 l2.1 -2.1" />
    </>
  ),
  splitIn: (
    <>
      <rect x="3.5" y="6" width="17" height="12" rx="1" />
      <path d="M5.8 12 h4.4 M8.3 9.9 l2.1 2.1 -2.1 2.1 M18.2 12 h-4.4 M15.7 9.9 l-2.1 2.1 2.1 2.1" />
    </>
  ),
  bounce: (
    <>
      <path d="M3.5 19.5 C6 7.5, 8.5 7.5, 11 19.5 C12.8 13, 14.6 13, 16.4 19.5" />
      <circle cx="19.3" cy="17.8" r="1.9" fill="currentColor" stroke="none" />
    </>
  ),
  flipIn: (
    <>
      <rect x="4" y="6" width="7" height="12" />
      <path d="M13.5 4.2 L20 6.4 V17.6 L13.5 19.8 Z" />
    </>
  ),
  zoom: (
    <>
      <rect x="9.2" y="9.2" width="5.6" height="5.6" fill="currentColor" stroke="none" />
      <path d="M8 8 L4.5 4.5 M4.5 8 V4.5 H8" />
      <path d="M16 16 L19.5 19.5 M19.5 16 V19.5 H16" />
    </>
  ),
  // emphasis
  pulse: (
    <>
      <circle cx="12" cy="12" r="2.2" fill="currentColor" stroke="none" />
      <circle cx="12" cy="12" r="5.6" opacity="0.65" />
      <circle cx="12" cy="12" r="9" opacity="0.35" />
    </>
  ),
  spin: (
    <>
      <path d="M12 5.4 A 6.6 6.6 0 1 1 5.4 12" />
      <path d="M3.3 13.9 L5.4 11.5 7.7 13.8" />
    </>
  ),
  grow: (
    <>
      <rect x="7.5" y="7.5" width="9" height="9" />
      <path d="M4.5 19.5 L19.5 4.5 M4.5 15.2 V19.5 H8.8 M19.5 8.8 V4.5 H15.2" />
    </>
  ),
  teeter: (
    <>
      <rect x="5" y="9.5" width="14" height="9" rx="1" transform="rotate(-8 12 14)" />
      <path d="M6 6.6 A 7.5 4.5 0 0 1 18 6.6" />
      <path d="M5.2 4.2 L6 6.8 8.6 6.1 M18.8 4.2 L18 6.8 15.4 6.1" />
    </>
  ),
  // exit
  disappear: <rect x="4" y="5" width="16" height="14" rx="1.5" strokeDasharray="3 2.2" />,
  fadeOut: (
    <>
      <rect x="4.5" y="6.5" width="4.6" height="11" fill="currentColor" stroke="none" />
      <rect
        x="9.7"
        y="6.5"
        width="4.6"
        height="11"
        fill="currentColor"
        stroke="none"
        opacity="0.5"
      />
      <rect
        x="14.9"
        y="6.5"
        width="4.6"
        height="11"
        fill="currentColor"
        stroke="none"
        opacity="0.2"
      />
    </>
  ),
  flyOut: (
    <>
      <rect x="5.5" y="3.5" width="13" height="8" rx="1" />
      <path d="M12 14.2 V20.7 M8.8 17.5 L12 20.7 l3.2 -3.2" />
    </>
  ),
  wipeOut: (
    <>
      <rect x="3.5" y="6" width="17" height="12" rx="1" />
      <rect x="12.3" y="6" width="8.2" height="12" fill="currentColor" stroke="none" />
      <path d="M10.7 12 H5.9 M8.1 9.8 L5.9 12 l2.2 2.2" />
    </>
  ),
  shrink: (
    <>
      <rect
        x="9.6"
        y="9.6"
        width="4.8"
        height="4.8"
        fill="currentColor"
        stroke="none"
        transform="rotate(15 12 12)"
      />
      <path d="M4.5 4.5 L8.3 8.3 M8.3 5.1 V8.3 H5.1" />
      <path d="M19.5 19.5 L15.7 15.7 M15.7 18.9 V15.7 H18.9" />
    </>
  ),
  zoomOut: (
    <>
      <rect x="3.5" y="3.5" width="17" height="17" rx="1" strokeDasharray="2.6 2" />
      <path d="M5.6 5.6 L9.4 9.4 M9.4 6.2 V9.4 H6.2" />
      <path d="M18.4 18.4 L14.6 14.6 M14.6 17.8 V14.6 H17.8" />
    </>
  ),
  motionPath: (
    <>
      <path d="M4.5 6 C10 3, 14 9, 18.6 15.8" strokeDasharray="3 2.2" />
      <path d="M18.9 10.9 l0.5 5.4 -5.4 -0.6" />
    </>
  ),
}

export function AnimEffectIcon({ kind, size }: { kind: AnimEffectKind; size?: number }) {
  return <Svg size={size}>{ANIM_EFFECT_BODIES[kind]}</Svg>
}

/* ---------- transition gallery (drawn to the shared icon standard, replacing
   the old unicode text glyphs whose size and weight came from the font) ---- */

export function IconTransNone(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="4.75" y="6" width="14.5" height="12" rx="1.5" />
      <path d="M6.5 16.5 17.5 7.5" />
    </Svg>
  )
}

export function IconTransMorph(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M4.75 9h11.5M13.5 6.25 16.25 9l-2.75 2.75" />
      <path d="M19.25 15H7.75M10.5 17.75 7.75 15l2.75-2.75" />
    </Svg>
  )
}

export function IconTransFade(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="4.75" y="4.75" width="11" height="9.5" rx="1.4" strokeDasharray="2.4 2.2" />
      <rect x="8.25" y="9.75" width="11" height="9.5" rx="1.4" />
    </Svg>
  )
}

export function IconTransPush(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="4.75" y="6" width="14.5" height="12" rx="1.5" />
      <path d="M12 15.5v-6M9.25 12.25 12 9.5l2.75 2.75" />
    </Svg>
  )
}

export function IconTransWipe(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="4.75" y="6" width="14.5" height="12" rx="1.5" />
      <path d="M7.5 12h6M11.25 9.75 13.5 12l-2.25 2.25" />
    </Svg>
  )
}

export function IconTransSplit(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="4.75" y="6" width="14.5" height="12" rx="1.5" />
      <path d="M12 8.5v7" />
      <path d="M9.5 12H7M8.25 10.75 7 12l1.25 1.25M14.5 12H17M15.75 10.75 17 12l-1.25 1.25" />
    </Svg>
  )
}

export function IconTransCircle(props: IconProps) {
  return (
    <Svg {...props}>
      <circle cx="12" cy="12" r="7.25" />
      <circle cx="12" cy="12" r="3.25" />
    </Svg>
  )
}

export function IconTransCover(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M14.5 8.25V6.25a1.5 1.5 0 0 0-1.5-1.5H6.25a1.5 1.5 0 0 0-1.5 1.5V13a1.5 1.5 0 0 0 1.5 1.5h2" />
      <rect x="9.75" y="9.75" width="9.5" height="9.5" rx="1.5" />
    </Svg>
  )
}

export function IconTransPull(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="4.75" y="4.75" width="9.5" height="9.5" rx="1.5" />
      <path d="M9.5 15.75v2a1.5 1.5 0 0 0 1.5 1.5h6.75a1.5 1.5 0 0 0 1.5-1.5V11a1.5 1.5 0 0 0-1.5-1.5h-2" />
    </Svg>
  )
}

export function IconTransDissolve(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="4.75" y="6" width="14.5" height="12" rx="1.5" />
      <circle cx="9" cy="10" r="0.8" fill="currentColor" stroke="none" />
      <circle cx="13.75" cy="9.25" r="0.8" fill="currentColor" stroke="none" />
      <circle cx="15.75" cy="13.25" r="0.8" fill="currentColor" stroke="none" />
      <circle cx="8.25" cy="14" r="0.8" fill="currentColor" stroke="none" />
      <circle cx="12" cy="14.75" r="0.8" fill="currentColor" stroke="none" />
    </Svg>
  )
}

export function IconTransZoom(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M14 4.75h5.25V10M19 5 13.75 10.25" />
      <path d="M10 19.25H4.75V14M5 19 10.25 13.75" />
    </Svg>
  )
}

export function IconTransRandom(props: IconProps) {
  return (
    <Svg {...props}>
      <rect x="5" y="5" width="14" height="14" rx="2.5" />
      <circle cx="9.25" cy="9.25" r="1" fill="currentColor" stroke="none" />
      <circle cx="14.75" cy="9.25" r="1" fill="currentColor" stroke="none" />
      <circle cx="12" cy="12" r="1" fill="currentColor" stroke="none" />
      <circle cx="9.25" cy="14.75" r="1" fill="currentColor" stroke="none" />
      <circle cx="14.75" cy="14.75" r="1" fill="currentColor" stroke="none" />
    </Svg>
  )
}

/* ---------- animation gallery chrome (star / none / motion paths) -------- */

export function IconAnimStar(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="m12 4.5 2.15 4.9 5.35.5-4 3.6 1.15 5.25L12 16l-4.65 2.75 1.15-5.25-4-3.6 5.35-.5Z" />
    </Svg>
  )
}

export function IconAnimNone(props: IconProps) {
  return (
    <Svg {...props}>
      <circle cx="12" cy="12" r="7.25" />
      <path d="M6.9 17.1 17.1 6.9" />
    </Svg>
  )
}

export function IconPathRight(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M4.75 12H18M14.75 8.75 18 12l-3.25 3.25" />
    </Svg>
  )
}

export function IconPathDown(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M12 4.75V18M8.75 14.75 12 18l3.25-3.25" />
    </Svg>
  )
}

export function IconPathDiagonal(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M5.5 5.5 18 18M18 13.4V18h-4.6" />
    </Svg>
  )
}

export function IconPathCircle(props: IconProps) {
  return (
    <Svg {...props}>
      <circle cx="12" cy="12" r="6.75" />
      <path d="M15.5 3.9 12.7 5.2l1.3 2.75" />
    </Svg>
  )
}

export function IconPathZigzag(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="M4.75 15.5 8.75 9.5l3.75 5.5 4.25-6.25" />
      <path d="M17.5 12.9V8.25h-4.6" />
    </Svg>
  )
}

export function IconCrop(props: IconProps) {
  return <CraftIcon set={CRAFT} name="crop" size={props.size ?? 24} />
}

export function IconNoneX(props: IconProps) {
  return (
    <Svg {...props}>
      <path d="m6.5 6.5 11 11M17.5 6.5l-11 11" />
    </Svg>
  )
}

/** AI feature glyphs shared by the ribbon Home tab and the canvas AI bar.
 * Fixed 1.5-unit stroke (not pinnedStroke) to keep the ribbon rendering,
 * where CSS sizes them; `size` is for other hosts. */
function AiFeatureSvg({ size, children }: IconProps & { children: ReactNode }) {
  return (
    <svg
      width={size}
      height={size}
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth="1.5"
      strokeLinecap="round"
      strokeLinejoin="round"
      aria-hidden
    >
      {children}
    </svg>
  )
}

export function IconAiBeautify(props: IconProps) {
  return (
    <AiFeatureSvg {...props}>
      <path d="M10.4948 15.0904C9.91037 14.5539 9.12711 14.5196 9.12711 14.5196C9.12711 14.5196 8.37299 14.4408 7.8691 14.7887C7.03613 15.3629 7.02071 16.5283 7.01385 16.7014C6.995 17.2293 7.03785 17.8652 6.71563 18.4239C6.24088 19.2466 5 19.5894 5 19.5894C5 19.5894 5.21081 19.8448 5.51246 19.9939C5.51246 19.9939 6.47226 20.5423 8.45012 20.0727C8.99 19.9442 9.58816 19.7796 10.0732 19.4763C10.6662 19.1061 11.105 18.5799 11.2232 18.3194C11.7374 17.2053 11.4752 15.9902 10.4948 15.0904Z" />
      <path d="M18.9872 4.57628C18.9772 4.39067 18.7726 4.29501 18.7726 4.29501C18.7726 4.29501 18.5538 4.25444 18.3736 4.34695C18.0654 4.50498 17.6181 4.70138 15.4274 7.19175C13.9033 8.92525 12.0146 11.2553 10.9207 12.5777C10.1021 13.5619 10.0281 14.2654 10.1806 14.6721C10.2108 14.7514 10.5825 15.0484 10.9207 15.2661C11.1663 15.4249 11.3406 15.5257 11.6199 15.5437C12.0717 15.5739 12.5127 15.2661 12.5127 15.2661C12.5127 15.2661 12.7483 15.0623 13.1086 14.587C13.8466 13.6144 15.1527 11.7581 16.2918 9.96583C17.7356 7.70256 18.8308 5.34067 18.8308 5.34067C18.8308 5.34067 19.0033 4.88149 18.9872 4.57628Z" />
      <path d="M7 4L7.22106 4.59745C7.51094 5.38087 7.65589 5.77259 7.94166 6.05833C8.22743 6.34409 8.61914 6.48903 9.40257 6.77893L10 7L9.40257 7.22107C8.61914 7.51097 8.22743 7.65592 7.94166 7.94167C7.65589 8.22741 7.51094 8.61913 7.22106 9.40255L7 10L6.77894 9.40255C6.48906 8.61913 6.34411 8.22741 6.05834 7.94167C5.77257 7.65592 5.38086 7.51097 4.59743 7.22107L4 7L4.59743 6.77893C5.38086 6.48903 5.77257 6.34409 6.05834 6.05833C6.34411 5.77259 6.48906 5.38087 6.77894 4.59745L7 4Z" />
      <path d="M18 16L18.1474 16.3983C18.3406 16.9206 18.4373 17.1817 18.6278 17.3722C18.8183 17.5627 19.0794 17.6594 19.6017 17.8526L20 18L19.6017 18.1474C19.0794 18.3406 18.8183 18.4373 18.6278 18.6278C18.4373 18.8183 18.3406 19.0794 18.1474 19.6017L18 20L17.8526 19.6017C17.6594 19.0794 17.5627 18.8183 17.3722 18.6278C17.1817 18.4373 16.9206 18.3406 16.3983 18.1474L16 18L16.3983 17.8526C16.9206 17.6594 17.1817 17.5627 17.3722 17.3722C17.5627 17.1817 17.6594 16.9206 17.8526 16.3983L18 16Z" />
    </AiFeatureSvg>
  )
}

export function IconAiFactCheck(props: IconProps) {
  return (
    <AiFeatureSvg {...props}>
      <circle cx="11.9997" cy="12.0004" r="9.53571" />
      <path
        d="M18.2407 8.20112C18.5754 7.86675 19.118 7.86668 19.4526 8.20112C19.7873 8.53587 19.7873 9.07931 19.4526 9.41402L14.0727 14.7929C13.9119 14.9535 13.6936 15.044 13.4663 15.0439C13.2392 15.0437 13.0213 14.9535 12.8608 14.7929L11.3921 13.3222C11.1247 13.0544 11.0714 12.6529 11.2319 12.332L8.70555 14.8593C8.37087 15.194 7.82741 15.1939 7.49266 14.8593L4.53563 11.9023C4.20123 11.5676 4.20121 11.025 4.53563 10.6904C4.87039 10.3559 5.41388 10.3557 5.74852 10.6904L8.0991 13.041L12.8725 8.26753C13.2072 7.93333 13.7498 7.93331 14.0845 8.26753C14.4189 8.60218 14.4188 9.14572 14.0845 9.48042L11.6147 11.9492C11.936 11.7883 12.3371 11.8423 12.605 12.1103L13.4673 12.9736L18.2407 8.20112Z"
        fill="currentColor"
        stroke="none"
      />
    </AiFeatureSvg>
  )
}

export function IconAiImage(props: IconProps) {
  return (
    <AiFeatureSvg {...props}>
      <path d="M7.55042 9.62398C8.33741 9.62398 8.97538 8.986 8.97538 8.19902C8.97538 7.41203 8.33741 6.77405 7.55042 6.77405C6.76343 6.77405 6.12545 7.41203 6.12545 8.19902C6.12545 8.986 6.76343 9.62398 7.55042 9.62398Z" />
      <path d="M20.8474 11.05C20.8496 11.4966 20.8496 11.9708 20.8496 12.475C20.8496 16.7293 20.8496 18.8565 19.528 20.1782C18.2063 21.4998 16.0791 21.4998 11.8248 21.4998C7.57047 21.4998 5.44331 21.4998 4.12165 20.1782C2.8 18.8565 2.8 16.7293 2.8 12.475C2.8 8.22066 2.8 6.0935 4.12165 4.77184C5.44331 3.4502 7.57047 3.4502 11.8248 3.4502C12.3289 3.4502 12.8032 3.4502 13.2498 3.45239" />
      <path d="M18.4756 2.49915L18.7206 3.16131C19.0419 4.02958 19.2026 4.46372 19.5193 4.78041C19.836 5.09712 20.2701 5.25776 21.1384 5.57905L21.8006 5.82407L21.1384 6.06909C20.2701 6.39038 19.836 6.55103 19.5193 6.86772C19.2026 7.18442 19.0419 7.61856 18.7206 8.48683L18.4756 9.14899L18.2306 8.48683C17.9093 7.61856 17.7487 7.18442 17.432 6.86772C17.1153 6.55103 16.6811 6.39038 15.8128 6.06909L15.1507 5.82407L15.8128 5.57905C16.6811 5.25776 17.1153 5.09712 17.432 4.78041C17.7487 4.46372 17.9093 4.02958 18.2306 3.16131L18.4756 2.49915Z" />
      <path d="M5.17599 21.0236C9.32974 16.06 13.9862 9.51375 20.8483 13.9391" />
    </AiFeatureSvg>
  )
}

export function IconRotateRight(props: IconProps) {
  return <CraftIcon set={CRAFT} name="rotate-right" size={props.size ?? 24} />
}

export function IconRotateLeft(props: IconProps) {
  return <CraftIcon set={CRAFT} name="rotate-left" size={props.size ?? 24} />
}

export function IconReplacePicture(props: IconProps) {
  return <CraftIcon set={CRAFT} name="change-picture" size={props.size ?? 24} />
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
