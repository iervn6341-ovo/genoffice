/**
 * 3.1/3.2 Konva canvas — renders one RenderSlide + selection/transform + text-edit triggering.
 */
import React, {
  useEffect,
  useImperativeHandle,
  useLayoutEffect,
  useMemo,
  useRef,
  useState,
  useSyncExternalStore,
} from 'react'
import {
  Stage,
  Layer,
  Rect,
  Group,
  Transformer,
  Line,
  Arrow,
  Text,
  Circle,
  Path,
} from 'react-konva'
import type Konva from 'konva'
import {
  createSelectionDrag,
  selectionMoveRequests,
  stopSelectionDrags,
  type SelectionDrag,
} from './selection-drag'
import type {
  RenderSlide,
  RenderNode,
  RenderFill,
  ShapeRenderNode,
  TableRenderNode,
  ChartRenderNode,
  PictureRenderNode,
  GroupRenderNode,
} from '@genoffice/pptx-render'
import { boxPivotProps, fillToKonva, isEditableText } from './konva-adapter'
import {
  TABLE_BORDER_GRIP_PX,
  tableCellAtPoint,
  tableCellInterior,
  tableCellOverlayBox,
  tableLocalPointFromStage,
} from './table-hit'
import { isPromptPlaceholder, textHitAtPoint } from './text-hit-area'
import {
  clampAnchorPos,
  colDragRange,
  ALL_ANCHORS,
  anchorsForScreenSize,
  clampTransformBox,
  resolveColDrag,
  resolveRowDrag,
  rowDragFloor,
  type AnchorRect,
  type MinSize,
} from './min-size'
import { isConnectorNode, minSizeOfNodes, selectionBounds } from './selection-min'
import { emitNodeGesture } from './node-gesture'
import type { GridChange } from './table-grid-preview'
import { unstretchGlyphLayer } from './text-unstretch'
import { centerShift, contentBoundHeight, drawnTableHeight } from './resize-geometry'
import type { EditCaret } from './action-context'
import {
  computeSnap,
  computeSpacingSnap,
  type SnapTarget,
  type Guide,
  type SpacingIndicator,
} from './snap'
import { NodeBody, StaticNode } from './NodeBody'
import { ZOOM_PREVIEW_EVENT } from './zoom-preview'
import { useI18n } from './i18n/locale'
import {
  defaultDrawSize,
  isLineDrawKind,
  isStraightLineKind,
  resolveDrawRect,
  type DrawRect,
} from './draw-shape'
import { shapePreviewPath } from './components/gallery-previews'
import { adjustHandleSpecs } from './adjust-handles'

/** Luminance (0..1) + alpha (0..1) pair — the unit the background-darkness composite works in. */
type LumAlpha = { lum: number; alpha: number }

/** Perceived luminance and alpha of a CSS color (#rgb/#rrggbb[aa]/rgb[a]()); null when
 * unparseable. Alpha must survive parsing: a transparent full-page overlay otherwise
 * reads as its base color and flips the chrome to white on a white slide. */
function colorLumAlpha(color: string): LumAlpha | null {
  const c = color.trim()
  let r: number, g: number, b: number
  let a = 1
  const hex6 = /^#([0-9a-f]{6})([0-9a-f]{2})?/i.exec(c)
  const hex3 = /^#([0-9a-f]{3})$/i.exec(c)
  const rgb = /^rgba?\(\s*(\d+)[,\s]+(\d+)[,\s]+(\d+)(?:[,\s/]+([\d.]+%?))?/i.exec(c)
  if (hex6) {
    r = parseInt(hex6[1]!.slice(0, 2), 16)
    g = parseInt(hex6[1]!.slice(2, 4), 16)
    b = parseInt(hex6[1]!.slice(4, 6), 16)
    if (hex6[2]) a = parseInt(hex6[2], 16) / 255
  } else if (hex3) {
    r = parseInt(hex3[1]![0]! + hex3[1]![0]!, 16)
    g = parseInt(hex3[1]![1]! + hex3[1]![1]!, 16)
    b = parseInt(hex3[1]![2]! + hex3[1]![2]!, 16)
  } else if (rgb) {
    r = Number(rgb[1])
    g = Number(rgb[2])
    b = Number(rgb[3])
    if (rgb[4]) a = rgb[4].endsWith('%') ? Number(rgb[4].slice(0, -1)) / 100 : Number(rgb[4])
  } else {
    return null
  }
  if (!Number.isFinite(a)) a = 1
  return {
    lum: (0.299 * r + 0.587 * g + 0.114 * b) / 255,
    alpha: Math.max(0, Math.min(1, a)),
  }
}

/** Alpha-weighted average luminance + mean coverage of an image (8×8 downsample), cached
 * per element; null while not loaded. Transparent pixels must not count as black — a
 * mostly-transparent decor PNG is (near-)invisible, not a dark background. */
const imageLumCache = new WeakMap<HTMLImageElement, LumAlpha | null>()
function imageLumAlpha(img: HTMLImageElement | undefined): LumAlpha | null {
  if (!img || !img.complete || !img.naturalWidth) return null
  const hit = imageLumCache.get(img)
  if (hit !== undefined) return hit
  let out: LumAlpha | null
  try {
    const c = document.createElement('canvas')
    c.width = 8
    c.height = 8
    const ctx = c.getContext('2d')!
    ctx.drawImage(img, 0, 0, 8, 8)
    const d = ctx.getImageData(0, 0, 8, 8).data
    let sum = 0
    let cover = 0
    for (let i = 0; i < d.length; i += 4) {
      const a = d[i + 3]! / 255
      sum += ((0.299 * d[i]! + 0.587 * d[i + 1]! + 0.114 * d[i + 2]!) / 255) * a
      cover += a
    }
    out = cover > 0 ? { lum: sum / cover, alpha: cover / (d.length / 4) } : { lum: 1, alpha: 0 }
  } catch {
    out = null // tainted canvas etc. — treat as unknown
  }
  imageLumCache.set(img, out)
  return out
}

/** Luminance/alpha of a render fill; gradients alpha-weight their stops, images sample the bitmap. */
function fillLumAlpha(fill: RenderFill, images: Map<string, HTMLImageElement>): LumAlpha | null {
  if (fill.kind === 'solid') return colorLumAlpha(fill.color)
  if (fill.kind === 'gradient') {
    const stops = fill.stops
      .map((s) => colorLumAlpha(s.color))
      .filter((v): v is LumAlpha => v != null)
    if (!stops.length) return null
    const cover = stops.reduce((acc, s) => acc + s.alpha, 0)
    return {
      lum: cover > 0 ? stops.reduce((acc, s) => acc + s.lum * s.alpha, 0) / cover : 1,
      alpha: cover / stops.length,
    }
  }
  if (fill.kind === 'image') {
    const la = imageLumAlpha(fill.dataUrl ? images.get(fill.dataUrl) : undefined)
    return la && { lum: la.lum, alpha: la.alpha * (fill.alpha ?? 1) }
  }
  return null
}

/** Whether the slide's effective background is dark (selection chrome flips to white on it).
 * The slide background, full-page master/layout decorations and background-flagged elements
 * composite bottom-up, each blended by its alpha over an assumed light page — so a
 * transparent or lightly-tinted overlay can't flip the verdict with its base color, while
 * an opaque or heavy dark scrim still does. Unknown layers (unloaded images, unparseable
 * colors) are skipped and keep whatever is below. */
function slideBackgroundIsDark(slide: RenderSlide, images: Map<string, HTMLImageElement>): boolean {
  let lum: number | null = null
  const blend = (la: LumAlpha | null) => {
    if (!la || la.alpha <= 0) return
    lum = la.alpha >= 1 ? la.lum : la.lum * la.alpha + (lum ?? 1) * (1 - la.alpha)
  }
  const tol = 2
  const coversPage = (n: RenderNode, ox: number, oy: number): boolean =>
    !n.box.rotationDeg &&
    ox + n.box.x <= tol &&
    oy + n.box.y <= tol &&
    ox + n.box.x + n.box.w >= slide.widthPx - tol &&
    oy + n.box.y + n.box.h >= slide.heightPx - tol
  // Group children are in group-local coords; a master plate nested in a decoration group
  // still paints the page, so walk into covering groups with the group's offset applied.
  const visit = (nodes: RenderNode[], ox: number, oy: number, inherited: boolean) => {
    for (const n of nodes) {
      const flagged = inherited || !!n.background || !!n.decoration
      if (!n.background && !(flagged && coversPage(n, ox, oy))) continue
      if (n.type === 'group') {
        if (!n.box.rotationDeg)
          visit((n as GroupRenderNode).children, ox + n.box.x, oy + n.box.y, true)
        continue
      }
      if (n.type === 'picture') {
        const pic = n as PictureRenderNode
        const la = imageLumAlpha(pic.dataUrl ? images.get(pic.dataUrl) : undefined)
        blend(la && { lum: la.lum, alpha: la.alpha * (pic.opacity ?? 1) })
      } else if (n.type === 'shape' || n.type === 'text') {
        blend(fillLumAlpha((n as ShapeRenderNode).fill, images))
      }
    }
  }
  blend(fillLumAlpha(slide.background, images))
  visit(slide.nodes, 0, 0, false)
  return (lum ?? 1) < 0.5
}

/** Whether the slide looks dark UNDER a specific region: composite the slide background
 * plus every node whose box overlaps the region, bottom-up in z-order, each blended by
 * alpha × the fraction of the region it covers. The page-average verdict fails the most
 * common template — a colored page with a white content plate: a text box ON the plate
 * needs dark chrome even though the page itself is dark. Rotation is
 * approximated by the unrotated box; tables blend their cell fills (the cell-edit overlay
 * samples one cell), charts their chart-space and plot-area fills; unknown layers
 * (unloaded images) are skipped and keep whatever is below. */
function regionIsDark(
  slide: RenderSlide,
  images: Map<string, HTMLImageElement>,
  region: { x: number; y: number; w: number; h: number },
): boolean {
  const area = region.w * region.h
  if (!(area > 0)) return slideBackgroundIsDark(slide, images)
  let lum: number | null = null
  const blend = (la: LumAlpha | null, cover: number) => {
    if (!la || la.alpha <= 0 || cover <= 0) return
    const a = Math.min(1, la.alpha * cover)
    lum = a >= 1 ? la.lum : la.lum * a + (lum ?? 1) * (1 - a)
  }
  const coverage = (x: number, y: number, w: number, h: number): number => {
    const ix = Math.max(0, Math.min(x + w, region.x + region.w) - Math.max(x, region.x))
    const iy = Math.max(0, Math.min(y + h, region.y + region.h) - Math.max(y, region.y))
    return (ix * iy) / area
  }
  // Group children are in group-local coords (same walk as slideBackgroundIsDark)
  const visit = (nodes: RenderNode[], ox: number, oy: number) => {
    for (const n of nodes) {
      const cover = coverage(ox + n.box.x, oy + n.box.y, n.box.w, n.box.h)
      // a rotated table's cells can lie outside its unrotated box; they gate themselves below
      if (cover <= 0 && n.type !== 'table') continue
      if (n.type === 'group') {
        if (!n.box.rotationDeg) visit((n as GroupRenderNode).children, ox + n.box.x, oy + n.box.y)
        continue
      }
      if (n.type === 'picture') {
        const pic = n as PictureRenderNode
        const la = imageLumAlpha(pic.dataUrl ? images.get(pic.dataUrl) : undefined)
        blend(la && { lum: la.lum, alpha: la.alpha * (pic.opacity ?? 1) }, cover)
      } else if (n.type === 'shape' || n.type === 'text') {
        blend(fillLumAlpha((n as ShapeRenderNode).fill, images), cover)
      } else if (n.type === 'table') {
        const tbl = n as TableRenderNode
        if (tbl.bgFill && cover > 0) blend(fillLumAlpha(tbl.bgFill, images), cover)
        // same placement as the cell-edit overlay, so a rotated or flipped table samples the edited cell
        const placed = { ...n.box, x: ox + n.box.x, y: oy + n.box.y }
        for (const c of tbl.cells) {
          const cb = tableCellOverlayBox(placed, c)
          blend(fillLumAlpha(c.fill, images), coverage(cb.x, cb.y, cb.w, cb.h))
        }
      } else if (n.type === 'chart') {
        const chart = n as ChartRenderNode
        if (chart.bgFill) blend(fillLumAlpha(chart.bgFill, images), cover)
        const plot = chart.plotRect
        if (plot?.fill) {
          const px = ox + n.box.x + plot.x
          const py = oy + n.box.y + plot.y
          blend(fillLumAlpha(plot.fill, images), coverage(px, py, plot.w, plot.h))
        }
      }
    }
  }
  blend(fillLumAlpha(slide.background, images), 1)
  visit(slide.nodes, 0, 0)
  return (lum ?? 1) < 0.5
}

/** Selection/edit chrome color: white on dark backgrounds, near-black otherwise (shared
 * with the text-edit overlay so the edit frame matches the selection frame). With a
 * region (the selection's bounding box in slide px) the verdict samples what is actually
 * painted under it; without one it falls back to the whole-page verdict. */
export function selectionChromeColor(
  slide: RenderSlide,
  images: Map<string, HTMLImageElement>,
  region?: { x: number; y: number; w: number; h: number } | null,
): string {
  const dark = region ? regionIsDark(slide, images, region) : slideBackgroundIsDark(slide, images)
  return dark ? '#ffffff' : '#232425'
}

/** PowerPoint-style rotate handle: white disc with a clockwise circular arrow.
 * Rendered once to an offscreen canvas (4× for retina) and applied to the
 * Transformer's `rotater` anchor as a fill pattern (Konva anchors are Rects). */
const ROTATER_SIZE = 19
let rotaterIcon: HTMLCanvasElement | null = null
function getRotaterIcon(): HTMLCanvasElement {
  if (rotaterIcon) return rotaterIcon
  const k = 4
  const s = ROTATER_SIZE * k
  const c = document.createElement('canvas')
  c.width = s
  c.height = s
  const ctx = c.getContext('2d')!
  const mid = s / 2
  // disc: white fill + selection-color rim (~0.85px at final size, matching the frame hairline)
  ctx.beginPath()
  ctx.arc(mid, mid, mid - k, 0, Math.PI * 2)
  ctx.fillStyle = '#ffffff'
  ctx.fill()
  ctx.lineWidth = k * 0.85
  ctx.strokeStyle = '#232425'
  ctx.stroke()
  // circular arrow: 300° arc with the gap at 12 o'clock (canvas angles grow clockwise)
  const r = mid * 0.46
  const end = (Math.PI * 4) / 3 // 240°, upper-left — the arrowhead end
  ctx.beginPath()
  ctx.arc(mid, mid, r, -Math.PI / 3, end)
  ctx.lineWidth = 1.4 * k
  ctx.lineCap = 'round'
  ctx.stroke()
  // arrowhead at the 240° end, pointing along the clockwise tangent (up-right, into the gap)
  const ex = mid + r * Math.cos(end)
  const ey = mid + r * Math.sin(end)
  const tx = -Math.sin(end)
  const ty = Math.cos(end)
  const len = 2.7 * k
  const half = 1.75 * k
  ctx.beginPath()
  ctx.moveTo(ex + tx * len, ey + ty * len)
  ctx.lineTo(ex - ty * half, ey + tx * half)
  ctx.lineTo(ex + ty * half, ey - tx * half)
  ctx.closePath()
  ctx.fillStyle = '#232425'
  ctx.fill()
  rotaterIcon = c
  return c
}

const setStageCursor = (e: Konva.KonvaEventObject<Event>, cursor: string) => {
  const st = e.target.getStage()
  if (st) st.container().style.cursor = cursor
}

/** Hover cursor for the rotate anchor: the same clockwise-arrow glyph as the
 * handle, black with a white halo so it stays readable on any background
 * (replaces Konva's default crosshair). The canvas is scaled with a CSS
 * transform, which does not scale cursors, so callers pass the on-screen
 * handle size (ROTATER_SIZE × zoom) × 1.3; hotspot at the image center. */
function makeRotateCursor(sizePx: number): string {
  const size = Math.max(10, Math.round(sizePx))
  const arc = 'M16 5.07A8 8 0 1 1 8 5.07'
  const head = 'M11.64 2.97L9.35 7.41L6.65 2.73Z'
  const svg =
    `<svg xmlns="http://www.w3.org/2000/svg" width="${size}" height="${size}" viewBox="0 0 24 24">` +
    '<g fill="none" stroke-linecap="round">' +
    `<path d="${arc}" stroke="#fff" stroke-width="5.5"/>` +
    `<path d="${head}" fill="#fff" stroke="#fff" stroke-width="3" stroke-linejoin="round"/>` +
    `<path d="${arc}" stroke="#232425" stroke-width="2.4"/>` +
    `<path d="${head}" fill="#232425"/>` +
    '</g></svg>'
  const hot = Math.floor(size / 2)
  return `url("data:image/svg+xml,${encodeURIComponent(svg)}") ${hot} ${hot}, crosshair`
}

/** Restyle the Transformer's rotate anchor from the default square into the rotate icon.
 * `size` is in canvas px (the caller applies zoom compensation and low-zoom shrinking). */
function styleRotaterAnchor(anchor: Konva.Rect, size: number): void {
  if (!anchor.hasName('rotater')) return
  anchor.setAttrs({
    width: size,
    height: size,
    offsetX: size / 2,
    offsetY: size / 2,
    cornerRadius: size / 2,
    strokeEnabled: false,
    fillPriority: 'pattern',
    fillPatternImage: getRotaterIcon(),
    fillPatternRepeat: 'no-repeat',
    fillPatternScale: { x: size / (ROTATER_SIZE * 4), y: size / (ROTATER_SIZE * 4) },
  })
}

/** Find a node by sourceId (top level + recursively inside groups; while editing inside a group the selection may contain children). */
function findNodeDeep(nodes: RenderNode[], id: string): RenderNode | undefined {
  for (const n of nodes) {
    if (n.sourceId === id) return n
    if (n.type === 'group') {
      const c = findNodeDeep((n as GroupRenderNode).children, id)
      if (c) return c
    }
  }
  return undefined
}

export interface SlideCanvasHandle {
  /** Start dragging a node from a press that landed outside the stage (the text-edit overlay's frame). */
  startNodeDrag: (sourceId: string, ev: MouseEvent) => void
}

interface Props {
  ref?: React.Ref<SlideCanvasHandle>
  slide: RenderSlide
  selectedIds: string[]
  onSelect: (sourceId: string | null, additive?: boolean) => void
  /** caret = viewport coordinates of the double-click (editor places the caret/selects the word there; defaults to caret at end) */
  onEditText: (sourceId: string, caret?: EditCaret) => void
  /** preview=true: live preview commit during drag (not added to undo history, see EditTransformOp.preview);
   * groupId: geometry commit for a child while editing inside a group (box is in group-local coordinates) */
  onTransform: (
    sourceId: string,
    box: {
      x: number
      y: number
      w: number
      h: number
      rotationDeg: number
      tableRowHeightsPx?: number[]
    },
    preview?: boolean,
    groupId?: string,
  ) => void | Promise<void>
  /** Double-click a table cell: row/col are model coordinates */
  onEditTableCell: (sourceId: string, row: number, col: number) => void
  /** Press inside a cell of a not-yet-selected table: select the table AND that cell and let the
   * cell layer take over the gesture (drag = block selection). Only borders / the frame move a table. */
  onTableCellPress?: (sourceId: string, cell: { row: number; col: number }) => void
  /** Double-click an audio/video element: trigger the playback overlay */
  onPlayMedia?: (sourceId: string) => void
  /** Right-click: hit element gives sourceId, blank area gives null; table hits include cell model coordinates */
  onContextMenu: (
    sourceId: string | null,
    clientX: number,
    clientY: number,
    cell?: { row: number; col: number },
  ) => void
  /** Drag to resize a column (col is the tc index, wPx is the column's new width in slide px) */
  onTableColResize: (
    sourceId: string,
    col: number,
    wPx: number,
    neighbor?: { col: number; wPx: number },
  ) => void
  onTableRowResize?: (sourceId: string, row: number, hPx: number) => void
  /** A border drag is in flight: the table follows the pointer live (nothing is committed until release) */
  onTableGridMove?: (sourceId: string, change: GridChange) => void
  /** The border drag was abandoned (Esc): drop the live frames */
  onTableGridCancel?: () => void
  images: Map<string, HTMLImageElement>
  /** Element/cell currently edited in the DOM overlay: canvas skips drawing its text (avoids ghosting under the overlay) */
  editingText?: { sourceId: string; cell?: { row: number; col: number } } | null
  /** Canvas CSS zoom (App's zoom): screen-distance semantics such as drag thresholds must be converted back to canvas coordinates */
  zoom?: number
  /** Rubber-band selection (drag a rectangle on blank area): set of elements fully inside the rectangle */
  onMarqueeSelect?: (sourceIds: string[]) => void
  /** Option+drag duplicate: original snaps back, a copy is created at the drop offset (dx/dy in slide px) */
  onDuplicateTo?: (sourceId: string, dxPx: number, dyPx: number) => void
  /** Group being edited from inside after double-click-into-group (its children are selectable/editable) */
  enteredGroupId?: string | null
  /** Double-click a group to enter in-group editing (childId is the child hit by the double-click, may be empty) */
  onEnterGroup?: (groupId: string, childId: string | null) => void
  /** Connector endpoint drag committed: new endpoints in slide px; start/end = attachment change for that end (undefined keep, null detach, object attach) */
  onEditConnectorEndpoints?: (
    sourceId: string,
    ep: {
      x1: number
      y1: number
      x2: number
      y2: number
      start?: { targetId: string; idx: number } | null
      end?: { targetId: string; idx: number } | null
    },
  ) => void
  /** Shape draw mode (ribbon gallery pick): crosshair cursor, mousedown draws over anything; click = default size */
  drawMode?: { kind: string } | null
  /** Draw gesture committed: box to insert (slide px; flipH/flipV restore a line's drag direction) */
  onDrawCommit?: (rect: DrawRect) => void
  /** Draw mode cancelled from inside the canvas (Escape) */
  onDrawCancel?: () => void
  /** Yellow adjust-handle drag: full avLst map; preview=true during the drag, false on release */
  onAdjust?: (sourceId: string, adjust: Record<string, number>, preview: boolean) => void
}

/**
 * Canvas bleed margin (slide px): PowerPoint's edit view draws elements that extend past the
 * slide boundary on the surrounding workspace, but pixels outside the canvas cannot be drawn.
 * Add a bleed ring around the Stage; content is wrapped in an offset Group (slide coordinate
 * system unchanged), and the Stage container is absolutely positioned at -BLEED to align.
 */
export const CANVAS_BLEED = 160

/* Screenshot-automation hook (fidelity-compare): window.__genofficeHidePhPrompts = true +
 * dispatching 'genoffice:hide-ph-prompts' hides empty-placeholder hints on the edit canvas. */
const hidePhPromptsListeners = new Set<() => void>()
window.addEventListener('genoffice:hide-ph-prompts', () => {
  for (const l of hidePhPromptsListeners) l()
})
const subscribeHidePhPrompts = (cb: () => void) => {
  hidePhPromptsListeners.add(cb)
  return () => {
    hidePhPromptsListeners.delete(cb)
  }
}
const getHidePhPrompts = () =>
  !!(window as { __genofficeHidePhPrompts?: boolean }).__genofficeHidePhPrompts

/** Default rotate-handle snapping: lock onto 45° multiples (Shift switches to 15° steps) */
const ROTATION_SNAPS = [0, 45, 90, 135, 180, 225, 270, 315]
const ROTATION_SNAP_TOLERANCE = 5

/** Walk up the Konva parent chain to the owning node_<sourceId> Group (null for background/decoration/stage). */
function nodeIdFromTarget(t: Konva.Node | null): string | null {
  while (t && t !== t.getStage()) {
    const id = typeof t.id === 'function' ? t.id() : ''
    if (id && id.startsWith('node_')) return id.slice('node_'.length)
    t = t.getParent()
  }
  return null
}

/**
 * Node count from which a slide counts as "dense" and gets its raster pressure capped:
 * every edit redraws the whole layer, and on a ~600-element page at pixelRatio 3 that keeps the
 * GPU command buffer near-full — the prime suspect for the sporadic renderer main-thread freeze.
 */
export const DENSE_SLIDE_NODE_COUNT = 300

/** Total node count including group children (a dense page may hide its elements inside groups). */
function countNodes(nodes: RenderNode[]): number {
  let n = 0
  for (const node of nodes) {
    n += 1
    if (node.type === 'group') n += countNodes((node as GroupRenderNode).children)
  }
  return n
}

/** Below this node count a slide is "light": full-layer redraws are cheap enough to afford
 * full-resolution rasterization at deep zoom (retina × 300% needs ratio 6; at that point the
 * whole bitmap — content and selection chrome — would otherwise be a 2× upscale and read soft). */
export const LIGHT_SLIDE_NODE_COUNT = 100

/**
 * Main-canvas rasterization ratio. Zoom is a pure CSS transform on the Stage's container, so
 * without this the bitmap stays at devicePixelRatio and zoom>1 just stretches it (blurry).
 * Never below devicePixelRatio (zoom<1 keeps the old quality); capped to bound canvas memory —
 * the cap is tiered by node count because full-layer redraw cost scales with ratio² × node
 * count: light pages get the full dpr×zoom (up to 6, ~160 MB transient on a slide-sized stage),
 * mid pages keep the historical 3, and dense slides skip the zoom upscale and cap harder.
 */
export function canvasPixelRatio(devicePixelRatio: number, zoom: number, nodeCount = 0): number {
  if (nodeCount >= DENSE_SLIDE_NODE_COUNT) return Math.min(devicePixelRatio || 1, 2)
  const cap = nodeCount < LIGHT_SLIDE_NODE_COUNT ? 6 : 3
  return Math.min((devicePixelRatio || 1) * Math.max(zoom, 1), cap)
}

/** Selection-chrome stroke width (canvas px): as thin as possible while staying SOLID.
 * Chrome geometry sits at arbitrary fractional coordinates, so a stroke near 1 raster px
 * anti-aliases into two ~50%-alpha pixels — faint gray, not a line. ~1.7 raster px keeps a
 * fully-covered core at any alignment; the screen floor (~1.5 device px) keeps it solid
 * when zooming out shrinks canvas px below device px. */
function chromeHairline(zoom: number, nodeCount = 0): number {
  const dpr = window.devicePixelRatio || 1
  const z = Math.max(zoom, 0.1)
  return Math.max(1.5 / (dpr * z), 1.7 / canvasPixelRatio(dpr, zoom, nodeCount))
}

/** Counter-scale the Transformer chrome for a given zoom so the frame, anchors and
 * rotate handle hold a constant on-screen size. Called from JSX values on commits and
 * imperatively per frame during a live zoom gesture — the Transformer sits on its own
 * layer, so the per-frame redraw touches only the chrome, never the slide raster. */
function applyChromeZoom(tr: Konva.Transformer, zoom: number, nodeCount: number): void {
  const z = Math.max(zoom, 0.1)
  const cs = Math.min(1, Math.sqrt(z))
  const hl = chromeHairline(zoom, nodeCount)
  tr.setAttrs({
    borderStrokeWidth: hl,
    anchorStrokeWidth: hl,
    anchorSize: (8 * cs) / z,
    rotateAnchorOffset: (50 * cs) / z,
    anchorStyleFunc: (a: Konva.Rect) => styleRotaterAnchor(a, (ROTATER_SIZE * cs) / z),
  })
  tr.forceUpdate()
  tr.getLayer()?.batchDraw()
}

/** PowerPoint's adjust-handle yellow — canvas editing chrome, identical in both
 * themes (like the Transformer's white anchors). */
const ADJUST_HANDLE_FILL = '#ffc94d'
const ADJUST_HANDLE_STROKE = '#a97d00'

/**
 * Yellow adjust handles for the single selected adjustable shape (preset avLst
 * edit). The handle position derives from the committed node; during a drag it
 * is pinned imperatively to the constraint path, and onAdjust fires throttled
 * preview commits (the rebuilt slide re-renders the geometry live) with a final
 * non-preview commit on release — one whole drag = one undo step.
 */
function AdjustHandles({
  node,
  zoom,
  chromeScale,
  onAdjust,
}: {
  node: ShapeRenderNode
  zoom: number
  chromeScale: number
  onAdjust: (adjust: Record<string, number>, preview: boolean) => void
}) {
  const specs = adjustHandleSpecs(node)
  if (!specs.length) return null
  const b = node.box
  const val = (k: string, d: number) => node.adjust?.[k] ?? d
  /** Committed values for every handle key (the engine rewrites the whole avLst) */
  const fullAdjust = (): Record<string, number> => {
    const out: Record<string, number> = { ...(node.adjust ?? {}) }
    for (const sp of specs) for (const k of sp.keys) if (out[k.name] == null) out[k.name] = k.def
    return out
  }
  const rot = ((b.rotationDeg || 0) * Math.PI) / 180
  const cos = Math.cos(rot)
  const sin = Math.sin(rot)
  const toStage = (px: number, py: number) => {
    const lx = b.flipH ? b.w - px : px
    const ly = b.flipV ? b.h - py : py
    const dx = lx - b.w / 2
    const dy = ly - b.h / 2
    return { x: b.x + b.w / 2 + dx * cos - dy * sin, y: b.y + b.h / 2 + dx * sin + dy * cos }
  }
  const toLocal = (sx: number, sy: number) => {
    const dx = sx - (b.x + b.w / 2)
    const dy = sy - (b.y + b.h / 2)
    let lx = b.w / 2 + dx * cos + dy * sin
    let ly = b.h / 2 - dx * sin + dy * cos
    if (b.flipH) lx = b.w - lx
    if (b.flipV) ly = b.h - ly
    return { x: lx, y: ly }
  }
  const z = Math.max(zoom, 0.1)
  // A yellow handle sitting on a resize handle (tiny shape, or a preset whose handle lies on a corner /
  // edge middle) would swallow the press meant for the resize handle: the resize handle wins there.
  const anchorPts = anchorsForScreenSize(b.w * zoom, b.h * zoom, 8 * chromeScale).map((name) => ({
    x: name.includes('left') ? 0 : name.includes('right') ? b.w : b.w / 2,
    y: name.startsWith('top') ? 0 : name.startsWith('bottom') ? b.h : b.h / 2,
  }))
  const reach = (4 + 5) * chromeScale // anchor half-size + adjust-handle radius, screen px
  return (
    <>
      {specs.map((spec) => {
        const p = spec.pos(b.w, b.h, val)
        if (anchorPts.some((a) => Math.hypot(a.x - p.x, a.y - p.y) * zoom < reach)) return null
        const s = toStage(p.x, p.y)
        const valuesAt = (t: Konva.Node) => {
          const l = toLocal(t.x(), t.y())
          return spec.values(b.w, b.h, l.x, l.y, val)
        }
        return (
          <Circle
            key={spec.keys.map((k) => k.name).join('-')}
            x={s.x}
            y={s.y}
            radius={(5 * chromeScale) / z}
            fill={ADJUST_HANDLE_FILL}
            stroke={ADJUST_HANDLE_STROKE}
            strokeWidth={1 / z}
            draggable
            onDragMove={(e) => {
              const merged = { ...fullAdjust(), ...valuesAt(e.target) }
              // pin the handle onto its constraint path
              const np = spec.pos(b.w, b.h, (k, d) => merged[k] ?? d)
              const ns = toStage(np.x, np.y)
              e.target.position(ns)
              onAdjust(merged, true)
            }}
            onDragEnd={(e) => {
              onAdjust({ ...fullAdjust(), ...valuesAt(e.target) }, false)
            }}
          />
        )
      })}
    </>
  )
}

/** A column / row border drag as the press found the table (the live preview re-grids it afterwards) */
interface GripDrag {
  id: string
  kind: 'col' | 'row'
  /** Grid column / row whose far border is dragged */
  idx: number
  /** Column widths / row heights at the press (px) */
  sizes: number[]
  /** Content need per column / row at the press (px) */
  needs?: number[]
  /** Slide-space start of the dragged column / row (its logical left / top edge) */
  origin: number
  /** rtl columns: slide-space visual-left edge of the dragged column */
  edge: number
  /** The grip's own resting coordinate (to put it back after a cancel) */
  home: number
  cancelled: boolean
}

function guidesEqual(a: Guide[], b: Guide[]): boolean {
  if (a === b) return true
  if (a.length !== b.length) return false
  return a.every((g, i) => g.axis === b[i]?.axis && Math.abs(g.pos - b[i]!.pos) < 0.5)
}

function spacingEqual(a: SpacingIndicator[], b: SpacingIndicator[]): boolean {
  if (a === b) return true
  if (a.length !== b.length) return false
  return a.every(
    (s, i) =>
      s.axis === b[i]?.axis &&
      Math.abs(s.from - b[i]!.from) < 0.5 &&
      Math.abs(s.to - b[i]!.to) < 0.5 &&
      Math.abs(s.at - b[i]!.at) < 0.5,
  )
}

export function SlideCanvas({
  slide,
  selectedIds,
  onSelect,
  onEditText,
  onTransform,
  onEditTableCell,
  onTableCellPress,
  onTableColResize,
  onTableRowResize,
  onTableGridMove,
  onTableGridCancel,
  onPlayMedia,
  onContextMenu,
  images,
  editingText,
  zoom = 1,
  onMarqueeSelect,
  onDuplicateTo,
  enteredGroupId,
  onEnterGroup,
  onEditConnectorEndpoints,
  drawMode,
  onDrawCommit,
  onDrawCancel,
  onAdjust,
  ref,
}: Props) {
  const trRef = useRef<Konva.Transformer>(null)
  const lastSnapRef = useRef<{ guides: Guide[]; spacing: SpacingIndicator[] } | null>(null)
  const gripRef = useRef<GripDrag | null>(null)
  const gripEscRef = useRef<(() => void) | null>(null)
  /** Esc while a column / row border is held abandons the drag (the table goes back, no undo step) */
  const armGripEscape = (node: Konva.Node) => {
    gripEscRef.current?.()
    const onKey = (ev: KeyboardEvent) => {
      if (ev.key !== 'Escape' || !gripRef.current) return
      ev.preventDefault()
      ev.stopPropagation()
      gripRef.current.cancelled = true
      node.stopDrag()
    }
    window.addEventListener('keydown', onKey, true)
    gripEscRef.current = () => window.removeEventListener('keydown', onKey, true)
  }
  const disarmGripEscape = () => {
    gripEscRef.current?.()
    gripEscRef.current = null
  }
  useEffect(() => () => gripEscRef.current?.(), [])

  /** Minimum size of the current selection for resize gestures (see min-size.ts) */
  /** The floor is fixed when a drag starts: preview frames re-lay the element out smaller as the pointer
   * moves, and a floor derived from the shrinking geometry would ratchet down with it. */
  const gestureMinRef = useRef<{ anchor: string; min: MinSize } | null>(null)
  const selectionMin = (anchor?: string): MinSize => {
    if (
      trActiveRef.current &&
      gestureMinRef.current &&
      gestureMinRef.current.anchor === (anchor ?? '')
    )
      return gestureMinRef.current.min
    return computeSelectionMin(anchor)
  }
  const computeSelectionMin = (anchor?: string): MinSize => {
    const targets = selectedIds
      .map((id) => findNodeDeep(slide.nodes, id))
      .filter((n): n is RenderNode => !!n)
    const min = minSizeOfNodes(targets)
    // Proportional (picture) corner drags keep w:h, so the floor must hold in both sides at once
    const corner = !!anchor && !anchor.includes('center') && !anchor.includes('middle')
    if (corner && targets.length > 0 && targets.every((n) => n.type === 'picture')) {
      const bb = selectionBounds(targets)
      if (bb.w > 0 && bb.h > 0) {
        const f = Math.max(min.w / bb.w, min.h / bb.h)
        return { w: bb.w * f, h: bb.h * f }
      }
    }
    return min
  }
  /** True while a handle is being dragged: the handle set must not change under the pointer */
  const trActiveRef = useRef(false)
  const anchorsRef = useRef<string[]>([...ALL_ANCHORS])
  const liveAnchors = (): string[] => {
    if (trActiveRef.current) return anchorsRef.current
    const targets = selectedIds
      .map((id) => findNodeDeep(slide.nodes, id))
      .filter((n): n is RenderNode => !!n && !isConnectorNode(n))
    if (targets.length === 0) return anchorsRef.current
    const b = targets.length === 1 ? targets[0]!.box : selectionBounds(targets)
    const w = 'w' in b ? b.w : 0
    const h = 'h' in b ? b.h : 0
    const next = anchorsForScreenSize(w * zoom, h * zoom, 8 * chromeScale)
    if (next.join() !== anchorsRef.current.join()) anchorsRef.current = next
    return anchorsRef.current
  }
  const layerRef = useRef<Konva.Layer>(null)
  const stageRef = useRef<Konva.Stage>(null)
  const selectionDrag = useRef(createSelectionDrag()).current
  const selectionMoveCommit = useRef<((delta: { x: number; y: number }) => Promise<void>) | null>(
    null,
  )

  useImperativeHandle(ref, () => ({
    startNodeDrag(sourceId, ev) {
      const stage = stageRef.current
      const n = stage?.findOne<Konva.Node>(`#node_${sourceId}`)
      if (!stage || !n?.draggable()) return
      // Same path as a real press on the node: Konva's own mousedown listener arms the drag and
      // the window mousemove handler starts it once the pointer travels dragDistance
      stage.setPointersPositions(ev)
      n.fire('mousedown', { evt: ev, type: 'mousedown', target: n, currentTarget: n }, false)
    },
  }))

  const nodeCount = useMemo(() => countNodes(slide.nodes), [slide])
  const dense = nodeCount >= DENSE_SLIDE_NODE_COUNT

  // Zoom the raster and the selection chrome react to: during an active zoom (slider drag,
  // repeated wheel commits) both hold their last settled value — the CSS transform scales
  // them visually, which is smooth — and re-fit once the zoom stops changing. Recomputing
  // per commit forced a full-layer re-raster each time (very expensive at high pixel
  // ratios), which made the gesture stutter.
  const [settledZoom, setSettledZoom] = useState(zoom)
  useEffect(() => {
    if (settledZoom === zoom) return
    const t = window.setTimeout(() => setSettledZoom(zoom), 150)
    return () => window.clearTimeout(t)
  }, [zoom, settledZoom])

  // Bundled @font-face fonts (Carlito) and Office-private FontFaces (doc-fonts.ts) may finish
  // loading after the first draw; canvas text drawn with a fallback face must be redrawn once
  // the real font is available. 'loadingdone' covers faces added at any later point.
  useEffect(() => {
    let live = true
    const redraw = () => {
      if (!live) return
      for (const l of stageRef.current?.getLayers() ?? []) l.batchDraw()
    }
    document.fonts?.ready?.then(redraw).catch(() => {})
    document.fonts?.addEventListener?.('loadingdone', redraw)
    return () => {
      live = false
      document.fonts?.removeEventListener?.('loadingdone', redraw)
    }
  }, [])

  // Force a layer batchDraw whenever slide data changes so same-page updates (reloads, undo,
  // deck synchronization) immediately repaint on the Konva canvas.
  useLayoutEffect(() => {
    for (const l of stageRef.current?.getLayers() ?? []) {
      l.batchDraw()
    }
  }, [slide])

  // Re-rasterize on settled zoom changes. The slide dep covers layers (re)created between
  // zoom changes; the resolution media query covers devicePixelRatio changes (window dragged
  // to another display) — it matches the current dpr, so it must be re-armed after each change.
  useEffect(() => {
    const apply = () => {
      const ratio = canvasPixelRatio(window.devicePixelRatio, settledZoom, nodeCount)
      for (const l of stageRef.current?.getLayers() ?? []) {
        if (l.getCanvas().getPixelRatio() !== ratio) {
          l.getCanvas().setPixelRatio(ratio)
          l.batchDraw()
        }
      }
    }
    apply()
    let mq: MediaQueryList | null = null
    const onDprChange = () => {
      apply()
      arm()
    }
    const arm = () => {
      mq?.removeEventListener('change', onDprChange)
      mq = window.matchMedia(`(resolution: ${window.devicePixelRatio}dppx)`)
      mq.addEventListener('change', onDprChange)
    }
    arm()
    return () => mq?.removeEventListener('change', onDprChange)
  }, [settledZoom, slide, nodeCount])
  const [guides, setGuides] = useState<Guide[]>([])
  // Equal-spacing double-headed arrows (while dragging); same-size matched elements (while resizing, listed per dimension)
  const [spacing, setSpacing] = useState<SpacingIndicator[]>([])
  const [sizeMatch, setSizeMatch] = useState<{ w: string[]; h: string[] } | null>(null)
  const sizeMatchKeyRef = useRef('')
  // A marquee drag just ended on this gesture: swallow the trailing click so it doesn't select the node under the cursor
  const suppressClickRef = useRef(false)
  // Full-page background-like nodes: a drag on one that is not yet selected rubber-bands
  // instead of moving it; once selected it drags like any other node (PowerPoint)
  const backgroundIds = useMemo(
    () => new Set(slide.nodes.filter((n) => n.background).map((n) => n.sourceId)),
    [slide],
  )
  // Rubber-band selection rectangle (slide coordinates); null = not rubber-band selecting
  const [marquee, setMarquee] = useState<{ x1: number; y1: number; x2: number; y2: number } | null>(
    null,
  )
  const marqueeRef = useRef<{ x1: number; y1: number; x2: number; y2: number } | null>(null)

  // Handles keep a constant on-screen size at zoom ≥ 1, but shrink gently (√zoom) when zoomed
  // out so they stay proportionate to the shrinking content. Chrome sizing follows the
  // committed zoom directly (the Transformer lives on its own cheap layer), not settledZoom —
  // only the expensive raster ratio waits for the zoom to settle.
  const chromeScale = Math.min(1, Math.sqrt(Math.max(zoom, 0.1)))

  // Rotate hover cursor: 30% larger than the handle's current on-screen size
  const rotateCursor = useMemo(
    () => makeRotateCursor(ROTATER_SIZE * chromeScale * 1.3),
    [chromeScale],
  )

  // Selection chrome flips to white on dark backgrounds so the frame stays visible —
  // judged under the selection itself (a text box on a white plate over a colored page
  // needs dark chrome), falling back to the page verdict when nothing is selected or a
  // selected id lives inside a group (group-local coords).
  const selStroke = useMemo(() => {
    const boxes = selectedIds
      .map((id) => slide.nodes.find((n) => n.sourceId === id)?.box)
      .filter((b): b is NonNullable<typeof b> => !!b)
    let region: { x: number; y: number; w: number; h: number } | null = null
    if (boxes.length === selectedIds.length && boxes.length > 0) {
      const x = Math.min(...boxes.map((b) => b.x))
      const y = Math.min(...boxes.map((b) => b.y))
      region = {
        x,
        y,
        w: Math.max(...boxes.map((b) => b.x + b.w)) - x,
        h: Math.max(...boxes.map((b) => b.y + b.h)) - y,
      }
    }
    return selectionChromeColor(slide, images, region)
  }, [slide, images, selectedIds])

  // Hairline width for the current zoom + raster resolution (canvas px)
  const hairline = useMemo(() => chromeHairline(zoom, nodeCount), [zoom, nodeCount])

  // During a live zoom gesture (pinch/slider) only the CSS transform moves — no React
  // commit — so the chrome would stretch with the content and then snap back at commit.
  // Counter-scale it imperatively every previewed frame instead; at commit the JSX lands
  // on exactly the same values, so the gesture ends with no visible jump.
  useEffect(() => {
    const onPreview = (e: Event) => {
      const z = (e as CustomEvent<number>).detail
      if (typeof z !== 'number' || !trRef.current) return
      applyChromeZoom(trRef.current, z, nodeCount)
    }
    window.addEventListener(ZOOM_PREVIEW_EVENT, onPreview)
    return () => window.removeEventListener(ZOOM_PREVIEW_EVENT, onPreview)
  }, [nodeCount])

  // In-flight shape draw gesture (slide coordinates); preview only shows past the click threshold
  const drawRef = useRef<{ x1: number; y1: number; x2: number; y2: number; shift: boolean } | null>(
    null,
  )
  const [drawPreview, setDrawPreview] = useState<DrawRect | null>(null)

  // Draw mode: Escape cancels; toggling Shift mid-drag re-constrains the preview without mouse movement
  useEffect(() => {
    if (!drawMode) {
      drawRef.current = null
      setDrawPreview(null)
      return
    }
    const onKey = (ev: KeyboardEvent) => {
      if (ev.key === 'Escape') {
        drawRef.current = null
        setDrawPreview(null)
        onDrawCancel?.()
        return
      }
      if (ev.key === 'Shift') {
        const d = drawRef.current
        if (!d) return
        d.shift = ev.type === 'keydown'
        setDrawPreview(resolveDrawRect(drawMode.kind, d.x1, d.y1, d.x2, d.y2, d.shift))
      }
    }
    window.addEventListener('keydown', onKey)
    window.addEventListener('keyup', onKey)
    return () => {
      window.removeEventListener('keydown', onKey)
      window.removeEventListener('keyup', onKey)
    }
  }, [drawMode, onDrawCancel])

  // Hold Shift while rotating -> snap in 15° steps; release restores the default 45° snaps
  useEffect(() => {
    const setSnaps = (fine: boolean) => {
      const tr = trRef.current
      if (!tr) return
      tr.rotationSnaps(fine ? Array.from({ length: 24 }, (_, i) => i * 15) : ROTATION_SNAPS)
      tr.rotationSnapTolerance(fine ? 7.5 : ROTATION_SNAP_TOLERANCE)
    }
    const down = (e: KeyboardEvent) => e.key === 'Shift' && setSnaps(true)
    const up = (e: KeyboardEvent) => e.key === 'Shift' && setSnaps(false)
    window.addEventListener('keydown', down)
    window.addEventListener('keyup', up)
    return () => {
      window.removeEventListener('keydown', down)
      window.removeEventListener('keyup', up)
    }
  }, [])

  // Selection changed -> attach the Transformer to all selected nodes (except connectors: selectable but no transform handles)
  useEffect(() => {
    const tr = trRef.current
    const layer = layerRef.current
    if (!tr || !layer) return
    const nodes = selectedIds
      .filter((id) => {
        const n = findNodeDeep(slide.nodes, id)
        return n && !isConnectorNode(n)
      })
      .map((id) => layer.findOne(`#node_${id}`))
      .filter((n): n is Konva.Node => !!n)
    tr.nodes(nodes)
    tr.getLayer()?.batchDraw()
  }, [selectedIds, slide, enteredGroupId])

  // Snap target edges of the other elements (excluding the dragged selection and decoration layer) + page center lines
  const snapTargets = (excludeIds: string[]): SnapTarget[] => {
    const list: SnapTarget[] = [
      { v: [slide.widthPx / 2], h: [slide.heightPx / 2] }, // Page center
      { v: [0, slide.widthPx], h: [0, slide.heightPx] }, // Page edges
    ]
    for (const n of slide.nodes) {
      if (excludeIds.includes(n.sourceId) || n.decoration || n.background) continue
      const b = n.box
      list.push({ v: [b.x, b.x + b.w / 2, b.x + b.w], h: [b.y, b.y + b.h / 2, b.y + b.h] })
    }
    return list
  }

  // Neighbor boxes for equal-spacing snapping (same exclusion rules as snapTargets)
  const spacingBoxes = (excludeIds: string[]) =>
    slide.nodes
      .filter((n) => !excludeIds.includes(n.sourceId) && !n.decoration && !n.background)
      .map((n) => ({ x: n.box.x, y: n.box.y, w: n.box.w, h: n.box.h }))

  // Bounding box of a multi-select drag (snapped as a whole; if it contains group children the coordinate systems differ, so degrade to no snapping)
  const selBBox = useMemo(() => {
    if (selectedIds.length < 2) return null
    const boxes = selectedIds
      .map((id) => slide.nodes.find((n) => n.sourceId === id))
      .filter((n): n is RenderNode => !!n)
      .map((n) => n.box)
    if (boxes.length !== selectedIds.length) return null
    const x = Math.min(...boxes.map((b) => b.x))
    const y = Math.min(...boxes.map((b) => b.y))
    return {
      x,
      y,
      w: Math.max(...boxes.map((b) => b.x + b.w)) - x,
      h: Math.max(...boxes.map((b) => b.y + b.h)) - y,
    }
  }, [selectedIds, slide])

  const beginSelectionDrag = (sourceId: string) => {
    selectionMoveCommit.current = null
    selectionDrag.reset()
    lastSnapRef.current = null
    const pointer = stageRef.current?.getPointerPosition()
    if (!pointer || selectedIds.length < 2 || !selectedIds.includes(sourceId)) return
    const members = selectedIds
      .map((id) => findNodeDeep(slide.nodes, id))
      .filter((n): n is RenderNode => !!n)
    const snapshots = members.map((n) => ({ ...n, box: { ...n.box } }))
    const positions = new Map<string, { x: number; y: number }>()
    const parents = new Map<string, Konva.Transform>()
    for (const id of selectedIds) {
      const member = layerRef.current?.findOne<Konva.Node>(`#node_${id}`)
      if (member) {
        positions.set(id, { ...member.absolutePosition() })
        const parent = member.getParent()
        if (parent) parents.set(id, parent.getAbsoluteTransform().copy().invert())
      }
    }
    selectionMoveCommit.current = async (delta) => {
      const requests = selectionMoveRequests(
        snapshots,
        delta,
        enteredGroupId ?? undefined,
        (id, d) => {
          const inverse = parents.get(id)
          if (!inverse) return d
          const origin = inverse.point({ x: 0, y: 0 })
          const point = inverse.point(d)
          return { x: point.x - origin.x, y: point.y - origin.y }
        },
      )
      await Promise.all(requests.map((r) => onTransform(r.sourceId, r.box, undefined, r.groupId)))
    }
    selectionDrag.begin(pointer, positions, (delta) => {
      if (!selBBox) return delta
      const bb = { ...selBBox, x: selBBox.x + delta.x, y: selBBox.y + delta.y }
      const threshold = 6 / Math.max(zoom, 0.1)
      const snap = computeSnap(bb, snapTargets(selectedIds), threshold)
      const spaced = computeSpacingSnap(
        { ...bb, x: snap.x ?? bb.x, y: snap.y ?? bb.y },
        spacingBoxes(selectedIds),
        threshold,
      )
      lastSnapRef.current = {
        guides: snap.guides,
        spacing: spaced.indicators.filter((i) =>
          i.axis === 'x' ? snap.x == null : snap.y == null,
        ),
      }
      return {
        x: delta.x + (snap.x ?? spaced.x ?? bb.x) - bb.x,
        y: delta.y + (snap.y ?? spaced.y ?? bb.y) - bb.y,
      }
    })
  }

  const endSelectionDrag = (): Promise<void> => {
    lastSnapRef.current = null
    const delta = selectionDrag.finish()
    if (!delta) return Promise.resolve()
    // Proxy drags created on the last mousemove may miss Konva's mouseup filter
    // (their pointer id has not been initialized yet). End them explicitly so
    // later pointer movement cannot keep moving peers after this drop.
    stopSelectionDrags([
      ...selectedIds.map((id) => layerRef.current?.findOne<Konva.Node>(`#node_${id}`)),
      trRef.current,
    ])
    return selectionMoveCommit.current?.(delta) ?? Promise.resolve()
  }
  const syncSelectionDrag = () => {
    if (lastSnapRef.current) {
      const { guides: g, spacing: sp } = lastSnapRef.current
      setGuides((prev) => (guidesEqual(prev, g) ? prev : g))
      setSpacing((prev) => (spacingEqual(prev, sp) ? prev : sp))
    }
    const pointer = stageRef.current?.getPointerPosition()
    if (!pointer) return
    // Connectors are not attached to Transformer proxy drag: synchronize them synchronously here
    for (const id of selectedIds) {
      const node = findNodeDeep(slide.nodes, id)
      if (node && isConnectorNode(node)) {
        const position = selectionDrag.bound(id, pointer)
        if (position)
          layerRef.current?.findOne<Konva.Node>(`#node_${id}`)?.absolutePosition(position)
      }
    }
  }

  return (
    <Stage
      ref={stageRef}
      width={slide.widthPx + CANVAS_BLEED * 2}
      height={slide.heightPx + CANVAS_BLEED * 2}
      onMouseDown={(e) => {
        // Draw mode swallows the gesture before any selection logic (drawing works over elements too)
        if (drawMode) {
          if (e.evt.button !== 0) return
          const raw = e.target.getStage()?.getPointerPosition()
          if (!raw) return
          const x = raw.x - CANVAS_BLEED
          const y = raw.y - CANVAS_BLEED
          drawRef.current = { x1: x, y1: y, x2: x, y2: y, shift: e.evt.shiftKey }
          return
        }
        suppressClickRef.current = false
        if (e.evt.button === 0 && e.target.getParent()?.getClassName() === 'Transformer') {
          beginSelectionDrag(selectedIds[0] ?? '')
        }
        // Blank = the Stage itself (bleed area), the slide base/background (name=slide-bg),
        // or a full-page background-like node (still click-selectable, but a drag on it
        // rubber-bands instead of moving it)
        const isBlank =
          e.target === e.target.getStage() ||
          (typeof e.target.name === 'function' && e.target.name() === 'slide-bg')
        const hitId = isBlank ? null : nodeIdFromTarget(e.target)
        const onBackground =
          hitId != null && backgroundIds.has(hitId) && !selectedIds.includes(hitId)
        if (!isBlank && !onBackground) return
        if (isBlank) onSelect(null)
        // Mouse-down on blank area -> start rubber-band selection (on release, elements fully inside the rectangle are selected)
        if (e.evt.button !== 0) return
        const raw = e.target.getStage()?.getPointerPosition()
        if (!raw || !onMarqueeSelect) return
        const start = { x: raw.x - CANVAS_BLEED, y: raw.y - CANVAS_BLEED }
        marqueeRef.current = { x1: start.x, y1: start.y, x2: start.x, y2: start.y }
        setMarquee(null) // Only show after moving past the threshold
      }}
      onMouseMove={(e) => {
        if (drawMode) {
          const d = drawRef.current
          if (!d) return
          const raw = e.target.getStage()?.getPointerPosition()
          if (!raw) return
          d.x2 = raw.x - CANVAS_BLEED
          d.y2 = raw.y - CANVAS_BLEED
          d.shift = e.evt.shiftKey
          // Within 3 screen px still counts as a click; don't show the preview yet
          if (Math.hypot(d.x2 - d.x1, d.y2 - d.y1) * zoom > 3)
            setDrawPreview(resolveDrawRect(drawMode.kind, d.x1, d.y1, d.x2, d.y2, d.shift))
          return
        }
        const m = marqueeRef.current
        if (!m) return
        const raw = e.target.getStage()?.getPointerPosition()
        if (!raw) return
        m.x2 = raw.x - CANVAS_BLEED
        m.y2 = raw.y - CANVAS_BLEED
        // Within 3 screen px counts as a click; don't show the selection rectangle
        if (Math.hypot(m.x2 - m.x1, m.y2 - m.y1) * zoom > 3) setMarquee({ ...m })
      }}
      onMouseUp={() => {
        if (drawMode) {
          const d = drawRef.current
          drawRef.current = null
          setDrawPreview(null)
          if (!d || !onDrawCommit) return
          // Click (no real drag) = insert at the predefined default size, PowerPoint-style
          const rect =
            Math.hypot(d.x2 - d.x1, d.y2 - d.y1) * zoom <= 3
              ? { x: d.x1, y: d.y1, ...defaultDrawSize(drawMode.kind) }
              : resolveDrawRect(drawMode.kind, d.x1, d.y1, d.x2, d.y2, d.shift)
          onDrawCommit(rect)
          return
        }
        const m = marqueeRef.current
        marqueeRef.current = null
        if (!m) return
        setMarquee(null)
        if (Math.hypot(m.x2 - m.x1, m.y2 - m.y1) * zoom <= 3) return // A click, not a rubber-band selection
        suppressClickRef.current = true
        const [lx, rx] = m.x1 < m.x2 ? [m.x1, m.x2] : [m.x2, m.x1]
        const [ty, by] = m.y1 < m.y2 ? [m.y1, m.y2] : [m.y2, m.y1]
        const ids = slide.nodes
          .filter((n) => {
            if (n.decoration || n.background) return false
            const b = n.box
            return b.x >= lx && b.x + b.w <= rx && b.y >= ty && b.y + b.h <= by
          })
          .map((n) => n.sourceId)
        if (ids.length) onMarqueeSelect!(ids)
        else onSelect(null) // A marquee that started on a background node skipped the mousedown clear
      }}
      onContextMenu={(e) => {
        e.evt.preventDefault()
        // Walk up the parent chain to find the node_<sourceId> Group (decoration layer/background has none -> blank-area menu)
        let t: Konva.Node | null = e.target
        let sourceId: string | null = null
        while (t && t !== t.getStage()) {
          const id = typeof t.id === 'function' ? t.id() : ''
          if (id && id.startsWith('node_')) {
            sourceId = id.slice('node_'.length)
            break
          }
          t = t.getParent()
        }
        let cell: { row: number; col: number } | undefined
        if (sourceId) {
          const n = slide.nodes.find((x) => x.sourceId === sourceId)
          const raw = e.target.getStage()?.getPointerPosition()
          // Stage coordinates -> slide coordinates (remove the bleed offset)
          const pos = raw ? { x: raw.x - CANVAS_BLEED, y: raw.y - CANVAS_BLEED } : null
          if (n && n.type === 'table' && pos) {
            const local = tableLocalPointFromStage(pos, n.box)
            const hit = tableCellAtPoint(n, local)
            if (hit) cell = { row: hit.row, col: hit.col }
          }
        }
        onContextMenu(sourceId, e.evt.clientX, e.evt.clientY, cell)
      }}
      style={{
        position: 'absolute',
        left: -CANVAS_BLEED,
        top: -CANVAS_BLEED,
        cursor: drawMode ? 'crosshair' : undefined,
      }}
    >
      {/* Draw mode disables node hit-testing: the crosshair gesture must not select/drag elements underneath */}
      <Layer ref={layerRef} x={CANVAS_BLEED} y={CANVAS_BLEED} listening={!drawMode}>
        {/* Slide base: white background + shadow (used to be the Stage's CSS background; the bleed area must show the gray workspace behind).
            Dense slides drop the decorative blur: a canvas shadow re-rasterizes on every full-layer
            redraw and is pure GPU pressure on exactly the pages where the freeze bites. */}
        <Rect
          name="slide-bg"
          x={0}
          y={0}
          width={slide.widthPx}
          height={slide.heightPx}
          fill="#ffffff"
          {...(dense
            ? { stroke: 'rgba(0,0,0,0.15)', strokeWidth: 1 }
            : { shadowColor: 'rgba(0,0,0,0.15)', shadowBlur: 16, shadowOffsetY: 2 })}
        />
        <Rect
          name="slide-bg"
          x={0}
          y={0}
          width={slide.widthPx}
          height={slide.heightPx}
          {...bgFill(slide, images)}
        />
        {slide.nodes.map((n) => (
          <NodeView
            key={n.id}
            node={n}
            onSelect={onSelect}
            onEditText={onEditText}
            onTransform={onTransform}
            onEditTableCell={onEditTableCell}
            onTableCellPress={onTableCellPress}
            onPlayMedia={onPlayMedia}
            onDragGuides={(g, sp) => {
              const nextSp = sp ?? []
              setGuides((prev) => (guidesEqual(prev, g) ? prev : g))
              setSpacing((prev) => (spacingEqual(prev, nextSp) ? prev : nextSp))
            }}
            snapTargets={snapTargets}
            images={images}
            editingText={editingText}
            livePreview={selectedIds.length === 1}
            zoom={zoom}
            multiDrag={selectedIds.length > 1 && selectedIds.includes(n.sourceId)}
            selectionDrag={selectionDrag}
            onSelectionDragPress={beginSelectionDrag}
            onSelectionDragEnd={endSelectionDrag}
            onSelectionDragMove={syncSelectionDrag}
            onDuplicateTo={onDuplicateTo}
            entered={n.sourceId === enteredGroupId}
            onEnterGroup={onEnterGroup}
            selectedIds={selectedIds}
            selBBox={selBBox}
            spacingBoxes={spacingBoxes}
            suppressClickRef={suppressClickRef}
            selStroke={selStroke}
            selHairline={hairline}
          />
        ))}
        {guides.map((g, i) =>
          g.axis === 'v' ? (
            <Line
              key={`v${i}`}
              points={[g.pos, 0, g.pos, slide.heightPx]}
              stroke="#ff2d55"
              strokeWidth={1}
              dash={[4, 4]}
            />
          ) : (
            <Line
              key={`h${i}`}
              points={[0, g.pos, slide.widthPx, g.pos]}
              stroke="#ff2d55"
              strokeWidth={1}
              dash={[4, 4]}
            />
          ),
        )}
        {/* Equal-spacing double-headed arrows (axis=x horizontal spacing, axis=y vertical) */}
        {spacing.map((s, i) => (
          <Arrow
            key={`sp${i}`}
            points={s.axis === 'x' ? [s.from, s.at, s.to, s.at] : [s.at, s.from, s.at, s.to]}
            stroke="#ff2d55"
            fill="#ff2d55"
            strokeWidth={1 / Math.max(zoom, 0.1)}
            pointerAtBeginning
            pointerAtEnding
            pointerLength={5 / Math.max(zoom, 0.1)}
            pointerWidth={4 / Math.max(zoom, 0.1)}
            listening={false}
          />
        ))}
        {/* Same-size match: matched elements highlighted with a dashed outline */}
        {sizeMatch &&
          [...new Set([...sizeMatch.w, ...sizeMatch.h])].map((id) => {
            const n = slide.nodes.find((x) => x.sourceId === id)
            if (!n) return null
            return (
              <Rect
                key={`sz_${id}`}
                x={n.box.x + n.box.w / 2}
                y={n.box.y + n.box.h / 2}
                offsetX={n.box.w / 2 + 2}
                offsetY={n.box.h / 2 + 2}
                width={n.box.w + 4}
                height={n.box.h + 4}
                rotation={n.box.rotationDeg}
                stroke="#ff2d55"
                strokeWidth={1 / Math.max(zoom, 0.1)}
                dash={[4, 3]}
                listening={false}
              />
            )
          })}
        {marquee && (
          <Rect
            x={Math.min(marquee.x1, marquee.x2)}
            y={Math.min(marquee.y1, marquee.y2)}
            width={Math.abs(marquee.x2 - marquee.x1)}
            height={Math.abs(marquee.y2 - marquee.y1)}
            fill="rgba(64,128,255,0.08)"
            stroke="#4080ff"
            strokeWidth={1 / Math.max(zoom, 0.1)}
            dash={[4, 4]}
            listening={false}
          />
        )}
        {/* Shape-draw preview: live ghost of the actual shape being drawn (PowerPoint drag behavior) */}
        {drawMode &&
          drawPreview &&
          (() => {
            const p = drawPreview
            // Line kinds preview with their real endpoints (flips encode leftward/upward drags)
            const sx = p.flipH ? p.x + p.w : p.x
            const sy = p.flipV ? p.y + p.h : p.y
            const ex = p.flipH ? p.x : p.x + p.w
            const ey = p.flipV ? p.y : p.y + p.h
            const lineW = 1.33 // 1pt, same as the inserted stroke
            if (isStraightLineKind(drawMode.kind)) {
              return drawMode.kind === 'line' ? (
                <Line
                  points={[sx, sy, ex, ey]}
                  stroke="#000000"
                  strokeWidth={lineW}
                  listening={false}
                />
              ) : (
                <Arrow
                  points={[sx, sy, ex, ey]}
                  stroke="#000000"
                  fill="#000000"
                  strokeWidth={lineW}
                  pointerAtEnding
                  pointerAtBeginning={drawMode.kind === 'lineArrowDouble'}
                  pointerLength={8}
                  pointerWidth={6}
                  listening={false}
                />
              )
            }
            if (isLineDrawKind(drawMode.kind)) {
              const mx = (sx + ex) / 2
              const d =
                drawMode.kind === 'lineBent'
                  ? `M ${sx} ${sy} L ${mx} ${sy} L ${mx} ${ey} L ${ex} ${ey}`
                  : `M ${sx} ${sy} C ${sx + (ex - sx) * 0.6} ${sy} ${sx + (ex - sx) * 0.4} ${ey} ${ex} ${ey}`
              return <Path data={d} stroke="#000000" strokeWidth={lineW} listening={false} />
            }
            const d = shapePreviewPath(drawMode.kind, Math.max(p.w, 1), Math.max(p.h, 1))
            if (d)
              return (
                <Path
                  x={p.x}
                  y={p.y}
                  data={d}
                  fill="rgba(196,62,28,0.45)"
                  stroke="#C43E1C"
                  strokeWidth={1 / Math.max(zoom, 0.1)}
                  listening={false}
                />
              )
            // Preset not covered by the preview geometry: dashed box fallback
            return (
              <Rect
                x={p.x}
                y={p.y}
                width={p.w}
                height={p.h}
                stroke="#4080ff"
                strokeWidth={1 / Math.max(zoom, 0.1)}
                dash={[4, 4]}
                listening={false}
              />
            )
          })()}
        {slide.nodes.flatMap((n) => {
          if (n.type !== 'table' || n.decoration || n.box.rotationDeg) return []
          const tbl = n as TableRenderNode
          // One grip per grid boundary; only boundaries a merged cell spans across are skipped
          const spansX = (b: number) => tbl.cells.some((c) => c.x < b - 0.5 && c.x + c.w > b + 0.5)
          const spansY = (b: number) => tbl.cells.some((c) => c.y < b - 0.5 && c.y + c.h > b + 0.5)
          // rtl tables render mirrored: gridX stays logical, so visual boundary = width − gridX
          const tblGridW = tbl.gridX[tbl.gridX.length - 1] ?? n.box.w
          const visX = (i: number) => (tbl.rtl ? tblGridW - tbl.gridX[i]! : tbl.gridX[i]!)
          // While a border is dragged the live preview re-grids the table under the pointer: the
          // limits and the pointer→size conversion keep reading the table as the press found it
          const held = (kind: 'col' | 'row', idx: number) => {
            const g = gripRef.current
            return g && g.id === n.sourceId && g.kind === kind && g.idx === idx ? g : null
          }
          // Scale grip hit area so it is easy to hover and click across all zoom levels
          const gripHit = Math.max(10 / Math.max(zoom, 0.1), 8)
          const gripHalf = gripHit / 2
          const edge = Math.max(3 / Math.max(zoom, 0.1), 3)
          const colWant = (g: GripDrag, x: number) =>
            tbl.rtl ? g.edge - (x + gripHalf) : x + gripHalf - g.origin
          const rowWant = (g: GripDrag, y: number) => y + gripHalf - g.origin
          const lastCol = tbl.gridX.length - 2
          const lastRow = tbl.gridY.length - 2
          return [
            ...tbl.gridX.slice(1).flatMap((_, col) =>
              col === lastCol || spansX(visX(col + 1)) || n.box.h - 2 * edge <= 0
                ? []
                : [
                    <Rect
                      key={`colgrip_${n.sourceId}_${col}`}
                      x={n.box.x + visX(col + 1) - gripHalf}
                      y={n.box.y + edge}
                      width={gripHit}
                      height={Math.max(1, n.box.h - 2 * edge)}
                      fill="transparent"
                      draggable
                      dragBoundFunc={(pos) => {
                        // the grip stops where a neighbouring column would start cutting a word apart
                        const g = held('col', col)
                        const widths =
                          g?.sizes ?? tbl.gridX.slice(1).map((v, i) => v - tbl.gridX[i]!)
                        const { lo, hi } = colDragRange(widths, col, g?.needs ?? tbl.colNeedPx)
                        const start =
                          CANVAS_BLEED + (g?.origin ?? n.box.x + tbl.gridX[col]!) - gripHalf
                        const x = tbl.rtl
                          ? pos.x
                          : Math.min(Math.max(pos.x, start + lo), start + hi)
                        return { x, y: n.box.y + edge + CANVAS_BLEED }
                      }}
                      onMouseEnter={(e) => {
                        const st = e.target.getStage()
                        if (st) st.container().style.cursor = 'col-resize'
                      }}
                      onMouseMove={(e) => {
                        const st = e.target.getStage()
                        if (st) st.container().style.cursor = 'col-resize'
                      }}
                      onMouseLeave={(e) => {
                        const st = e.target.getStage()
                        if (st && !gripRef.current) st.container().style.cursor = 'default'
                      }}
                      onMouseDown={() => {
                        if (!selectedIds.includes(n.sourceId)) {
                          onSelect(n.sourceId, false)
                        }
                      }}
                      onDragStart={(e) => {
                        gripRef.current = {
                          id: n.sourceId,
                          kind: 'col',
                          idx: col,
                          sizes: tbl.gridX.slice(1).map((v, i) => v - tbl.gridX[i]!),
                          needs: tbl.colNeedPx,
                          origin: n.box.x + tbl.gridX[col]!,
                          edge: n.box.x + visX(col),
                          home: n.box.x + visX(col + 1) - gripHalf,
                          cancelled: false,
                        }
                        const st = e.target.getStage()
                        if (st) st.container().style.cursor = 'col-resize'
                        armGripEscape(e.target)
                      }}
                      onDragMove={(e) => {
                        const g = held('col', col)
                        if (!g) return
                        const st = e.target.getStage()
                        if (st) st.container().style.cursor = 'col-resize'
                        // rtl: the grip is the column's visual-left edge; width grows leftward
                        const r = resolveColDrag(g.sizes, col, colWant(g, e.target.x()), g.needs)
                        onTableGridMove?.(n.sourceId, {
                          kind: 'col',
                          col,
                          wPx: r.wPx,
                          ...(r.neighbor ? { neighbor: r.neighbor } : {}),
                        })
                      }}
                      onDragEnd={(e) => {
                        const g = held('col', col)
                        disarmGripEscape()
                        gripRef.current = null
                        const st = e.target.getStage()
                        if (st) st.container().style.cursor = 'default'
                        if (!g) return
                        if (g.cancelled) {
                          e.target.x(g.home)
                          onTableGridCancel?.()
                          return
                        }
                        const r = resolveColDrag(g.sizes, col, colWant(g, e.target.x()), g.needs)
                        onTableColResize(n.sourceId, col, r.wPx, r.neighbor)
                      }}
                    />,
                  ],
            ),
            ...(onTableRowResize
              ? tbl.gridY.slice(1).flatMap((b, row) =>
                  row === lastRow || spansY(b) || n.box.w - 2 * edge <= 0
                    ? []
                    : [
                        <Rect
                          key={`rowgrip_${n.sourceId}_${row}`}
                          x={n.box.x + edge}
                          y={n.box.y + b - gripHalf}
                          width={Math.max(1, n.box.w - 2 * edge)}
                          height={gripHit}
                          fill="transparent"
                          draggable
                          dragBoundFunc={(pos) => {
                            // a row stops where its wrapped text starts to need the height (the engine
                            // would grow it straight back, which would show as a snap on release)
                            const g = held('row', row)
                            const cur = g ? g.sizes[row]! : tbl.gridY[row + 1]! - tbl.gridY[row]!
                            const need = g ? g.needs?.[row] : tbl.rowNeedPx?.[row]
                            const start =
                              CANVAS_BLEED + (g?.origin ?? n.box.y + tbl.gridY[row]!) - gripHalf
                            return {
                              x: n.box.x + edge + CANVAS_BLEED,
                              y: Math.max(pos.y, start + rowDragFloor(cur, need)),
                            }
                          }}
                          onMouseEnter={(e) => {
                            const st = e.target.getStage()
                            if (st) st.container().style.cursor = 'row-resize'
                          }}
                          onMouseMove={(e) => {
                            const st = e.target.getStage()
                            if (st) st.container().style.cursor = 'row-resize'
                          }}
                          onMouseLeave={(e) => {
                            const st = e.target.getStage()
                            if (st && !gripRef.current) st.container().style.cursor = 'default'
                          }}
                          onMouseDown={() => {
                            if (!selectedIds.includes(n.sourceId)) {
                              onSelect(n.sourceId, false)
                            }
                          }}
                          onDragStart={(e) => {
                            gripRef.current = {
                              id: n.sourceId,
                              kind: 'row',
                              idx: row,
                              sizes: tbl.gridY.slice(1).map((v, i) => v - tbl.gridY[i]!),
                              needs: tbl.rowNeedPx,
                              origin: n.box.y + tbl.gridY[row]!,
                              edge: 0,
                              home: n.box.y + b - gripHalf,
                              cancelled: false,
                            }
                            const st = e.target.getStage()
                            if (st) st.container().style.cursor = 'row-resize'
                            armGripEscape(e.target)
                          }}
                          onDragMove={(e) => {
                            const g = held('row', row)
                            if (!g) return
                            const st = e.target.getStage()
                            if (st) st.container().style.cursor = 'row-resize'
                            onTableGridMove?.(n.sourceId, {
                              kind: 'row',
                              row,
                              hPx: resolveRowDrag(
                                rowWant(g, e.target.y()),
                                g.sizes[row]!,
                                g.needs?.[row],
                              ),
                            })
                          }}
                          onDragEnd={(e) => {
                            const g = held('row', row)
                            disarmGripEscape()
                            gripRef.current = null
                            const st = e.target.getStage()
                            if (st) st.container().style.cursor = 'default'
                            if (!g) return
                            if (g.cancelled) {
                              e.target.y(g.home)
                              onTableGridCancel?.()
                              return
                            }
                            const newH = resolveRowDrag(
                              rowWant(g, e.target.y()),
                              g.sizes[row]!,
                              g.needs?.[row],
                            )
                            onTableRowResize(n.sourceId, row, newH)
                          }}
                        />,
                      ],
                )
              : []),
          ]
        })}
        {/* Connector selection highlight: dashed bounding box (no transform handles) */}
        {selectedIds.flatMap((id) => {
          const n = slide.nodes.find((x) => x.sourceId === id)
          if (!n || !isConnectorNode(n)) return []
          const b = n.box
          return [
            <Rect
              key={`cxn_sel_${id}`}
              x={b.x + b.w / 2}
              y={b.y + b.h / 2}
              offsetX={b.w / 2 + 3}
              offsetY={b.h / 2 + 3}
              width={b.w + 6}
              height={b.h + 6}
              rotation={b.rotationDeg}
              stroke={selStroke}
              strokeWidth={hairline}
              dash={[4, 3]}
              fill="transparent"
              listening={false}
            />,
          ]
        })}
        {/* Connector endpoint handles: drag an end to reposition it; near a shape's side midpoint it snaps and attaches (stCxn/endCxn) */}
        {onEditConnectorEndpoints &&
          selectedIds.length === 1 &&
          (() => {
            const n = slide.nodes.find((x) => x.sourceId === selectedIds[0])
            if (!n || !isConnectorNode(n) || n.decoration || n.box.rotationDeg) return null
            const line = (n as ShapeRenderNode).line
            if (!line || line.points.length < 4) return null
            return (
              <ConnectorEndpointHandles
                node={n as ShapeRenderNode}
                slide={slide}
                zoom={zoom}
                stroke={selStroke}
                hairline={hairline}
                onCommit={onEditConnectorEndpoints}
              />
            )
          })()}
      </Layer>
      {/* Selection chrome on its own layer: counter-scaling it during a zoom gesture then
          only redraws these few nodes — redrawing the content layer per frame is what
          made the gesture stutter. The Transformer tracks nodes across layers fine.
          Mirrors the content layer's draw-mode listening switch so the crosshair
          gesture isn't swallowed by the transformer handles. */}
      <Layer x={CANVAS_BLEED} y={CANVAS_BLEED} listening={!drawMode}>
        <Transformer
          ref={trRef}
          onDragMove={() => {
            if (selectedIds.length > 1) syncSelectionDrag()
          }}
          onDragEnd={() => {
            if (selectedIds.length > 1) void endSelectionDrag()
          }}
          rotateEnabled
          // Rotating locks onto the 45° multiples when within a few degrees (PowerPoint-style snap)
          rotationSnaps={ROTATION_SNAPS}
          rotationSnapTolerance={ROTATION_SNAP_TOLERANCE}
          // Pictures default to proportional scaling (corner handles); shapes/text boxes scale freely
          keepRatio={
            selectedIds.length > 0 &&
            selectedIds.every((id) => findNodeDeep(slide.nodes, id)?.type === 'picture')
          }
          borderStroke={selStroke}
          anchorStroke={selStroke}
          anchorFill="#ffffff"
          // The canvas is CSS-scaled: divide every chrome size by zoom so the frame keeps a constant on-screen weight
          borderStrokeWidth={hairline}
          anchorStrokeWidth={hairline}
          anchorSize={(8 * chromeScale) / Math.max(zoom, 0.1)}
          rotateAnchorOffset={(50 * chromeScale) / Math.max(zoom, 0.1)}
          anchorStyleFunc={(a) =>
            styleRotaterAnchor(a, (ROTATER_SIZE * chromeScale) / Math.max(zoom, 0.1))
          }
          rotateAnchorCursor={rotateCursor}
          // Tiny selections keep only the corner handles (middle ones would overlap them)
          enabledAnchors={liveAnchors()}
          onTransformStart={() => {
            const a = (trRef.current as unknown as { _movingAnchorName?: string } | null)
              ?._movingAnchorName
            gestureMinRef.current = { anchor: a ?? '', min: computeSelectionMin(a) }
            trActiveRef.current = true
          }}
          // Dragging a handle past the opposite edge must stop at a minimum size, not mirror the element
          flipEnabled={false}
          // Keep the dragged handle inside its legal range so the box can never go through the opposite edge
          // (Konva would otherwise swap handles mid-drag and re-form the box on the far side)
          anchorDragBoundFunc={(_old, pos) => {
            const tr = trRef.current as unknown as {
              _movingAnchorName?: string
              _getNodeRect?: () => AnchorRect
            } | null
            const name = tr?._movingAnchorName
            const rect = tr?._getNodeRect?.()
            return name && rect ? clampAnchorPos(name, pos, rect, selectionMin(name)) : pos
          }}
          boundBoxFunc={(old, next) => {
            const min = selectionMin(
              (trRef.current as unknown as { _movingAnchorName?: string } | null)
                ?._movingAnchorName,
            )
            const clamped = clampTransformBox(old, next, min)
            if (clamped !== next) return clamped
            // Same-size snapping: when resizing a single, unrotated, non-picture element (keepRatio conflicts),
            // snap when width/height is close to another element's and highlight the matched element
            if (
              selectedIds.length !== 1 ||
              next.rotation !== old.rotation ||
              Math.abs(next.width - old.width) + Math.abs(next.height - old.height) < 1e-6
            )
              return next
            const sel = slide.nodes.find((n) => n.sourceId === selectedIds[0])
            if (!sel || sel.box.rotationDeg || sel.type === 'picture' || isConnectorNode(sel))
              return next
            const thr = 6 / Math.max(zoom, 0.1)
            const cands = slide.nodes.filter((n) => !n.decoration && n.sourceId !== sel.sourceId)
            const nearest = (val: number, dim: 'w' | 'h') => {
              let best: number | null = null
              for (const c of cands) {
                const v = c.box[dim]
                if (
                  Math.abs(v - val) <= thr &&
                  (best == null || Math.abs(v - val) < Math.abs(best - val))
                )
                  best = v
              }
              return best
            }
            const out = { ...next }
            const match = { w: [] as string[], h: [] as string[] }
            const w = nearest(next.width, 'w')
            if (w != null) {
              // When dragging a left-side anchor keep the right edge fixed (x compensates with width)
              if (Math.abs(next.x - old.x) > 1e-6) out.x = next.x + next.width - w
              out.width = w
              match.w = cands.filter((c) => Math.abs(c.box.w - w) < 0.5).map((c) => c.sourceId)
            }
            const h = nearest(next.height, 'h')
            if (h != null) {
              if (Math.abs(next.y - old.y) > 1e-6) out.y = next.y + next.height - h
              out.height = h
              match.h = cands.filter((c) => Math.abs(c.box.h - h) < 0.5).map((c) => c.sourceId)
            }
            const key = `${match.w.join(',')}|${match.h.join(',')}`
            if (key !== sizeMatchKeyRef.current) {
              sizeMatchKeyRef.current = key
              setSizeMatch(w != null || h != null ? match : null)
            }
            return out
          }}
          onTransformEnd={() => {
            trActiveRef.current = false
            gestureMinRef.current = null
            sizeMatchKeyRef.current = ''
            setSizeMatch(null)
          }}
        />
        {onAdjust &&
          selectedIds.length === 1 &&
          !editingText &&
          (() => {
            // Top-level shapes only: entered-group children live in group-local
            // coordinates, which the handle math doesn't map yet
            const n = slide.nodes.find((x) => x.sourceId === selectedIds[0])
            if (!n || (n.type !== 'shape' && n.type !== 'text') || isConnectorNode(n)) return null
            const id = n.sourceId
            return (
              <AdjustHandles
                node={n as ShapeRenderNode}
                zoom={zoom}
                chromeScale={chromeScale}
                onAdjust={(adjust, preview) => onAdjust(id, adjust, preview)}
              />
            )
          })()}
      </Layer>
    </Stage>
  )
}

export { computeSnap }

// ── Connector endpoint handles ─────────────────────────────

interface AnchorPt {
  targetId: string
  idx: number
  x: number
  y: number
}

/**
 * Attachable connection points of the other elements: side midpoints of each box
 * (idx matches the engine's rectangle approximation — 0 top, 1 left, 2 bottom, 3 right).
 * Uses the unrotated model box, same as the engine's move-following computation.
 */
function connectorAnchors(slide: RenderSlide, excludeId: string): AnchorPt[] {
  const out: AnchorPt[] = []
  for (const n of slide.nodes) {
    if (n.decoration || n.sourceId === excludeId) continue
    if (n.type === 'placeholder-chip' || isConnectorNode(n)) continue
    const b = n.box
    out.push(
      { targetId: n.sourceId, idx: 0, x: b.x + b.w / 2, y: b.y },
      { targetId: n.sourceId, idx: 1, x: b.x, y: b.y + b.h / 2 },
      { targetId: n.sourceId, idx: 2, x: b.x + b.w / 2, y: b.y + b.h },
      { targetId: n.sourceId, idx: 3, x: b.x + b.w, y: b.y + b.h / 2 },
    )
  }
  return out
}

function ConnectorEndpointHandles({
  node,
  slide,
  zoom,
  stroke,
  hairline,
  onCommit,
}: {
  node: ShapeRenderNode
  slide: RenderSlide
  zoom: number
  /** Selection chrome color (flips to white on dark slide backgrounds) */
  stroke: string
  /** Selection-chrome stroke width (canvas px, zoom- and raster-compensated) */
  hairline: number
  onCommit: NonNullable<Props['onEditConnectorEndpoints']>
}) {
  const [drag, setDrag] = useState<{
    which: 'start' | 'end'
    x: number
    y: number
    snap: AnchorPt | null
  } | null>(null)
  const z = Math.max(zoom, 0.1)
  const pts = node.line!.points
  const start = { x: node.box.x + pts[0]!, y: node.box.y + pts[1]! }
  const end = { x: node.box.x + pts[pts.length - 2]!, y: node.box.y + pts[pts.length - 1]! }
  const anchors = useMemo(() => connectorAnchors(slide, node.sourceId), [slide, node.sourceId])
  const thr = 8 / z
  const nearestAnchor = (x: number, y: number): AnchorPt | null => {
    let best: AnchorPt | null = null
    let bestD = thr
    for (const a of anchors) {
      const d = Math.hypot(a.x - x, a.y - y)
      if (d <= bestD) {
        bestD = d
        best = a
      }
    }
    return best
  }
  const handle = (which: 'start' | 'end', p: { x: number; y: number }) => (
    <Circle
      key={which}
      x={p.x}
      y={p.y}
      radius={4.5 / z}
      fill="#ffffff"
      stroke={stroke}
      strokeWidth={1.5 * hairline}
      hitStrokeWidth={12 / z}
      draggable
      onMouseEnter={(e) => setStageCursor(e, 'crosshair')}
      onMouseLeave={(e) => setStageCursor(e, 'default')}
      onDragMove={(e) => {
        const t = e.target
        const snap = nearestAnchor(t.x(), t.y())
        if (snap) t.position({ x: snap.x, y: snap.y })
        setDrag({ which, x: t.x(), y: t.y(), snap })
      }}
      onDragEnd={(e) => {
        const t = e.target
        const x = t.x()
        const y = t.y()
        const snap = nearestAnchor(x, y)
        // Snap the Konva handle back; the model commit re-renders the connector at its new geometry
        t.position(p)
        setDrag(null)
        const att = snap ? { targetId: snap.targetId, idx: snap.idx } : null
        onCommit(
          node.sourceId,
          which === 'start'
            ? { x1: x, y1: y, x2: end.x, y2: end.y, start: att }
            : { x1: start.x, y1: start.y, x2: x, y2: y, end: att },
        )
      }}
    />
  )
  const fixed = drag?.which === 'start' ? end : start
  return (
    <>
      {/* While dragging: show all attachable anchor dots, highlight the snapped one, and preview the new run as a dashed line */}
      {drag &&
        anchors.map((a) => {
          const active = drag.snap && a.targetId === drag.snap.targetId && a.idx === drag.snap.idx
          return (
            <Circle
              key={`anchor_${a.targetId}_${a.idx}`}
              x={a.x}
              y={a.y}
              radius={(active ? 4 : 2.5) / z}
              fill={active ? '#4ea72e' : '#8e8e93'}
              listening={false}
            />
          )
        })}
      {drag && (
        <Line
          points={[fixed.x, fixed.y, drag.x, drag.y]}
          stroke={stroke}
          strokeWidth={hairline}
          dash={[4, 3]}
          listening={false}
        />
      )}
      {handle('start', start)}
      {handle('end', end)}
    </>
  )
}

function bgFill(slide: RenderSlide, images: Map<string, HTMLImageElement>) {
  const f = fillToKonva(slide.background, slide.widthPx, slide.heightPx, images)
  return Object.keys(f).length ? f : { fill: '#ffffff' }
}

interface NodeProps {
  selectionDrag?: SelectionDrag
  onSelectionDragPress?: (sourceId: string) => void
  onSelectionDragEnd?: () => Promise<void>
  onSelectionDragMove?: () => void
  node: RenderNode
  onSelect: (id: string | null, additive?: boolean) => void
  onEditText: (id: string, caret?: EditCaret) => void
  onTransform: Props['onTransform']
  onEditTableCell: Props['onEditTableCell']
  onTableCellPress?: Props['onTableCellPress']
  onPlayMedia?: Props['onPlayMedia']
  onDragGuides: (g: Guide[], spacing?: SpacingIndicator[]) => void
  snapTargets: (excludeIds: string[]) => SnapTarget[]
  /** Neighbor boxes for equal-spacing snapping (excluding the dragged selection) */
  spacingBoxes?: (excludeIds: string[]) => Array<{ x: number; y: number; w: number; h: number }>
  /** Bounding box of a multi-select drag (snapped as a whole) */
  selBBox?: { x: number; y: number; w: number; h: number } | null
  images: Map<string, HTMLImageElement>
  editingText?: Props['editingText']
  /** Live resize preview only for single selection (multi-select gestures commit per node; the preview's undo-merge semantics don't hold, so it's off) */
  livePreview: boolean
  zoom: number
  /** This node is part of a multi-select drag: disable per-node snapping (each node snapping separately would break relative positions within the selection) */
  multiDrag?: boolean
  onDuplicateTo?: Props['onDuplicateTo']
  /** This group is in in-group editing mode: children are interactive, whole-group hit area is off */
  entered?: boolean
  onEnterGroup?: Props['onEnterGroup']
  /** In-group editing: this node is a direct child of the group (geometry commits carry groupId; snapping/preview disabled) */
  insideGroupId?: string
  /** Children may double-click into text editing (only when the group is unrotated/unflipped/unscaled — the overlay can't be positioned otherwise) */
  allowChildTextEdit?: boolean
  /** Used by multiDrag computation for in-group-editing children */
  selectedIds?: string[]
  /** Set when a marquee drag just completed on this gesture: the trailing click must not select the node under the cursor */
  suppressClickRef?: React.MutableRefObject<boolean>
  /** Selection chrome color (flips to white on dark slide backgrounds) */
  selStroke?: string
  /** Selection-chrome stroke width (canvas px, zoom- and raster-compensated) */
  selHairline?: number
}

const boxKey = (b: { x: number; y: number; w: number; h: number; rotationDeg?: number }): string =>
  `${b.x}|${b.y}|${b.w}|${b.h}|${b.rotationDeg}`

/** Shortest gap between two live resize previews (ms): each one is an IPC round-trip + full-page relayout,
 * and only one is in flight at a time (the reply carries the layout of the pointer's size at send time) */
const PREVIEW_MIN_GAP_MS = 12

function NodeView({
  selectionDrag,
  onSelectionDragPress,
  onSelectionDragEnd,
  onSelectionDragMove,
  node,
  onSelect,
  onEditText,
  onTransform,
  onEditTableCell,
  onTableCellPress,
  onPlayMedia,
  onDragGuides,
  snapTargets,
  spacingBoxes,
  selBBox,
  images,
  editingText,
  livePreview,
  zoom,
  multiDrag,
  onDuplicateTo,
  entered,
  onEnterGroup,
  insideGroupId,
  allowChildTextEdit,
  selectedIds,
  suppressClickRef,
  selStroke = '#232425',
  selHairline,
}: NodeProps) {
  const { t } = useI18n()
  // screenshot automation (fidelity-compare) hides the hint to match print output; the
  // event-backed store re-renders even when no slide switch follows (single-slide decks)
  const hidePhPrompts = useSyncExternalStore(subscribeHidePhPrompts, getHidePhPrompts)
  const { box } = node
  const groupRef = useRef<Konva.Group>(null)
  /** Most recently rendered model width/height (preview responses update it; during a gesture, scale is always relative to it) */
  const boxRef = useRef({ w: box.w, h: box.h })
  const transformingRef = useRef(false)
  const pivotKeyRef = useRef('')
  const lastPreviewRef = useRef(0)
  /** The pointer's size of the running resize (the node itself may be drawn taller, see the table path below) */
  const rawSizeRef = useRef<{ w: number; h: number } | null>(null)
  /** One preview request at a time: the next one carries the pointer's latest size */
  const previewBusyRef = useRef(false)
  const previewPendingRef = useRef(false)
  /** Size asked for by the preview whose reply landed last (a reply taller than that is content-bound) */
  const repliedAskRef = useRef<{ w: number; h: number } | null>(null)
  /** Height the latest content-bound table layout cannot go under (0 = none) */
  const contentFloorRef = useRef(0)
  const gestureIdRef = useRef(0)
  /** Latest model box (the committed layout of a table resize lands here while the drawn size is held) */
  const modelBoxRef = useRef(box)
  /** Finishes a resize whose drawn size is held until the committed layout lands (see onTransformEnd) */
  const settleRef = useRef<{ key: string; armed: boolean; run: () => void } | null>(null)
  /** Name of the Transformer handle being dragged ('bottom-right', …) */
  const movingAnchor = (): string | undefined =>
    (
      groupRef.current?.getStage()?.find<Konva.Transformer>('Transformer')[0] as unknown as {
        _movingAnchorName?: string
      } | null
    )?._movingAnchorName
  /** Node position captured on every transform event: boxPivotProps derives Konva x/y from
   * box.w/h, so a live-preview re-render would otherwise teleport the node mid-gesture. */
  const gesturePosRef = useRef<{ x: number; y: number } | null>(null)
  // Option+drag ghost: while Alt is held mid-drag, the dragged node becomes the semi-transparent
  // "copy" (dashed frame) and a static render of the original stays at the source position.
  const [altDragging, setAltDragging] = useState(false)
  const altDraggingRef = useRef(false)
  const setAltDrag = (on: boolean) => {
    if (altDraggingRef.current === on) return
    altDraggingRef.current = on
    setAltDragging(on)
  }
  /** Window key listeners active during a drag, so pressing/releasing Alt with a still pointer
   * still toggles the ghost (drop semantics keep reading Alt from the mouse event on release). */
  const altKeyCleanupRef = useRef<(() => void) | null>(null)
  /** Live drag position: the ghost toggle re-renders mid-drag, and the controlled x/y from
   * boxPivotProps must not teleport the node back to the model position (same idea as gesturePosRef). */
  const dragPosRef = useRef<{ x: number; y: number } | null>(null)
  /** Press landed inside a cell of a table that was not selected: that gesture selects cells, it must not move the table.
   * Holds the table's absolute position at press time (the drag-bound pins it there); null = a normal move gesture. */
  const tablePinRef = useRef<{ x: number; y: number } | null>(null)
  /** Press landed inside text of an editable shape: that gesture edits or selects text, it must not move the shape.
   * Holds the node's absolute position at press time (the drag-bound pins it there); null = a normal move gesture. */
  const textPinRef = useRef<{ x: number; y: number } | null>(null)
  useLayoutEffect(() => {
    const g = groupRef.current
    if (g && dragPosRef.current && g.isDragging()) g.position(dragPosRef.current)
  })
  useEffect(() => () => altKeyCleanupRef.current?.(), [])

  // The Transformer's frame/scale basis defaults to getClientRect() (content bounding box). When text
  // overflows the shape box (autofit off and content too tall), overflowing glyphs inflate the bounding
  // box; after shrinking, the Transformer refresh pops the frame back to content size, which looks like
  // "resize doesn't work". Match PowerPoint: handles always sit on the shape box (model box), text may
  // overflow — override getClientRect so the Transformer measures by the model box (live boxRef value).
  useEffect(() => {
    const g = groupRef.current
    if (!g) return
    g.getClientRect = (config?: { skipTransform?: boolean; relativeTo?: Konva.Node }) => {
      const { w, h } = boxRef.current
      const local = { x: 0, y: 0, width: w, height: h }
      if (config?.skipTransform) return local
      const t = g.getAbsoluteTransform(config?.relativeTo)
      const pts = [
        t.point({ x: 0, y: 0 }),
        t.point({ x: w, y: 0 }),
        t.point({ x: w, y: h }),
        t.point({ x: 0, y: h }),
      ]
      const xs = pts.map((p) => p.x)
      const ys = pts.map((p) => p.y)
      const minX = Math.min(...xs)
      const minY = Math.min(...ys)
      return { x: minX, y: minY, width: Math.max(...xs) - minX, height: Math.max(...ys) - minY }
    }
  }, [])

  // Resize-preview response arrived (children already rendered at the new box): rebase the gesture's
  // temporary scale onto the new box so visual size (box × scale) stays continuous — frame follows the hand, content doesn't stretch or flicker.
  // Position must be restored too: the controlled x/y (box center per boxPivotProps) depend on
  // box.w/h, so the re-render would otherwise re-pin the node to the stale model x/y.
  useLayoutEffect(() => {
    const g = groupRef.current
    const prev = boxRef.current
    if (g && transformingRef.current && (prev.w !== box.w || prev.h !== box.h)) {
      const raw = rawSizeRef.current
      if (node.type === 'table' && raw && !box.flipH && !box.flipV) {
        // A table's rows grow with re-wrapped text: the layout that came back may be taller than the
        // pointer asked for. Draw it at its own height (the frame follows the content) instead of
        // squashing it into the pointer's — squashing is what made it pop to full height on release.
        contentFloorRef.current = contentBoundHeight(repliedAskRef.current?.h, box.h)
        const curW = prev.w * Math.abs(g.scaleX())
        const curH = prev.h * Math.abs(g.scaleY())
        const nextW = raw.w
        const nextH = drawnTableHeight(raw.h, contentFloorRef.current)
        if (box.w > 0) g.scaleX(nextW / box.w)
        if (box.h > 0) g.scaleY(nextH / box.h)
        const pos = gesturePosRef.current ?? { x: g.x(), y: g.y() }
        const d = centerShift(
          movingAnchor(),
          g.rotation(),
          { w: curW, h: curH },
          { w: nextW, h: nextH },
        )
        g.position({ x: pos.x + d.dx, y: pos.y + d.dy })
        gesturePosRef.current = { x: g.x(), y: g.y() }
        const tr = g.getStage()?.find<Konva.Transformer>('Transformer')[0]
        tr?.forceUpdate()
      } else {
        if (box.w > 0) g.scaleX(g.scaleX() * (prev.w / box.w))
        if (box.h > 0) g.scaleY(g.scaleY() * (prev.h / box.h))
        if (gesturePosRef.current) g.position(gesturePosRef.current)
      }
    }
    // keep glyphs unstretched across preview responses (the rebase above changed the group scale)
    if (
      g &&
      transformingRef.current &&
      (node.type === 'shape' || node.type === 'text' || node.type === 'table')
    )
      unstretchGlyphLayer(g)
    // A committed box can differ from where the gesture left the group (a table's rows grow with
    // wrapped text, an autofit text box re-measures): the controlled x/y props only change when the
    // model box does, so a position the gesture restored imperatively is never re-applied and the
    // drawn node stays offset from the model — and from every overlay pinned to the model box
    // (cell-selection layer, AI button). Re-pin the node to its model pivot whenever the model box
    // changes outside a gesture.
    const key = `${box.x}|${box.y}|${box.w}|${box.h}|${box.rotationDeg}`
    boxRef.current = { w: box.w, h: box.h }
    modelBoxRef.current = box
    // the committed layout of a held table resize arrived: let go of the gesture's drawn size
    if (settleRef.current?.armed && boxKey(box) !== settleRef.current.key) settleRef.current.run()
    // `key` is only marked applied outside a gesture: the preview frames of a resize already carry
    // the final box, so the committed render must still re-pin the node after the gesture ended.
    if (g && key !== pivotKeyRef.current && !transformingRef.current && !g.isDragging()) {
      pivotKeyRef.current = key
      const pv = boxPivotProps(box)
      if (Math.abs(g.x() - pv.x) > 1e-3 || Math.abs(g.y() - pv.y) > 1e-3)
        g.position({ x: pv.x, y: pv.y })
      // The Transformer measures this group through the getClientRect override above, which reads
      // boxRef — updated only now, after Konva already applied the new props (and fired the change
      // events the Transformer refreshes on). Poke it once more so its frame matches the settled
      // box (a table whose rows grew, an autofit text box that re-measured).
      g.fire('heightChange')
    }
  })

  // master/layout decoration layer: read-only display, not selectable/draggable
  if (node.decoration) return <StaticNode node={node} images={images} />

  // Chips are select-only; tables/charts support p:xfrm patch persistence, so they can be dragged/resized.
  // An unselected full-page background rubber-bands on drag (Stage-level marquee); selecting it first
  // makes it movable, so a full-bleed picture can still be repositioned like in PowerPoint.
  const selected = !!selectedIds?.includes(node.sourceId)
  const draggable = node.type !== 'placeholder-chip' && (!node.background || selected)
  const mathNode = (node.type === 'shape' || node.type === 'text') && !!node.hasEquation
  // Rotation/flip pivot on the box center (boxPivotProps): the Konva position IS the box
  // center, so model x/y = position − half size — including mid-gesture, since the offset
  // point stays the drawn box's center under any scale.

  /** Live layout preview: one request at a time, always carrying the pointer's latest size */
  const sendPreview = (): void => {
    const raw = rawSizeRef.current
    if (!raw) return
    if (previewBusyRef.current) {
      previewPendingRef.current = true
      return
    }
    const now = performance.now()
    if (now - lastPreviewRef.current < PREVIEW_MIN_GAP_MS) return
    lastPreviewRef.current = now
    previewBusyRef.current = true
    previewPendingRef.current = false
    const asked = { ...raw }
    const gid = gestureIdRef.current
    void Promise.resolve(
      onTransform(
        node.sourceId,
        { x: box.x, y: box.y, w: asked.w, h: asked.h, rotationDeg: box.rotationDeg },
        true,
      ),
    )
      .catch(() => undefined)
      .then(() => {
        if (gestureIdRef.current !== gid) return
        repliedAskRef.current = asked
        previewBusyRef.current = false
        lastPreviewRef.current = 0
        if (transformingRef.current && previewPendingRef.current) sendPreview()
      })
  }

  const editable = isEditableText(node) && (!insideGroupId || allowChildTextEdit)
  const clickOnText = (): boolean => {
    const g = groupRef.current
    const pos = g?.getStage()?.getPointerPosition()
    if (!g || !pos) return false
    const local = g.getAbsoluteTransform().copy().invert().point(pos)
    // NodeBody counter-flips the text, so mirror the point back into text coordinates
    const p = { x: box.flipH ? box.w - local.x : local.x, y: box.flipV ? box.h - local.y : local.y }
    return textHitAtPoint(node as ShapeRenderNode, box, p, 4 / Math.max(zoom, 0.1))
  }

  const common = {
    id: `node_${node.sourceId}`,
    ...boxPivotProps(box),
    draggable,
    dragBoundFunc: (pos: { x: number; y: number }) => {
      if (tablePinRef.current) return tablePinRef.current
      if (textPinRef.current) return textPinRef.current
      const pointer = groupRef.current?.getStage()?.getPointerPosition()
      return (multiDrag && pointer && selectionDrag?.bound(node.sourceId, pointer)) || pos
    },
    onMouseDown: (e: Konva.KonvaEventObject<MouseEvent>) => {
      if (e.target !== e.currentTarget && insideGroupId == null && node.type === 'group' && entered)
        return
      const onText = editable && !mathNode && clickOnText()
      if (onText) {
        if (groupRef.current) groupRef.current.draggable(false)
        textPinRef.current = groupRef.current?.absolutePosition() ?? null
      } else {
        textPinRef.current = null
      }
      if (e.evt.button === 0 && !onText) onSelectionDragPress?.(node.sourceId)
      tablePinRef.current = null
      if (node.type !== 'table' || insideGroupId || selected || !onTableCellPress) {
        // When already selected: a left-press on text immediately starts editing (PowerPoint/WPS)
        // so the user can drag-select text or place the caret without accidentally moving the shape
        const additive = e.evt.shiftKey || e.evt.metaKey || e.evt.ctrlKey || e.evt.altKey
        if (e.evt.button === 0 && !additive && selected && onText) {
          onEditText(node.sourceId, { x: e.evt.clientX, y: e.evt.clientY })
        }
        return
      }
      if (e.evt.button !== 0 || e.evt.shiftKey || e.evt.metaKey || e.evt.ctrlKey || e.evt.altKey)
        return
      const pos = e.target.getStage()?.getPointerPosition()
      if (!pos) return
      const local = tableLocalPointFromStage(pos, box, CANVAS_BLEED)
      // Cell borders (and the frame) stay a grip for moving the table; the cell body selects cells
      const cell = tableCellInterior(node, local, TABLE_BORDER_GRIP_PX / Math.max(zoom, 0.1))
      if (!cell) return
      tablePinRef.current = groupRef.current?.absolutePosition() ?? null
      onTableCellPress(node.sourceId, { row: cell.row, col: cell.col })
    },
    // A slight hand slip on click-select (3~5px is common on trackpads) shouldn't trigger a drag: Konva's
    // default 3px threshold is too sensitive, and once it becomes a drag, onDragMove snapping amplifies it into a visible 6px+ jump that commits to the model.
    // The threshold's semantics are "6 screen px": Konva compares in canvas coordinates, so divide by the canvas CSS zoom.
    dragDistance: 6 / Math.max(zoom, 0.1),
    // Over the text of an editable shape the pointer is an I-beam (a click there places the caret);
    // anywhere else on a movable node it is the move cursor, like PowerPoint/WPS.
    // The node is only draggable when the pointer is in the move area (perimeter/border), never while over text.
    onMouseMove: (e: Konva.KonvaEventObject<MouseEvent>) => {
      const onText = editable && !mathNode && clickOnText()
      const cursor = !draggable ? '' : onText ? 'text' : 'move'
      setStageCursor(e, cursor)
      if (groupRef.current && draggable) {
        groupRef.current.draggable(cursor === 'move')
      }
    },
    onMouseEnter: (e: Konva.KonvaEventObject<MouseEvent>) => {
      const onText = editable && !mathNode && clickOnText()
      const cursor = !draggable ? '' : onText ? 'text' : 'move'
      setStageCursor(e, cursor)
      if (groupRef.current && draggable) {
        groupRef.current.draggable(cursor === 'move')
      }
    },
    onMouseLeave: (e: Konva.KonvaEventObject<MouseEvent>) => {
      setStageCursor(e, '')
      textPinRef.current = null
      if (groupRef.current && draggable) {
        groupRef.current.draggable(true)
      }
    },
    onMouseUp: () => {
      textPinRef.current = null
      if (groupRef.current && draggable) {
        const onText = editable && !mathNode && clickOnText()
        groupRef.current.draggable(!onText)
      }
    },
    onClick: (e: Konva.KonvaEventObject<MouseEvent>) => {
      if (suppressClickRef?.current) {
        suppressClickRef.current = false
        return
      }
      // Konva fires click for every button, and on Windows the DOM contextmenu event
      // arrives only after mouseup — selecting here on a right click would collapse a
      // multi-selection before the menu reads it (menu built without Group, second
      // frame visibly dropped). Right-click selection is owned by onContextMenu, whose
      // guard keeps an existing multi-selection.
      if (e.evt.button !== 0) return
      const additive = e.evt.shiftKey || e.evt.metaKey
      // When clicking an already-selected element in a multi-selection without modifiers:
      // Keep the multi-selection intact so the user can click-to-drag without accidentally collapsing it
      // or popping open the text editor (double-click enters text editing).
      if (
        !additive &&
        selectedIds &&
        selectedIds.length > 1 &&
        selectedIds.includes(node.sourceId)
      ) {
        return
      }
      const wasSelected = selectedIds?.length === 1 && selectedIds[0] === node.sourceId
      onSelect(node.sourceId, additive)
      // When already selected: a single left click on the text starts editing with the caret at the click (PowerPoint/WPS).
      // When not yet selected: the first click selects the shape for dragging/nudging without immediately opening the text editor.
      if (!additive && wasSelected && editable && !mathNode && clickOnText())
        onEditText(node.sourceId, { x: e.evt.clientX, y: e.evt.clientY })
    },
    onTap: () => onSelect(node.sourceId),
    onDragStart: (e: Konva.KonvaEventObject<DragEvent>) => {
      dragPosRef.current = null
      if (tablePinRef.current || textPinRef.current) {
        e.target.stopDrag()
        return
      }
      // Konva stops routing mousemove to shapes while dragging: pin the move cursor here
      setStageCursor(e, 'move')
      if (!(onDuplicateTo && !multiDrag && !insideGroupId)) return
      const onKey = (ev: KeyboardEvent) => setAltDrag(ev.altKey)
      window.addEventListener('keydown', onKey)
      window.addEventListener('keyup', onKey)
      altKeyCleanupRef.current = () => {
        window.removeEventListener('keydown', onKey)
        window.removeEventListener('keyup', onKey)
        altKeyCleanupRef.current = null
      }
    },
    onDragMove: (e: Konva.KonvaEventObject<DragEvent>) => {
      if (tablePinRef.current || textPinRef.current) return
      // Alt held mid-drag = duplicate gesture: show the ghost (Alt can be pressed/released at any time during the drag)
      setAltDrag(!!(e.evt?.altKey && onDuplicateTo && !multiDrag && !insideGroupId))
      dragPosRef.current = { x: e.target.x(), y: e.target.y() }
      emitNodeGesture({
        sourceId: node.sourceId,
        kind: 'move',
        dx: e.target.x() - box.w / 2 - box.x,
        dy: e.target.y() - box.h / 2 - box.y,
      })
      // The shared bound function has already snapped the whole selection. Re-snapping
      // each proxy here introduces different pointer offsets after the first move.
      if (multiDrag && selectionDrag) {
        onSelectionDragMove?.()
        return
      }
      // Children in in-group editing use a different coordinate system from page snap targets; don't snap
      if (insideGroupId) {
        onDragGuides([])
        return
      }
      const t = e.target
      // Snap threshold semantics are "6 screen px": convert back by the canvas CSS zoom (same as dragDistance)
      const thr = 6 / Math.max(zoom, 0.1)
      const raw = { x: t.x() - box.w / 2, y: t.y() - box.h / 2 }
      // Multi-select drag: snap the selection bounding box as a whole. All selected
      // nodes fire their own dragmove under the Transformer's proxy drag, but each computes the same snap
      // delta from the same model bounding box + its own fixed offset -> relative positions stay unchanged.
      const bb = multiDrag
        ? selBBox && {
            x: selBBox.x + raw.x - box.x,
            y: selBBox.y + raw.y - box.y,
            w: selBBox.w,
            h: selBBox.h,
          }
        : { ...raw, w: box.w, h: box.h }
      if (!bb) {
        onDragGuides([])
        return
      }
      const exclude = multiDrag ? selectedIds! : [node.sourceId]
      const snap = computeSnap(bb, snapTargets(exclude), thr)
      let fx = snap.x
      let fy = snap.y
      // Axes not consumed by edge snapping then try equal-spacing (smart guides)
      const indicators: SpacingIndicator[] = []
      if (spacingBoxes && (fx == null || fy == null)) {
        const sp = computeSpacingSnap(
          { x: fx ?? bb.x, y: fy ?? bb.y, w: bb.w, h: bb.h },
          spacingBoxes(exclude),
          thr,
        )
        if (fx == null && sp.x != null) {
          fx = sp.x
          indicators.push(...sp.indicators.filter((i) => i.axis === 'x'))
        }
        if (fy == null && sp.y != null) {
          fy = sp.y
          indicators.push(...sp.indicators.filter((i) => i.axis === 'y'))
        }
      }
      if (fx != null) t.x(raw.x + (fx - bb.x) + box.w / 2)
      if (fy != null) t.y(raw.y + (fy - bb.y) + box.h / 2)
      dragPosRef.current = { x: t.x(), y: t.y() }
      emitNodeGesture({
        sourceId: node.sourceId,
        kind: 'move',
        dx: t.x() - box.w / 2 - box.x,
        dy: t.y() - box.h / 2 - box.y,
      })
      onDragGuides(snap.guides, indicators)
    },
    onDragEnd: (e: Konva.KonvaEventObject<DragEvent>) => {
      if (textPinRef.current) {
        textPinRef.current = null
        e.target.position({ x: box.x + box.w / 2, y: box.y + box.h / 2 })
        return
      }
      if (tablePinRef.current) {
        // the gesture was a cell-block selection; the table never moved
        tablePinRef.current = null
        e.target.position({ x: box.x + box.w / 2, y: box.y + box.h / 2 })
        return
      }
      altKeyCleanupRef.current?.()
      dragPosRef.current = null
      emitNodeGesture({ sourceId: node.sourceId, kind: 'end' })
      setAltDrag(false)
      onDragGuides([])
      if (multiDrag && onSelectionDragEnd) {
        void onSelectionDragEnd().finally(() =>
          emitNodeGesture({ sourceId: node.sourceId, kind: 'committed' }),
        )
        return
      }
      const dropX = e.target.x() - box.w / 2
      const dropY = e.target.y() - box.h / 2
      // Option+drag = duplicate at the drop point;
      // original snaps back (model unchanged, the Konva node position must be manually restored); multi-select/in-group gestures not supported yet
      if (e.evt?.altKey && onDuplicateTo && !multiDrag && !insideGroupId) {
        e.target.position({ x: box.x + box.w / 2, y: box.y + box.h / 2 })
        onDuplicateTo(node.sourceId, dropX - box.x, dropY - box.y)
        return
      }
      void Promise.resolve(
        onTransform(
          node.sourceId,
          {
            x: dropX,
            y: dropY,
            w: box.w,
            h: box.h,
            rotationDeg: box.rotationDeg,
          },
          undefined,
          insideGroupId,
        ),
      ).finally(() => emitNodeGesture({ sourceId: node.sourceId, kind: 'committed' }))
    },
    onTransformStart: () => {
      emitNodeGesture({ sourceId: node.sourceId, kind: 'resize' })
      gestureIdRef.current += 1
      transformingRef.current = true
      lastPreviewRef.current = 0
      gesturePosRef.current = null
      rawSizeRef.current = { w: boxRef.current.w, h: boxRef.current.h }
      previewBusyRef.current = false
      previewPendingRef.current = false
      repliedAskRef.current = null
      contentFloorRef.current = 0
    },
    onTransform: (e: Konva.KonvaEventObject<Event>) => {
      // glyphs must not be stretched by the handle drag (see text-unstretch.ts)
      if (node.type === 'shape' || node.type === 'text' || node.type === 'table')
        unstretchGlyphLayer(groupRef.current)
      const t = e.target
      const sx = Math.abs(t.scaleX())
      const sy = Math.abs(t.scaleY())
      // The pointer's size, before anything below draws the node differently. An axis the dragged
      // handle does not control keeps what it was (the node may be drawn taller than the pointer
      // asked, and a middle handle must not feed that drawn height back as a request).
      {
        const a = movingAnchor() ?? ''
        const prev = rawSizeRef.current ?? { w: boxRef.current.w, h: boxRef.current.h }
        rawSizeRef.current = {
          w: a.includes('left') || a.includes('right') ? boxRef.current.w * sx : prev.w,
          h: a.includes('top') || a.includes('bottom') ? boxRef.current.h * sy : prev.h,
        }
      }
      // A table never draws shorter than its content needs (see the rebase effect): the pointer may
      // ask for less than the re-wrapped rows take, the frame then follows the content
      if (
        node.type === 'table' &&
        !box.flipH &&
        !box.flipV &&
        contentFloorRef.current > boxRef.current.h * sy + 0.25
      ) {
        const drawnH = drawnTableHeight(boxRef.current.h * sy, contentFloorRef.current)
        const d = centerShift(
          movingAnchor(),
          t.rotation(),
          { w: boxRef.current.w * sx, h: boxRef.current.h * sy },
          { w: boxRef.current.w * sx, h: drawnH },
        )
        t.scaleY(drawnH / boxRef.current.h)
        t.position({ x: t.x() + d.dx, y: t.y() + d.dy })
      }
      // Track the live position first: the preview re-render below re-applies controlled
      // x/y, and the rebase effect restores this value to keep the gesture continuous
      gesturePosRef.current = { x: t.x(), y: t.y() }
      // Live preview during drag: re-lay out text at the new box size (font size unchanged).
      // Commit width/height only (x/y are managed by the Konva gesture and land in the model with the final commit on release).
      // Offset conversion for flipped elements relies on a one-shot back-calculation at gesture end; preview stays disabled for them.
      if (!livePreview || box.flipH || box.flipV) return
      if (Math.abs(sx - 1) < 1e-3 && Math.abs(sy - 1) < 1e-3) return // Pure rotation / not scaled
      sendPreview()
    },
    onTransformEnd: (e: Konva.KonvaEventObject<Event>) => {
      emitNodeGesture({ sourceId: node.sourceId, kind: 'end' })
      const gid = gestureIdRef.current
      const t = e.target
      const sx = Math.abs(t.scaleX())
      const sy = Math.abs(t.scaleY())
      // A table is committed at the drawn size so the engine sets stored row heights matching the content;
      // the frame sits where it was DRAWN, which can be taller than the raw pointer when content needed room
      const raw = node.type === 'table' ? rawSizeRef.current : null
      const drawnW = box.w * sx
      const drawnH = box.h * sy
      const tbl = node.type === 'table' ? (node as TableRenderNode) : null
      const tableRowHeightsPx =
        tbl && tbl.gridY && tbl.gridY.length > 1
          ? tbl.gridY.slice(1).map((y, i) => (y - tbl.gridY[i]!) * sy)
          : undefined
      const settleNow = () => {
        const held = settleRef.current
        if (!held || gestureIdRef.current !== gid) return
        settleRef.current = null
        window.clearTimeout(timer)
        transformingRef.current = false
        gesturePosRef.current = null
        pivotKeyRef.current = '' // the next render re-pins the node to whatever box the commit yields
        t.scaleX(t.scaleX() < 0 ? -1 : 1)
        t.scaleY(t.scaleY() < 0 ? -1 : 1)
        unstretchGlyphLayer(t as Konva.Group)
        const pv = boxPivotProps(modelBoxRef.current)
        t.position({ x: pv.x, y: pv.y })
        t.fire('heightChange')
      }
      let timer = 0
      const done = onTransform(
        node.sourceId,
        {
          // The offset point stays the drawn box's center at any scale, so the model's
          // unrotated top-left is always position − half of the NEW size
          x: t.x() - drawnW / 2,
          y: t.y() - drawnH / 2,
          w: raw ? raw.w : drawnW,
          h: node.type === 'table' ? drawnH : raw ? raw.h : drawnH,
          rotationDeg: t.rotation(),
          ...(tableRowHeightsPx ? { tableRowHeightsPx } : {}),
        },
        undefined,
        insideGroupId,
      )
      void Promise.resolve(done).finally(() =>
        emitNodeGesture({ sourceId: node.sourceId, kind: 'committed' }),
      )
      if (node.type === 'table' && !insideGroupId) {
        // keep the drawn size until the committed layout lands: resetting now would show the stale
        // model box for the length of the round trip
        const held = { key: boxKey(modelBoxRef.current), armed: false, run: settleNow }
        settleRef.current = held
        void Promise.resolve(done).finally(() => {
          // the commit's reply is in: the next render carries its box. When that box equals the held
          // one (a no-op commit) there is no such render, so the size lets go on its own.
          held.armed = true
          if (boxKey(modelBoxRef.current) !== held.key) settleNow()
          else timer = window.setTimeout(settleNow, 250)
        })
      } else {
        transformingRef.current = false
        gesturePosRef.current = null
        pivotKeyRef.current = ''
        t.scaleX(t.scaleX() < 0 ? -1 : 1)
        t.scaleY(t.scaleY() < 0 ? -1 : 1)
        if (node.type === 'shape' || node.type === 'text' || node.type === 'table')
          unstretchGlyphLayer(groupRef.current)
      }
    },
  }

  // Group in in-group editing mode: whole-group hit area off, children become interactive NodeViews (geometry commits carry groupId)
  if (node.type === 'group' && entered) {
    const g = node as GroupRenderNode
    // The DOM text overlay doesn't follow group rotation/flip; in those cases children can't double-click into text editing.
    // ext/chExt scaling is already baked into child geometry (including text layout), so it no longer affects overlay alignment.
    const plain = !box.rotationDeg && !box.flipH && !box.flipV
    return (
      <Group id={`node_${node.sourceId}`} {...boxPivotProps(box)}>
        {/* In-group boundary indicator (dashed frame while editing a group) */}
        <Rect
          width={box.w}
          height={box.h}
          stroke="#8e8e93"
          strokeWidth={1 / Math.max(zoom, 0.1)}
          dash={[3, 3]}
          listening={false}
        />
        <Group>
          {g.children.map((c) => (
            <NodeView
              key={c.id}
              node={c}
              onSelect={onSelect}
              onEditText={onEditText}
              onTransform={onTransform}
              onEditTableCell={onEditTableCell}
              onTableCellPress={onTableCellPress}
              onPlayMedia={onPlayMedia}
              onDragGuides={onDragGuides}
              snapTargets={snapTargets}
              images={images}
              editingText={editingText}
              livePreview={false}
              zoom={zoom}
              multiDrag={(selectedIds?.length ?? 0) > 1 && !!selectedIds?.includes(c.sourceId)}
              selectionDrag={selectionDrag}
              onSelectionDragPress={onSelectionDragPress}
              onSelectionDragEnd={onSelectionDragEnd}
              onSelectionDragMove={onSelectionDragMove}
              selectedIds={selectedIds}
              insideGroupId={node.sourceId}
              allowChildTextEdit={plain}
              suppressClickRef={suppressClickRef}
              selStroke={selStroke}
              selHairline={selHairline}
            />
          ))}
        </Group>
      </Group>
    )
  }

  // Double-click a group = enter in-group editing and select the child hit by the double-click (pointer converted to group-local coordinates, bounding-box hit)
  const onGroupDblClick = (e: Konva.KonvaEventObject<Event>) => {
    if (!onEnterGroup) return
    const g = node as GroupRenderNode
    const p = e.target.getStage()?.getPointerPosition()
    let childId: string | null = null
    if (p && groupRef.current) {
      // children box already includes ext/chExt scaling (group-local px); compare directly after converting the pointer
      const local = groupRef.current.getAbsoluteTransform().copy().invert().point(p)
      const lx = local.x
      const ly = local.y
      const hit = [...g.children]
        .reverse()
        .find(
          (c) =>
            lx >= c.box.x && lx <= c.box.x + c.box.w && ly >= c.box.y && ly <= c.box.y + c.box.h,
        )
      childId = hit?.sourceId ?? null
    }
    onEnterGroup(node.sourceId, childId)
  }
  // Table: double-click hits a cell in table-local coordinates, regardless of rotation/flip.
  const onTableDblClick = (e: Konva.KonvaEventObject<Event>) => {
    if (node.type !== 'table') return
    const pos = e.target.getStage()?.getPointerPosition()
    if (!pos) return
    const local = tableLocalPointFromStage(pos, box, CANVAS_BLEED)
    const cell = tableCellAtPoint(node, local)
    if (cell) onEditTableCell(node.sourceId, cell.row, cell.col)
  }
  // Audio/video (image is the poster frame): double-click opens the playback overlay
  const isMedia = node.type === 'picture' && !!(node as PictureRenderNode).media && !!onPlayMedia
  // Empty placeholder: canvas draws a gray click hint (edit canvas only, not in thumbnails/export)
  const phPrompt = (() => {
    if (hidePhPrompts) return null
    if (node.type !== 'shape' && node.type !== 'text') return null
    const sh = node as ShapeRenderNode
    if (!isPromptPlaceholder(sh)) return null
    if (editingText && editingText.sourceId === node.sourceId) return null
    const kind = sh.placeholder
    return t(
      kind === 'subTitle'
        ? 'appPhPromptSubtitle'
        : kind === 'body'
          ? 'appPhPromptBody'
          : 'appPhPromptTitle',
    )
  })()
  return (
    <>
      {/* Option+drag: the original stays rendered at the source position while the dragged ghost travels */}
      {altDragging && <StaticNode key="ghost-src" node={node} images={images} />}
      <Group
        key="node"
        ref={groupRef}
        opacity={altDragging ? 0.5 : 1}
        {...common}
        {...(editable && !(mathNode && !insideGroupId)
          ? {
              onDblClick: (e: Konva.KonvaEventObject<MouseEvent>) =>
                onEditText(node.sourceId, { x: e.evt.clientX, y: e.evt.clientY, select: 'word' }),
              onDblTap: () => onEditText(node.sourceId),
            }
          : mathNode && !insideGroupId
            ? { onDblTap: () => onEditText(node.sourceId) }
            : node.type === 'group' && !insideGroupId && onEnterGroup
              ? { onDblClick: onGroupDblClick, onDblTap: onGroupDblClick }
              : node.type === 'table' && !insideGroupId
                ? { onDblClick: onTableDblClick, onDblTap: onTableDblClick }
                : isMedia && !insideGroupId
                  ? {
                      onDblClick: () => onPlayMedia!(node.sourceId),
                      onDblTap: () => onPlayMedia!(node.sourceId),
                    }
                  : {})}
      >
        {/* The selection frame is grabbable across the whole box: transparent fill ensures
            unfilled shapes, text insets, and transparent images are directly clickable without clicking through */}
        {selected && draggable && !isConnectorNode(node) && (
          <Rect
            width={box.w}
            height={box.h}
            fill="transparent"
            stroke="transparent"
            strokeWidth={0}
            hitStrokeWidth={8 / Math.max(zoom, 0.1)}
          />
        )}
        {/* group children don't take hits (listening=false); add a transparent hit area so the whole group can be selected/dragged */}
        {node.type === 'group' && <Rect width={box.w} height={box.h} fill="transparent" />}
        {/* The equation raster follows this shape's transform. Keep the whole box hittable
            so its first click selects before the native two-click editor gesture. */}
        {mathNode && <Rect width={box.w} height={box.h} fill="transparent" />}
        <NodeBody
          node={node}
          images={images}
          hideText={!!editingText && editingText.sourceId === node.sourceId && !editingText.cell}
          hideCellText={
            editingText && editingText.sourceId === node.sourceId ? editingText.cell : undefined
          }
        />
        {/* Multi-select: PowerPoint-style per-element border so every selected element is visibly selected
          (the shared Transformer only draws one combined box). Lives inside the node group so it follows drags/transforms. */}
        {multiDrag && !isConnectorNode(node) && (
          <Rect
            width={box.w}
            height={box.h}
            stroke={selStroke}
            strokeWidth={selHairline ?? chromeHairline(zoom)}
            strokeScaleEnabled={false}
            listening={false}
          />
        )}
        {/* Option+drag ghost: dashed frame marks the travelling copy */}
        {altDragging && (
          <Rect
            width={box.w}
            height={box.h}
            stroke={selStroke}
            strokeWidth={selHairline ?? chromeHairline(zoom)}
            strokeScaleEnabled={false}
            dash={[4, 4]}
            listening={false}
          />
        )}
        {phPrompt &&
          (() => {
            const sh = node as ShapeRenderNode
            const ins = sh.text?.insets ?? { l: 8, t: 4, r: 8, b: 4 }
            return (
              <Text
                text={phPrompt}
                x={ins.l}
                y={ins.t}
                width={Math.max(box.w - ins.l - ins.r, 20)}
                height={Math.max(box.h - ins.t - ins.b, 16)}
                align={sh.placeholder === 'body' ? 'left' : 'center'}
                verticalAlign={sh.text?.anchor ?? (sh.placeholder === 'body' ? 'top' : 'middle')}
                fontSize={Math.min(28, Math.max(14, box.h * 0.22))}
                fill="#8e8e93"
              />
            )
          })()}
      </Group>
    </>
  )
}

// Vertical/horizontal alignment offsets are already baked into glyph coordinates by pptx-render's layoutText; the renderer draws what it gets.
// Fallback reference (the Transformer frame still needs ShapeRenderNode type semantics)
export type { ShapeRenderNode }
