/**
 * Rendering layer for animation playback — shared by the show view and the edit-area "Preview".
 *
 * - useAnimPlayer: holds the playback cursor (steps played + current step's rAF progress) and
 *   outputs each node's NodeAnimState; advance() moves one step (if playing, completes the
 *   current step immediately).
 * - AnimatedSlideStage: static whole-page rendering same as SlideThumb, but each node is wrapped
 *   in a Konva Group driven by NodeAnimState (opacity/scale/rotation/offset/wipe clip,
 *   transformed around the node center).
 */
import React, { useEffect, useLayoutEffect, useRef } from 'react'
import { useAnimPlayer } from '../use-anim-player'
export { useAnimPlayer, type AnimPlayer } from '../use-anim-player'
import { Stage, Layer, Rect, Group } from 'react-konva'
import type { Context } from 'konva/lib/Context'
import type { Stage as KonvaStage } from 'konva/lib/Stage'
import type { RenderNode, RenderSlide, ShapeRenderNode } from '@genoffice/pptx-render'
import type { AnimationItem } from '../../shared/ipc'
import { parseParaStateKey, type NodeAnimState } from '../animation-play'
import { useI18n } from '../i18n/locale'
import { fillToKonva } from '../konva-adapter'
import { StaticNode } from '../NodeBody'

/** Whole-page rendering (with animation state); with an empty states Map it's equivalent to SlideThumb. */
export function AnimatedSlideStage({
  slide,
  images,
  width,
  states,
}: {
  slide: RenderSlide
  images: Map<string, HTMLImageElement>
  width: number
  states: Map<string, NodeAnimState>
}) {
  const stageRef = useRef<KonvaStage>(null)
  // Konva normally batches draws to the next frame. Flush a page replacement
  // before browser paint so the old canvas cannot flash with the new page.
  useLayoutEffect(() => {
    stageRef.current?.draw()
  }, [slide])
  const scale = width / slide.widthPx
  const h = slide.heightPx * scale
  const bg = fillToKonva(slide.background, slide.widthPx, slide.heightPx, images)
  // Paragraph animation states grouped by sourceId (key format: see animStateKey)
  const paraStates = new Map<string, Array<{ para: number; st: NodeAnimState }>>()
  for (const [key, st] of states) {
    const pk = parseParaStateKey(key)
    if (!pk) continue
    const arr = paraStates.get(pk.sourceId) ?? []
    arr.push({ para: pk.para, st })
    paraStates.set(pk.sourceId, arr)
  }
  return (
    <Stage
      ref={stageRef}
      width={width}
      height={h}
      listening={false}
      style={{ pointerEvents: 'none' }}
    >
      <Layer scaleX={scale} scaleY={scale} listening={false}>
        <Rect
          x={0}
          y={0}
          width={slide.widthPx}
          height={slide.heightPx}
          {...(Object.keys(bg).length ? bg : { fill: '#ffffff' })}
        />
        {slide.nodes.map((n) => (
          <AnimNode
            key={n.id}
            node={n}
            images={images}
            st={n.decoration ? undefined : states.get(n.sourceId)}
            paraSts={n.decoration ? undefined : paraStates.get(n.sourceId)}
          />
        ))}
      </Layer>
    </Stage>
  )
}

/**
 * Animation preview overlay on the edit canvas (ribbon "Preview" / animation pane ▶):
 * covers .stage-rel and plays all steps at 1:1 size (350ms between steps),
 * exiting when finished or on click.
 */
export function AnimPreviewOverlay({
  slide,
  images,
  items,
  onDone,
}: {
  slide: RenderSlide
  images: Map<string, HTMLImageElement>
  items: AnimationItem[]
  onDone: () => void
}) {
  const { t } = useI18n()
  const player = useAnimPlayer(slide.heightPx, slide.widthPx, { key: 'preview', items })
  const { playing, pending, played, epoch, advance } = player

  // Auto-advance: 350ms pause between steps; after everything plays, exit after 600ms
  useEffect(() => {
    if (playing) return
    const t = window.setTimeout(
      () => {
        if (!advance()) onDone()
      },
      pending ? 350 : 600,
    )
    return () => window.clearTimeout(t)
  }, [
    playing,
    pending,
    // Instant effects change the cursor without changing playing/pending.
    // A metadata reload also resets the pause even if that cursor is unchanged.
    played,
    epoch,
    advance,
    onDone,
  ])

  return (
    <div className="anim-preview-overlay" onClick={onDone} data-tip={t('paneAnimClickToEnd')}>
      <AnimatedSlideStage
        slide={slide}
        images={images}
        width={slide.widthPx}
        states={player.states}
      />
    </div>
  )
}

/** Transform Group props driven by st (scale/rotation around (cx,cy) + offset/opacity). */
function stateGroupProps(st: NodeAnimState, cx: number, cy: number) {
  return {
    x: cx + st.dx,
    y: cy + st.dy,
    offsetX: cx,
    offsetY: cy,
    scaleX: st.scale * st.scaleX,
    scaleY: st.scale,
    rotation: st.rotationDeg,
    opacity: st.opacity,
    listening: false as const,
  }
}

/** Wipe/split clip: narrows the visible area within rect according to t/mode. */
function wipedRect(
  clip: NonNullable<NodeAnimState['clip']>,
  rect: { x?: number; y: number; w?: number; h: number },
): { x: number; y: number; w: number; h: number } {
  const x = rect.x ?? 0
  const y = rect.y
  const w = rect.w ?? 0
  const h = rect.h
  const visH = h * clip.t
  const visW = w * clip.t
  if (clip.mode === 'top') return { x, y, w, h: visH }
  if (clip.mode === 'mid') return { x, y: y + (h - visH) / 2, w, h: visH }
  if (clip.mode === 'left') return { x, y, w: visW, h }
  if (clip.mode === 'right') return { x: x + (w - visW), y, w: visW, h }
  return { x, y: y + (h - visH), w, h: visH }
}

/**
 * Vertical line ranges per text paragraph (slide coordinates): lines grouped by paraStart,
 * missing means a new paragraph (same paragraph-split rule as the editor). Vertical/rotated/flipped
 * text is approximated as untransformed horizontal — a trade-off on par with the whole-box wipe clip.
 */
function paragraphSpans(node: RenderNode): Array<{ y0: number; y1: number }> {
  const text = (node as ShapeRenderNode).text
  if (!text) return []
  const base = node.box.y + text.insets.t
  const spans: Array<{ y0: number; y1: number }> = []
  for (const ln of text.lines) {
    const y0 = base + ln.top
    const y1 = y0 + ln.height
    if (ln.paraStart === false && spans.length) {
      spans[spans.length - 1]!.y1 = Math.max(spans[spans.length - 1]!.y1, y1)
    } else {
      spans.push({ y0, y1 })
    }
  }
  return spans
}

function AnimNode({
  node,
  images,
  st,
  paraSts,
}: {
  node: RenderNode
  images: Map<string, HTMLImageElement>
  st?: NodeAnimState
  /** Per-paragraph animation states for this node (only paragraphs covered by paragraph animations are listed) */
  paraSts?: Array<{ para: number; st: NodeAnimState }>
}) {
  const b = node.box
  const spans = paraSts?.length ? paragraphSpans(node) : []
  // Paragraph indexes may go out of range after text edits: paragraph animations without a matching line range are ignored whole-box
  const active = (paraSts ?? []).filter((p) => spans[p.para]).sort((x, y) => x.para - y.para)
  const cx = b.x + b.w / 2
  const cy = b.y + b.h / 2
  const m = Math.max(20, b.h * 0.1) // Margin on all sides to accommodate strokes/shadows

  if (!active.length) {
    if (!st) return <StaticNode node={node} images={images} />
    if (st.hidden) return null
    const clip = st.clip
    const clipFunc =
      clip != null
        ? (ctx: Context) => {
            // Wipe/split: clip the visible area to the node's bounding box
            const r = wipedRect(clip, { x: b.x, y: b.y, w: b.w, h: b.h })
            // The fixed edges are widened outward by m so edge strokes aren't clipped
            const top = clip.mode === 'top' ? m : 0
            const btm = clip.mode === 'btm' ? m : 0
            const left = clip.mode === 'left' ? m : 0
            const right = clip.mode === 'right' ? m : 0
            ctx.rect(r.x - left, r.y - top, r.w + left + right, r.h + top + btm)
          }
        : undefined
    return (
      <Group {...stateGroupProps(st, cx, cy)} clipFunc={clipFunc}>
        <StaticNode node={node} images={images} />
      </Group>
    )
  }

  // Paragraph layering: the base layer cuts out animated paragraphs' line ranges; each animated
  // paragraph renders a clone layer showing only its range with its own state. Cloning the whole node then
  // clipping — shape fill/stroke moves along with the paragraph strip; placeholder text boxes usually have no fill, acceptable approximation.
  if (st?.hidden) return null
  return (
    <Group {...(st ? stateGroupProps(st, cx, cy) : { listening: false as const })}>
      <Group
        listening={false}
        clipFunc={(ctx: Context) => {
          let y = b.y - m
          for (const p of active) {
            const s = spans[p.para]!
            if (s.y0 > y) ctx.rect(b.x - m, y, b.w + 2 * m, s.y0 - y)
            y = Math.max(y, s.y1)
          }
          const end = b.y + b.h + m
          if (end > y) ctx.rect(b.x - m, y, b.w + 2 * m, end - y)
        }}
      >
        <StaticNode node={node} images={images} />
      </Group>
      {active
        .filter((p) => !p.st.hidden) // Entrance paragraphs are invisible before playing (the base layer already cut them out)
        .map((p) => {
          const s = spans[p.para]!
          const cyP = (s.y0 + s.y1) / 2
          return (
            <Group
              key={p.para}
              {...stateGroupProps(p.st, cx, cyP)}
              clipFunc={(ctx: Context) => {
                // Paragraph line range; narrowed within the range during wipe. No top/bottom margin, avoiding bleeding into adjacent paragraphs
                const r = p.st.clip
                  ? wipedRect(p.st.clip, { x: b.x, y: s.y0, w: b.w, h: s.y1 - s.y0 })
                  : { x: b.x, y: s.y0, w: b.w, h: s.y1 - s.y0 }
                ctx.rect(b.x - m, r.y, b.w + 2 * m, r.h)
              }}
            >
              <StaticNode node={node} images={images} />
            </Group>
          )
        })}
    </Group>
  )
}
