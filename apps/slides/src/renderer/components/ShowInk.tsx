/**
 * Show ink — PowerPoint's pointer options during a show: laser pointer, pen, highlighter and
 * eraser. Coordinates are normalized 0..1 relative to the slide frame; the presenter and the
 * audience window each hold an InkBoard and restore them at their own frame size.
 *
 * Ink never goes through React state: a laser or pen stroke updates many times a frame, and
 * re-rendering the show view per pointer move (slide stage, thumbnails, notes) is what made the
 * mirrored laser stutter. The board redraws its canvas and moves the laser dot directly.
 */
import React, { useCallback, useEffect, useRef } from 'react'
import type { ShowInkEvent } from '../../shared/ipc'

/** Pointer mode: 'none' = the arrow (clicks advance the show) */
export type InkTool = 'none' | 'laser' | 'pen' | 'highlighter' | 'eraser'

export interface InkStroke {
  color: string
  kind: 'pen' | 'highlighter'
  /** Normalized coordinates alternating x,y */
  points: number[]
}

/**
 * PowerPoint's ink palette (Pointer Options), in its order. Ink is drawn onto the slide, so the
 * colours are content and identical in both UI themes.
 */
export const INK_COLORS = [
  { key: 'white', hex: '#FFFFFF' },
  { key: 'black', hex: '#000000' },
  { key: 'darkRed', hex: '#B3261E' },
  { key: 'red', hex: '#EA3323' },
  { key: 'orange', hex: '#F6B844' },
  { key: 'yellow', hex: '#FFFD54' },
  { key: 'lightGreen', hex: '#A1D16B' },
  { key: 'darkGreen', hex: '#4DA95A' },
  { key: 'lightBlue', hex: '#50B1EE' },
  { key: 'mediumBlue', hex: '#3470C3' },
  { key: 'darkBlue', hex: '#0E2366' },
  { key: 'purple', hex: '#8B2BE2' },
] as const

/** PowerPoint's defaults: a red pen, a yellow highlighter */
export const DEFAULT_PEN_COLOR = '#EA3323'
export const DEFAULT_HIGHLIGHTER_COLOR = '#FFFD54'

/** Eraser reach, as a share of the frame width */
const ERASER_RADIUS = 0.012

export class InkBoard {
  strokes: InkStroke[] = []
  private canvas: HTMLCanvasElement | null = null
  private laserEl: HTMLElement | null = null
  private size = { w: 0, h: 0 }
  private frame = 0
  private listeners = new Set<() => void>()

  attach(canvas: HTMLCanvasElement | null, laser: HTMLElement | null, w: number, h: number): void {
    this.canvas = canvas
    this.laserEl = laser
    this.size = { w, h }
    this.redraw()
  }

  /** Notified after any change to the strokes (e.g. to enable "Erase All Ink") */
  subscribe(fn: () => void): () => void {
    this.listeners.add(fn)
    return () => {
      this.listeners.delete(fn)
    }
  }

  /** Laser dot at a normalized point; null hides it */
  laser(p: { x: number; y: number } | null): void {
    const el = this.laserEl
    if (!el) return
    if (!p) {
      el.style.display = 'none'
      return
    }
    el.style.display = 'block'
    el.style.transform = `translate(${p.x * this.size.w}px, ${p.y * this.size.h}px)`
  }

  start(x: number, y: number, color: string, kind: InkStroke['kind']): void {
    this.strokes.push({ color, kind, points: [x, y] })
    this.changed()
  }

  extend(points: number[]): void {
    const last = this.strokes[this.strokes.length - 1]
    if (!last || !points.length) return
    last.points.push(...points)
    this.changed()
  }

  /** Remove every stroke passing within the eraser's reach of the point (PowerPoint erases whole strokes) */
  erase(x: number, y: number): void {
    const r = ERASER_RADIUS
    const ar = this.size.w && this.size.h ? this.size.h / this.size.w : 9 / 16
    const before = this.strokes.length
    this.strokes = this.strokes.filter((s) => {
      for (let i = 0; i < s.points.length; i += 2) {
        const dx = s.points[i]! - x
        const dy = (s.points[i + 1]! - y) * ar
        if (dx * dx + dy * dy <= r * r) return false
      }
      return true
    })
    if (this.strokes.length !== before) this.changed()
  }

  clear(): void {
    this.strokes = []
    this.laser(null)
    this.changed()
  }

  private changed(): void {
    // one repaint per display frame however many points arrive
    if (!this.frame) {
      this.frame = requestAnimationFrame(() => {
        this.frame = 0
        this.redraw()
      })
    }
    this.listeners.forEach((fn) => fn())
  }

  private redraw(): void {
    const canvas = this.canvas
    const { w, h } = this.size
    if (!canvas || !w || !h) return
    const dpr = window.devicePixelRatio || 1
    const pw = Math.round(w * dpr)
    const ph = Math.round(h * dpr)
    if (canvas.width !== pw || canvas.height !== ph) {
      canvas.width = pw
      canvas.height = ph
    }
    const ctx = canvas.getContext('2d')
    if (!ctx) return
    ctx.setTransform(dpr, 0, 0, dpr, 0, 0)
    ctx.clearRect(0, 0, w, h)
    ctx.lineCap = 'round'
    ctx.lineJoin = 'round'
    const pen = Math.max(2, w * 0.003)
    for (const s of this.strokes) {
      const pts = s.points
      if (pts.length < 2) continue
      ctx.globalAlpha = s.kind === 'highlighter' ? 0.45 : 1
      ctx.lineWidth = s.kind === 'highlighter' ? pen * 6 : pen
      ctx.strokeStyle = s.color
      ctx.beginPath()
      ctx.moveTo(pts[0]! * w, pts[1]! * h)
      // a single click still leaves a dot
      if (pts.length === 2) ctx.lineTo(pts[0]! * w + 0.01, pts[1]! * h)
      for (let i = 2; i < pts.length; i += 2) ctx.lineTo(pts[i]! * w, pts[i + 1]! * h)
      ctx.stroke()
    }
    ctx.globalAlpha = 1
  }
}

export function InkLayer({
  board,
  width,
  height,
}: {
  board: InkBoard
  width: number
  height: number
}): React.JSX.Element {
  const canvasRef = useRef<HTMLCanvasElement>(null)
  const laserRef = useRef<HTMLDivElement>(null)
  useEffect(() => {
    board.attach(canvasRef.current, laserRef.current, width, height)
    return () => board.attach(null, null, width, height)
  }, [board, width, height])
  return (
    <div className="ink-layer" style={{ width, height }}>
      <canvas ref={canvasRef} style={{ width, height }} />
      <div ref={laserRef} className="ink-laser" style={{ display: 'none' }} />
    </div>
  )
}

/**
 * Pointer handling for a show stage with a pointer tool active: draws on the board at once and
 * mirrors to the audience screen (`send`) at most once per display frame, so the laser moves
 * as smoothly there as here without flooding IPC.
 */
export function useInkPointer(
  board: InkBoard,
  tool: InkTool,
  color: string,
  stageRef: React.RefObject<HTMLElement | null>,
  send?: (ev: ShowInkEvent) => void,
) {
  const drawing = useRef(false)
  const pending = useRef<number[]>([])
  const laserAt = useRef<{ x: number; y: number } | null>(null)
  const frame = useRef(0)
  const sendRef = useRef(send)
  sendRef.current = send

  const flush = useCallback(() => {
    if (frame.current) cancelAnimationFrame(frame.current)
    frame.current = 0
    if (pending.current.length) {
      sendRef.current?.({ type: 'stroke-move', points: pending.current })
      pending.current = []
    }
    if (laserAt.current) {
      sendRef.current?.({ type: 'laser', ...laserAt.current })
      laserAt.current = null
    }
  }, [])
  const schedule = useCallback(() => {
    if (!frame.current) frame.current = requestAnimationFrame(flush)
  }, [flush])

  // Leaving the laser hides the dot on both screens
  useEffect(() => {
    if (tool === 'laser') return
    board.laser(null)
    laserAt.current = null
    sendRef.current?.({ type: 'laser', x: -1, y: -1 })
  }, [tool, board])
  useEffect(() => () => cancelAnimationFrame(frame.current), [])

  const norm = (e: React.PointerEvent): { x: number; y: number } | null => {
    const r = stageRef.current?.getBoundingClientRect()
    if (!r || !r.width || !r.height) return null
    return {
      x: Math.min(1, Math.max(0, (e.clientX - r.left) / r.width)),
      y: Math.min(1, Math.max(0, (e.clientY - r.top) / r.height)),
    }
  }

  return {
    onPointerDown: (e: React.PointerEvent) => {
      if (tool === 'none' || tool === 'laser' || e.button !== 0) return
      const p = norm(e)
      if (!p) return
      e.currentTarget.setPointerCapture(e.pointerId)
      drawing.current = true
      if (tool === 'eraser') {
        board.erase(p.x, p.y)
        sendRef.current?.({ type: 'erase', ...p })
        return
      }
      flush() // the previous stroke's last points must not land on this one
      board.start(p.x, p.y, color, tool)
      sendRef.current?.({ type: 'stroke-start', ...p, color, kind: tool })
    },
    onPointerMove: (e: React.PointerEvent) => {
      if (tool === 'laser') {
        const p = norm(e)
        if (!p) return
        board.laser(p)
        laserAt.current = p
        schedule()
        return
      }
      if (!drawing.current) return
      const p = norm(e)
      if (!p) return
      if (tool === 'eraser') {
        board.erase(p.x, p.y)
        sendRef.current?.({ type: 'erase', ...p })
        return
      }
      board.extend([p.x, p.y])
      pending.current.push(p.x, p.y)
      schedule()
    },
    onPointerUp: () => {
      drawing.current = false
      flush()
    },
    onPointerLeave: () => {
      if (tool !== 'laser') return
      board.laser(null)
      laserAt.current = null
      sendRef.current?.({ type: 'laser', x: -1, y: -1 })
    },
  }
}

/** Apply an ink event from the presenter to this screen's board */
export function applyInkEvent(board: InkBoard, ev: ShowInkEvent): void {
  if (ev.type === 'clear') board.clear()
  else if (ev.type === 'laser') board.laser(ev.x < 0 ? null : ev)
  else if (ev.type === 'stroke-start') board.start(ev.x, ev.y, ev.color, ev.kind)
  else if (ev.type === 'stroke-move') board.extend(ev.points)
  else if (ev.type === 'erase') board.erase(ev.x, ev.y)
}
