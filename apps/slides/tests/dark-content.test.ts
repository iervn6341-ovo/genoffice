import { describe, it, expect } from 'vitest'
import { readFileSync, readdirSync, statSync } from 'node:fs'
import { join } from 'node:path'
import type {
  ChartRenderNode,
  GroupRenderNode,
  PictureRenderNode,
  RenderSlide,
  ShapeRenderNode,
  TableRenderNode,
} from '@genoffice/pptx-render'
import {
  DARK_PAPER_HEX,
  darkContentColor,
  displaySlide,
  mapSlideColors,
} from '../src/renderer/dark-content'

const run = (color: string) => ({
  text: 'x',
  x: 0,
  baselineY: 10,
  fontFamily: 'Arial',
  fontSizePx: 12,
  color,
  bold: false,
  italic: false,
  underline: false,
  widthPx: 8,
})
const box = { x: 0, y: 0, w: 10, h: 10 } as never

function shape(id: string): ShapeRenderNode {
  return {
    id,
    type: 'shape',
    box,
    sourceId: id,
    fill: { kind: 'solid', color: '#ffffff' },
    stroke: { color: '#000000', widthPx: 1, widthPt: 1 },
    text: {
      lines: [{ runs: [run('#000000')], top: 0, height: 12 }],
      insets: { l: 0, t: 0, r: 0, b: 0 },
      anchor: 'top',
      fontScale: 1,
      contentHeight: 12,
      wrap: true,
    },
  } as ShapeRenderNode
}

function sampleSlide(): RenderSlide {
  const table = {
    id: 't',
    type: 'table',
    box,
    sourceId: 't',
    gridX: [0, 10],
    gridY: [0, 10],
    cells: [
      {
        x: 0,
        y: 0,
        w: 10,
        h: 10,
        row: 0,
        col: 0,
        fill: { kind: 'solid', color: '#ffffff' },
        borders: { l: { color: '#000000', widthPx: 1, widthPt: 1 } },
      },
    ],
  } as unknown as TableRenderNode
  const chart = {
    id: 'c',
    type: 'chart',
    box,
    sourceId: 'c',
    gridLines: [{ x1: 0, y1: 0, x2: 1, y2: 1, color: '#d9d9d9' }],
    axisLines: [],
    labels: [{ text: 'a', x: 0, y: 0, fontSizePx: 10, color: '#000000' }],
    bars: [{ x: 0, y: 0, w: 1, h: 1, color: '#4472c4' }],
    polylines: [],
    markers: [],
    swatches: [],
  } as unknown as ChartRenderNode
  const pic = {
    id: 'p',
    type: 'picture',
    box,
    sourceId: 'p',
    dataUrl: 'data:image/png;base64,AAAA',
  } as PictureRenderNode
  const group = {
    id: 'g',
    type: 'group',
    box,
    sourceId: 'g',
    children: [shape('s2')],
  } as GroupRenderNode
  return {
    widthPx: 100,
    heightPx: 50,
    scale: 1,
    background: { kind: 'solid', color: '#ffffff' },
    nodes: [shape('s1'), pic, table, chart, group],
  }
}

describe('darkContentColor', () => {
  it('maps authored white onto the dark paper tone and black to white', () => {
    expect(darkContentColor('#ffffff')).toBe(DARK_PAPER_HEX)
    expect(darkContentColor('#fff')).toBe(DARK_PAPER_HEX)
    expect(darkContentColor('#000000')).toBe('#ffffff')
  })
  it('keeps alpha and leaves keywords alone', () => {
    expect(darkContentColor('#ffffff80')).toBe(`${DARK_PAPER_HEX}80`)
    expect(darkContentColor('rgba(0, 0, 0, 0.5)')).toBe('rgb(255 255 255 / 0.5)')
    expect(darkContentColor('transparent')).toBe('transparent')
  })
  it('preserves hue direction for saturated colors', () => {
    const out = darkContentColor('#ff0000')
    const n = parseInt(out.slice(1), 16)
    expect(n >> 16).toBeGreaterThan((n >> 8) & 255)
  })
})

describe('mapSlideColors', () => {
  it('remaps vector colors, keeps ids/boxes/structure and picture bitmaps', () => {
    const src = sampleSlide()
    const out = mapSlideColors(src)
    expect(out.background).toEqual({ kind: 'solid', color: DARK_PAPER_HEX })
    expect(out.nodes.map((n) => n.id)).toEqual(src.nodes.map((n) => n.id))
    const s = out.nodes[0] as ShapeRenderNode
    expect(s.fill).toEqual({ kind: 'solid', color: DARK_PAPER_HEX })
    expect(s.stroke!.color).toBe('#ffffff')
    expect(s.text!.lines[0]!.runs[0]!.color).toBe('#ffffff')
    expect((out.nodes[1] as PictureRenderNode).dataUrl).toBe('data:image/png;base64,AAAA')
    const t = out.nodes[2] as TableRenderNode
    expect(t.cells[0]!.fill).toEqual({ kind: 'solid', color: DARK_PAPER_HEX })
    expect(t.cells[0]!.borders!.l!.color).toBe('#ffffff')
    const ch = out.nodes[3] as ChartRenderNode
    expect(ch.labels[0]!.color).toBe('#ffffff')
    expect(ch.bars[0]!.color).not.toBe('#4472c4')
    const g = out.nodes[4] as GroupRenderNode
    expect((g.children[0] as ShapeRenderNode).fill).toEqual({
      kind: 'solid',
      color: DARK_PAPER_HEX,
    })
  })

  it('never mutates the authored slide (model/export input stays original)', () => {
    const src = sampleSlide()
    const snapshot = JSON.stringify(src)
    mapSlideColors(src)
    expect(JSON.stringify(src)).toBe(snapshot)
  })

  it('an unset background paints the dark paper', () => {
    const src = { ...sampleSlide(), background: { kind: 'none' as const } }
    expect(mapSlideColors(src).background).toEqual({ kind: 'solid', color: DARK_PAPER_HEX })
  })

  it('is identity-stable per source slide (memoized thumbnails keep skipping renders)', () => {
    const src = sampleSlide()
    expect(mapSlideColors(src)).toBe(mapSlideColors(src))
  })
})

describe('displaySlide toggle', () => {
  it('light returns the authored slide untouched; dark returns the remapped copy', () => {
    const src = sampleSlide()
    expect(displaySlide(src, false)).toBe(src)
    expect(displaySlide(src, true)).not.toBe(src)
    expect(displaySlide(undefined, true)).toBeUndefined()
  })
})

describe('export / print / show paths stay authored', () => {
  const dir = join(__dirname, '../src/renderer')
  const files: string[] = []
  const walk = (d: string) => {
    for (const f of readdirSync(d)) {
      const p = join(d, f)
      if (statSync(p).isDirectory()) walk(p)
      else if (/\.(ts|tsx)$/.test(f)) files.push(p)
    }
  }
  walk(dir)
  const importers = files
    .filter((f) =>
      /from '(\.\/|\.\.\/|\.\/\.\.\/)*(\.\/)?dark-content'/.test(readFileSync(f, 'utf8')),
    )
    .map((f) => f.slice(dir.length + 1))
    .sort()

  it('only the on-screen editing surfaces import dark-content', () => {
    expect(importers).toEqual(['App.tsx', 'TextEditOverlay.tsx'])
  })

  it('App applies the preview only to canvas + thumbnails, not export/show/print call sites', () => {
    const app = readFileSync(join(dir, 'App.tsx'), 'utf8')
    expect(app).not.toMatch(/renderSlidesToPngBase64\([^)]*(canvasSlide|displaySlide)/)
    expect(app).not.toMatch(/<SlideShowView[^>]*displaySlide/)
    const show = app.slice(app.indexOf('<SlideShowView'))
    expect(show.slice(0, 600)).not.toMatch(/darkContent|canvasSlide/)
  })
})
