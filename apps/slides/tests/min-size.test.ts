import { describe, expect, it } from 'vitest'
import {
  MAX_ELEMENT_PX,
  anchorsForScreenSize,
  clampAnchorPos,
  MIN_CELL_COL_PX,
  MIN_CELL_ROW_PX,
  clampTransformBox,
  colDragRange,
  colWidthFloor,
  fitSizeToLimits,
  minSizeFor,
  minSizeForMembers,
  resolveColDrag,
  resolveRowDrag,
  rowDragFloor,
  sanitizePosPx,
  sanitizeSizePx,
} from '../src/renderer/min-size'

const box = (x: number, y: number, width: number, height: number, rotation = 0) => ({
  x,
  y,
  width,
  height,
  rotation,
})

describe('min size', () => {
  it('keeps a normal resize untouched', () => {
    const next = box(10, 10, 200, 100)
    expect(clampTransformBox(box(10, 10, 220, 120), next, minSizeFor({ type: 'shape' }))).toBe(next)
  })

  it('stops a right-edge drag at the minimum and keeps the left edge', () => {
    const out = clampTransformBox(
      box(50, 50, 100, 80),
      box(50, 50, 3, 80),
      minSizeFor({ type: 'text' }),
    )
    expect(out).toMatchObject({ x: 50, width: 64, height: 80 })
  })

  it('stops a left-edge drag at the minimum and keeps the right edge', () => {
    const out = clampTransformBox(
      box(50, 50, 100, 80),
      box(148, 50, 2, 80),
      minSizeFor({ type: 'text' }),
    )
    expect(out.width).toBe(64)
    expect(out.x + out.width).toBe(150)
  })

  it('refuses to flip an element dragged past the opposite edge', () => {
    const out = clampTransformBox(
      box(50, 50, 100, 80),
      box(50, 130, 100, -40),
      minSizeFor({ type: 'table', rows: 2, cols: 3 }),
    )
    expect(out.height).toBe(48)
    expect(out.y + out.height).toBe(130) // bottom edge stays
  })

  it('table minimum grows with its grid', () => {
    expect(minSizeFor({ type: 'table', rows: 5, cols: 4 })).toEqual({ w: 256, h: 120 })
  })

  it('lines keep a zero extent', () => {
    const next = box(0, 0, 120, 0)
    expect(
      clampTransformBox(box(0, 0, 100, 0), next, minSizeFor({ type: 'shape', isConnector: true })),
    ).toBe(next)
  })

  it('rotated boxes fall back to the previous box when too small', () => {
    const old = box(10, 10, 100, 60, 30)
    expect(clampTransformBox(old, box(10, 10, 4, 60, 30), minSizeFor({ type: 'shape' }))).toBe(old)
  })

  describe('anchor clamp', () => {
    const rect = { x: 100, y: 100, width: 200, height: 100, rotation: 0 }
    const min = { w: 40, h: 30 }

    it('stops a left handle at the minimum width, right edge fixed', () => {
      expect(clampAnchorPos('top-left', { x: 400, y: 80 }, rect, min)).toEqual({ x: 260, y: 80 })
    })

    it('stops a bottom-right handle dragged past the top-left corner', () => {
      expect(clampAnchorPos('bottom-right', { x: 10, y: 20 }, rect, min)).toEqual({
        x: 140,
        y: 130,
      })
    })

    it('leaves an in-range drag and edge handles on their free axis alone', () => {
      expect(clampAnchorPos('middle-right', { x: 350, y: 999 }, rect, min)).toEqual({
        x: 350,
        y: 999,
      })
    })

    it('works in the rotated frame', () => {
      const r = { ...rect, rotation: Math.PI / 2 }
      // rotated 90°: the box's local x axis points down the stage
      const out = clampAnchorPos('middle-right', { x: 100, y: 100 }, r, min)
      expect(out.x).toBeCloseTo(100, 6)
      expect(out.y).toBeCloseTo(140, 6)
    })
  })

  describe('table / group / multi floors', () => {
    it('a table with columns already thinner than the floor can still shrink to them', () => {
      const m = minSizeFor({
        type: 'table',
        rows: 2,
        cols: 3,
        colWidths: [30, 400, 80],
        rowHeights: [10, 100],
      })
      expect(m).toEqual({ w: 30 + 64 + 64, h: 10 + 24 })
    })

    describe('content-aware table width floor', () => {
      const base = {
        type: 'table',
        rows: 2,
        cols: 3,
        colWidths: [140, 260, 160],
        rowHeights: [50, 50],
      }
      it('stops at the column whose widest word reaches its width first (proportional scaling)', () => {
        // "Redmine Issue" column needs 112 of its 160 → the table may shrink to 112.5/160 of its width
        const m = minSizeFor({ ...base, colNeeds: [60, 70, 112] })
        expect(m.w).toBeCloseTo(560 * (112.5 / 160), 6)
        // no column ever ends narrower than the word it holds
        const k = m.w / 560
        ;[140, 260, 160].forEach((w, i) => expect(w * k).toBeGreaterThanOrEqual([60, 70, 112][i]!))
      })
      it('short words fall back to the tiny absolute floor, not 64px per column', () => {
        const m = minSizeFor({ ...base, colNeeds: [0, 0, 0] })
        expect(m.w).toBeCloseTo(560 * (24 / 140), 6)
        expect(m.w).toBeLessThan(3 * 64)
      })
      it('a column already narrower than its word does not lock the others, but it cannot shrink', () => {
        const m = minSizeFor({ ...base, colWidths: [30, 400, 130], colNeeds: [90, 50, 40] })
        // column 0 keeps its own 30px as floor → ratio 1 → the table cannot narrow at all
        expect(m.w).toBe(560)
      })
      it('the bottom-of-slide limit raises the width floor', () => {
        const free = minSizeFor({ ...base, colNeeds: [60, 70, 112] }).w
        expect(minSizeFor({ ...base, colNeeds: [60, 70, 112], fitMinW: 400 }).w).toBe(400)
        expect(minSizeFor({ ...base, colNeeds: [60, 70, 112], fitMinW: 10 }).w).toBe(free)
      })
      it('without measured needs the fixed per-column floor still applies', () => {
        expect(minSizeFor(base).w).toBe(3 * 64)
      })
    })

    it('a table cannot be dragged shorter than its wrapped text needs', () => {
      const m = minSizeFor({
        type: 'table',
        rows: 2,
        cols: 2,
        colWidths: [100, 100],
        rowHeights: [70, 75],
        rowNeeds: [40, 60],
      })
      expect(m.h).toBe(100)
      // a stale need above the row itself never locks the table beyond its current height
      expect(
        minSizeFor({
          type: 'table',
          rows: 1,
          cols: 1,
          colWidths: [100],
          rowHeights: [50],
          rowNeeds: [90],
        }).h,
      ).toBe(50)
    })

    it('a group cannot be squeezed past its tightest child', () => {
      const m = minSizeFor({
        type: 'group',
        w: 400,
        h: 300,
        children: [
          { type: 'shape', w: 200, h: 120 },
          { type: 'text', w: 100, h: 100 },
        ],
      })
      // text child: 64/100 of the width, shape child: 24/120 of the height → max(0.2, 0.24)
      expect(m.w).toBeCloseTo(256, 6)
      expect(m.h).toBeCloseTo(0.24 * 300, 6)
    })

    it('a child already under its floor may only halve, so a group is never locked', () => {
      const m = minSizeForMembers([{ type: 'shape', w: 10, h: 200 }], 100, 200)
      expect(m.w).toBe(50) // the 10px child may halve → the box may go to half its width, not lock
      expect(m.h).toBeCloseTo(24, 6)
    })

    it('connectors do not limit a multi-selection', () => {
      const m = minSizeForMembers(
        [
          { type: 'shape', isConnector: true, w: 300, h: 0 },
          { type: 'shape', w: 100, h: 100 },
        ],
        300,
        100,
      )
      expect(m.w).toBeCloseTo(72, 6) // 24/100 * 300
    })
  })

  describe('numeric size sanitising', () => {
    it('rejects non-finite values and clamps the rest', () => {
      expect(sanitizeSizePx(Number.NaN, 24)).toBeNull()
      expect(sanitizeSizePx(Infinity, 24)).toBeNull()
      expect(sanitizeSizePx(-5, 24)).toBe(24)
      expect(sanitizeSizePx(0, 24)).toBe(24)
      expect(sanitizeSizePx(1e9, 24)).toBe(MAX_ELEMENT_PX)
      expect(sanitizeSizePx(300, 24)).toBe(300)
    })

    it('positions stay finite and within the limit', () => {
      expect(sanitizePosPx(Number.NaN)).toBeNull()
      expect(sanitizePosPx(-1e9)).toBe(-MAX_ELEMENT_PX)
      expect(sanitizePosPx(12.5)).toBe(12.5)
    })
  })

  describe('tiny selections keep only corner handles', () => {
    it('shows all eight handles on a roomy box', () => {
      expect(anchorsForScreenSize(200, 100, 8)).toHaveLength(8)
    })

    it('drops the middle handle of a short edge', () => {
      const a = anchorsForScreenSize(200, 20, 8)
      expect(a).toContain('top-center')
      expect(a).not.toContain('middle-left')
      expect(a).not.toContain('middle-right')
    })

    it('leaves just the corners when both sides are tiny', () => {
      expect(anchorsForScreenSize(12, 12, 8)).toEqual([
        'top-left',
        'top-right',
        'bottom-left',
        'bottom-right',
      ])
    })
  })

  describe('typed size fit', () => {
    const min = { w: 24, h: 24 }
    it('clamps each side independently without a lock', () => {
      expect(fitSizeToLimits(2, 500, min)).toEqual({ w: 24, h: 500 })
      expect(fitSizeToLimits(99999, 10, min)).toEqual({ w: MAX_ELEMENT_PX, h: 24 })
    })
    it('keeps the proportions when locked', () => {
      const r = fitSizeToLimits(10, 5, min, { lockRatio: true })!
      expect(r.h).toBeCloseTo(24, 6)
      expect(r.w / r.h).toBeCloseTo(2, 6)
      const big = fitSizeToLimits(20000, 10000, min, { lockRatio: true })!
      expect(big.w).toBeCloseTo(MAX_ELEMENT_PX, 6)
      expect(big.w / big.h).toBeCloseTo(2, 6)
    })
    it('rejects NaN / Infinity', () => {
      expect(fitSizeToLimits(Number.NaN, 10, min)).toBeNull()
      expect(fitSizeToLimits(10, Infinity, min)).toBeNull()
    })
  })

  describe('table border drags', () => {
    const widths = [100, 200, 300]
    it('moves an inner border between its two columns, keeping the table width', () => {
      const r = resolveColDrag(widths, 0, 150)
      expect(r.wPx).toBe(150)
      expect(r.neighbor).toEqual({ col: 1, wPx: 150 })
    })
    it('stops when the neighbour would vanish', () => {
      const r = resolveColDrag(widths, 0, 5000)
      expect(r.wPx).toBe(300 - MIN_CELL_COL_PX)
      expect(r.neighbor!.wPx).toBe(MIN_CELL_COL_PX)
    })
    it('stops at the floor going the other way', () => {
      expect(resolveColDrag(widths, 1, -50).wPx).toBe(MIN_CELL_COL_PX)
    })
    it('the last column edge changes the table width instead', () => {
      const r = resolveColDrag(widths, 2, 500)
      expect(r).toEqual({ wPx: 500 })
      expect(resolveColDrag(widths, 2, 0).wPx).toBe(MIN_CELL_COL_PX)
    })
    it('a content floor keeps both neighbours wide enough for their words', () => {
      const needs = [60, 150, 40]
      // dragging border 0 right: column 1 (200) may shrink only to its 150.5 floor
      const r = resolveColDrag(widths, 0, 400, needs)
      expect(r.neighbor!.wPx).toBeCloseTo(150.5, 6)
      expect(r.wPx + r.neighbor!.wPx).toBeCloseTo(300, 6)
      // dragging it left: column 0 stops at its own 60.5 floor
      expect(resolveColDrag(widths, 0, 0, needs).wPx).toBeCloseTo(60.5, 6)
      // the last column's outer edge: its own floor
      expect(resolveColDrag(widths, 2, 0, [0, 0, 120]).wPx).toBeCloseTo(120.5, 6)
      expect(colDragRange(widths, 1, needs)).toEqual({ lo: 150.5, hi: 200 + 300 - 40.5 })
    })
    it('a column already thinner than its word is never forced wider by a drag', () => {
      const r = resolveColDrag([20, 200], 0, 400, [90, 0])
      expect(r.wPx).toBeLessThanOrEqual(200 + 20 - 32)
      expect(resolveColDrag([20, 200], 0, 0, [90, 0]).wPx).toBe(20)
    })
    it('a typed column width floor is the neediest selected column', () => {
      expect(colWidthFloor(widths, [0, 1], [90, 150, 0])).toBeCloseTo(150.5, 6)
    })
    it('ignores a non-finite drag', () => {
      expect(resolveColDrag(widths, 1, Number.NaN)).toEqual({ wPx: 200 })
      expect(resolveRowDrag(Number.NaN, 40)).toBe(40)
    })
    it('row heights clamp to the floor and the limit', () => {
      expect(resolveRowDrag(2, 40)).toBe(MIN_CELL_ROW_PX)
      expect(resolveRowDrag(1e9, 40)).toBe(MAX_ELEMENT_PX)
    })
    it('a row stops where its wrapped text needs the height', () => {
      expect(resolveRowDrag(2, 90, 73.6)).toBe(73.6)
      expect(resolveRowDrag(80, 90, 73.6)).toBe(80)
      expect(rowDragFloor(90, 73.6)).toBe(73.6)
    })
    it('a row is never forced taller than it already is by its need', () => {
      expect(rowDragFloor(30, 73.6)).toBe(30)
      expect(resolveRowDrag(2, 30, 73.6)).toBe(30)
      expect(rowDragFloor(5, 0)).toBe(MIN_CELL_ROW_PX)
    })
  })
})
