import { describe, expect, it } from 'vitest'
import { HorizontalAlign, VerticalAlign } from '@univerjs/core'
import {
  computeEditorFit,
  editorOverlayShift,
  editorPageWidth,
  fullPadding,
  normalizeEditorLayoutObject,
  shiftEditorOverlay,
  EDITOR_CARET_ROOM,
} from '../src/renderer/cell-editor-overlay-fix'

describe('cell-editor-overlay-fix', () => {
  describe('fullPadding', () => {
    it('fills in missing sides from default padding', () => {
      const padding = fullPadding({ l: 10 })
      expect(padding.l).toBe(10)
      expect(padding.r).toBe(2)
      expect(padding.t).toBe(0)
      expect(padding.b).toBe(2)
    })

    it('keeps existing values when all sides provided', () => {
      const padding = fullPadding({ t: 1, b: 2, l: 3, r: 4 })
      expect(padding).toEqual({ t: 1, b: 2, l: 3, r: 4 })
    })

    it('uses defaults when padding is null or undefined', () => {
      expect(fullPadding(null)).toEqual({ t: 0, b: 2, l: 2, r: 2 })
      expect(fullPadding(undefined)).toEqual({ t: 0, b: 2, l: 2, r: 2 })
    })
  })

  describe('editorPageWidth', () => {
    it('uses cell width if text fits within cell width', () => {
      const cellWidth = 100
      const textWidth = 50
      const padding = { t: 0, b: 0, l: 2, r: 2 }
      const width = editorPageWidth(cellWidth, textWidth, padding, 1)
      expect(width).toBe(100)
    })

    it('widens editor page width when text exceeds cell width', () => {
      const cellWidth = 50
      const textWidth = 80
      const padding = { t: 0, b: 0, l: 4, r: 4 }
      const width = editorPageWidth(cellWidth, textWidth, padding, 1)
      // 80 + (4 + 4 + 0.05) * 1 = 88.05
      expect(width).toBeCloseTo(88.05, 2)
    })
  })

  describe('editorOverlayShift', () => {
    it('shifts half-pixel for non-middle vertical align', () => {
      const shift = editorOverlayShift(1, 1, VerticalAlign.TOP)
      expect(shift.dx).toBe(0.5)
      expect(shift.dy).toBe(0.5)
    })

    it('compensates for MIDDLE_CELL_POS_MAGIC_NUMBER (1 unit) when middle aligned', () => {
      const shift = editorOverlayShift(1, 1, VerticalAlign.MIDDLE)
      expect(shift.dx).toBe(0.5)
      // (0.5 - 1) * 1 = -0.5
      expect(shift.dy).toBe(-0.5)
    })

    it('scales with zoom (scaleX and scaleY)', () => {
      const shift = editorOverlayShift(2, 2, VerticalAlign.TOP)
      expect(shift.dx).toBe(1.0)
      expect(shift.dy).toBe(1.0)
    })
  })

  describe('shiftEditorOverlay', () => {
    it('shifts start and end coordinates', () => {
      const rect = { startX: 10, startY: 20, endX: 110, endY: 50, show: true }
      const shifted = shiftEditorOverlay(rect, { dx: 0.5, dy: -0.5 })
      expect(shifted.startX).toBe(10.5)
      expect(shifted.startY).toBe(19.5)
      expect(shifted.endX).toBe(110.5)
      expect(shifted.endY).toBe(49.5)
      expect(shifted.show).toBe(true)
    })
  })

  describe('computeEditorFit', () => {
    it('computes exact margins for left-aligned text', () => {
      const fit = computeEditorFit({
        cellWidth: 100,
        cellHeight: 25,
        textWidth: 40,
        textHeight: 15,
        padding: { t: 2, b: 2, l: 5, r: 5 },
        scaleX: 1,
        scaleY: 1,
        horizontalAlign: HorizontalAlign.LEFT,
        verticalAlign: VerticalAlign.TOP,
      })
      expect(fit.width).toBe(100)
      expect(fit.height).toBe(25)
      expect(fit.offsetLeft).toBe(5)
      expect(fit.offsetTop).toBe(2)
    })

    it('computes exact margins for right-aligned text without rounding errors', () => {
      const fit = computeEditorFit({
        cellWidth: 100,
        cellHeight: 25,
        textWidth: 50,
        textHeight: 15,
        padding: { t: 2, b: 2, l: 3, r: 3 },
        scaleX: 1,
        scaleY: 1,
        horizontalAlign: HorizontalAlign.RIGHT,
        verticalAlign: VerticalAlign.MIDDLE,
      })
      // actualWidth = 50 + 6 = 56
      // width = 100
      // offsetLeft = (100 - 56) = 44
      expect(fit.width).toBe(100)
      expect(fit.offsetLeft).toBe(44)
    })
  })

  describe('normalizeEditorLayoutObject', () => {
    it('sets zeroWidthParagraphBreak on document renderConfig', () => {
      const layout = {
        paddingData: { l: 4 },
        documentModel: {
          documentStyle: {
            renderConfig: {},
          },
        },
      }
      normalizeEditorLayoutObject(layout)
      expect(layout.paddingData).toEqual({ t: 0, b: 2, l: 4, r: 2 })
      expect(layout.documentModel.documentStyle.renderConfig).toEqual({
        zeroWidthParagraphBreak: 1,
      })
      expect(layout.documentModel.documentStyle.marginLeft).toBe(4)
      expect(layout.documentModel.documentStyle.marginRight).toBe(2)
    })
  })
})
