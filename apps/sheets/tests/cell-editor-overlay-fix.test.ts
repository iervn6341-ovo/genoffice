import { describe, expect, it } from 'vitest'
import { HorizontalAlign, VerticalAlign } from '@univerjs/core'
import {
  computeEditorFit,
  computeSkeletonVerticalOffset,
  editorOverlayShift,
  editorPageWidth,
  fullPadding,
  normalizeEditorLayoutObject,
  normalizeEditorSkeletonGlyphs,
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
    it('anchors container vertically flush with row line (dy = 0) to prevent top protrusion', () => {
      const topShift = editorOverlayShift(1, 1, VerticalAlign.TOP)
      expect(topShift.dx).toBe(0.5)
      expect(topShift.dy).toBe(0)

      const middleShift = editorOverlayShift(1, 1, VerticalAlign.MIDDLE)
      expect(middleShift.dx).toBe(0.5)
      expect(middleShift.dy).toBe(0)
    })

    it('scales dx with zoom (scaleX)', () => {
      const shift = editorOverlayShift(2, 2, VerticalAlign.TOP)
      expect(shift.dx).toBe(1.0)
      expect(shift.dy).toBe(0)
    })
  })

  describe('computeSkeletonVerticalOffset', () => {
    it('returns half-pixel blur offset for non-middle vertical align', () => {
      expect(computeSkeletonVerticalOffset(VerticalAlign.TOP)).toBe(0.5)
    })

    it('compensates for asymmetric vertical padding when middle aligned', () => {
      // Default padding (t: 0, b: 2): (0 - 2) / 2 + 0.5 = -0.5
      expect(computeSkeletonVerticalOffset(VerticalAlign.MIDDLE)).toBe(-0.5)

      // Symmetric padding (t: 2, b: 2): (2 - 2) / 2 + 0.5 = 0.5
      expect(computeSkeletonVerticalOffset(VerticalAlign.MIDDLE, { t: 2, b: 2 })).toBe(0.5)

      // Custom asymmetric padding (t: 1, b: 3): (1 - 3) / 2 + 0.5 = -0.5
      expect(computeSkeletonVerticalOffset(VerticalAlign.MIDDLE, { t: 1, b: 3 })).toBe(-0.5)
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
      // offsetLeft keeps padding.l so the column spans the cell and DocumentSkeleton positions right-aligned text
      expect(fit.width).toBe(100)
      expect(fit.offsetLeft).toBe(3)
    })

    it('computes symmetric margins for centered text in a normal cell', () => {
      const fit = computeEditorFit({
        cellWidth: 100,
        cellHeight: 25,
        textWidth: 40,
        textHeight: 15,
        padding: { t: 2, b: 2, l: 2, r: 2 },
        scaleX: 1,
        scaleY: 1,
        horizontalAlign: HorizontalAlign.CENTER,
        verticalAlign: VerticalAlign.MIDDLE,
      })
      expect(fit.width).toBe(100)
      // DocSkeleton centers glyphs within column width: (column.width - textWidth) / 2.
      // offsetLeft keeps padding.l so the column starts at cell margin and text stays centered.
      expect(fit.offsetLeft).toBe(2)
      // (cellHeight - textHeight) / 2 = (25 - 15) / 2 = 5
      expect(fit.offsetTop).toBe(5)
    })

    it('adjusts offsetTop for asymmetric padding when vertically centered', () => {
      const fit = computeEditorFit({
        cellWidth: 100,
        cellHeight: 25,
        textWidth: 40,
        textHeight: 15,
        padding: { t: 0, b: 2, l: 2, r: 2 },
        scaleX: 1,
        scaleY: 1,
        horizontalAlign: HorizontalAlign.CENTER,
        verticalAlign: VerticalAlign.MIDDLE,
      })
      // (cellHeight - textHeight) / 2 + (0 - 2) / 2 = 5 - 1 = 4
      expect(fit.offsetTop).toBe(4)
    })

    it('clamps offsetTop to top padding when text overflows cell height', () => {
      const fit = computeEditorFit({
        cellWidth: 100,
        cellHeight: 25,
        textWidth: 40,
        textHeight: 35,
        padding: { t: 2, b: 2, l: 2, r: 2 },
        scaleX: 1,
        scaleY: 1,
        horizontalAlign: HorizontalAlign.CENTER,
        verticalAlign: VerticalAlign.MIDDLE,
      })
      expect(fit.offsetTop).toBe(2)
    })

    it('centers caret in an empty cell', () => {
      const fit = computeEditorFit({
        cellWidth: 100,
        cellHeight: 25,
        textWidth: 0,
        textHeight: 15,
        padding: { t: 2, b: 2, l: 2, r: 2 },
        scaleX: 1,
        scaleY: 1,
        horizontalAlign: HorizontalAlign.CENTER,
        verticalAlign: VerticalAlign.MIDDLE,
      })
      expect(fit.width).toBe(100)
      expect(fit.offsetLeft).toBe(2)
    })

    it('computes centered margins for merged cells', () => {
      const fit = computeEditorFit({
        cellWidth: 300,
        cellHeight: 50,
        textWidth: 60,
        textHeight: 15,
        padding: { t: 2, b: 2, l: 2, r: 2 },
        scaleX: 1,
        scaleY: 1,
        horizontalAlign: HorizontalAlign.CENTER,
        verticalAlign: VerticalAlign.MIDDLE,
      })
      expect(fit.width).toBe(300)
      expect(fit.offsetLeft).toBe(2)
    })

    it('clamps centered text to minimum padding when overflowing cell', () => {
      const fit = computeEditorFit({
        cellWidth: 50,
        cellHeight: 25,
        textWidth: 100,
        textHeight: 15,
        padding: { t: 2, b: 2, l: 4, r: 4 },
        scaleX: 1,
        scaleY: 1,
        horizontalAlign: HorizontalAlign.CENTER,
        verticalAlign: VerticalAlign.MIDDLE,
      })
      // actualWidth = 100 + 8 = 108
      expect(fit.width).toBe(108)
      expect(fit.offsetLeft).toBe(4)
    })
  })

  describe('normalizeEditorLayoutObject', () => {
    it('sets zeroWidthParagraphBreak on document renderConfig', () => {
      const layout = {
        paddingData: { l: 4 },
        documentModel: {
          documentStyle: {
            renderConfig: {},
            marginLeft: 0,
            marginRight: 0,
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

  describe('normalizeEditorSkeletonGlyphs', () => {
    it('zeros trailing \\r paragraph terminator glyph width', () => {
      const skeletonData: any = {
        pages: [
          {
            renderConfig: { horizontalAlign: HorizontalAlign.LEFT },
            sections: [
              {
                columns: [
                  {
                    lines: [
                      {
                        paragraphIndex: 0,
                        divides: [
                          {
                            width: 100,
                            paddingLeft: 0,
                            glyphGroupWidth: 24,
                            glyphGroup: [
                              { content: 'A', width: 10, left: 0, xOffset: 0, bBox: { width: 10 } },
                              { content: 'B', width: 10, left: 10, xOffset: 0, bBox: { width: 10 } },
                              { content: '\r', width: 4, left: 20, xOffset: 0, bBox: { width: 4 } },
                            ],
                          },
                        ],
                      },
                    ],
                  },
                ],
              },
            ],
          },
        ],
      }
      normalizeEditorSkeletonGlyphs(skeletonData)
      const divide = skeletonData.pages[0].sections[0].columns[0].lines[0].divides[0]
      expect(divide.glyphGroup[2].width).toBe(0)
      expect(divide.glyphGroupWidth).toBe(20)
      expect(skeletonData.pages[0].sections[0].columns[0].lines[0].width).toBe(20)
    })

    it('normalizes CJK glyphs adjacent to Latin characters by removing artificial spacing', () => {
      const skeletonData: any = {
        pages: [
          {
            renderConfig: { horizontalAlign: HorizontalAlign.CENTER },
            sections: [
              {
                columns: [
                  {
                    lines: [
                      {
                        paragraphIndex: 0,
                        divides: [
                          {
                            width: 200,
                            paddingLeft: 0,
                            glyphGroupWidth: 40,
                            glyphGroup: [
                              { content: 'A', width: 10, left: 0, xOffset: 0, bBox: { width: 10 } },
                              { content: 'I', width: 6, left: 10, xOffset: 0, bBox: { width: 6 } },
                              // Univer addCJKLatinSpacing inflated width from 16 to 20 and added xOffset: 4
                              { content: '驅', width: 20, left: 16, xOffset: 4, bBox: { width: 16 } },
                              { content: '\r', width: 4, left: 36, xOffset: 0, bBox: { width: 4 } },
                            ],
                          },
                        ],
                      },
                    ],
                  },
                ],
              },
            ],
          },
        ],
      }
      normalizeEditorSkeletonGlyphs(skeletonData)
      const divide = skeletonData.pages[0].sections[0].columns[0].lines[0].divides[0]
      const quGlyph = divide.glyphGroup[2]
      expect(quGlyph.width).toBe(16)
      expect(quGlyph.xOffset).toBe(0)
      expect(quGlyph.left).toBe(16) // 10 + 6
      expect(divide.glyphGroup[3].width).toBe(0)
      expect(divide.glyphGroupWidth).toBe(32) // 10 + 6 + 16
      // Centered in width 200: (200 - 32) / 2 = 84
      expect(divide.paddingLeft).toBe(84)
    })

    it('normalizes line heights, ascents, and page height across lines', () => {
      const skeletonData: any = {
        pages: [
          {
            height: 36,
            sections: [
              {
                columns: [
                  {
                    lines: [
                      {
                        top: 0,
                        asc: 14,
                        lineHeight: 19,
                        divides: [
                          {
                            glyphGroup: [
                              { content: '執', width: 14, left: 0, fontStyle: '10pt sans-serif' },
                            ],
                          },
                        ],
                      },
                      {
                        top: 19,
                        asc: 13,
                        lineHeight: 17,
                        divides: [
                          {
                            glyphGroup: [
                              { content: '(', width: 7, left: 0, fontStyle: '10pt sans-serif' },
                            ],
                          },
                        ],
                      },
                    ],
                  },
                ],
              },
            ],
          },
        ],
      }
      normalizeEditorSkeletonGlyphs(skeletonData)
      const column = skeletonData.pages[0].sections[0].columns[0]
      // In a DOM environment FontCache measures font bounding box; even without DOM,
      // lines top positions accumulate monotonically.
      expect(column.lines[0].top).toBe(0)
      expect(column.lines[1].top).toBe(column.lines[0].lineHeight)
      expect(skeletonData.pages[0].height).toBe(column.lines[0].lineHeight + column.lines[1].lineHeight)
    })
  })
})
