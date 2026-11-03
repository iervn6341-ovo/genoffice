import { describe, expect, it } from 'vitest'

import type { WorkbookVisualObject } from '../src/shared/desktop-api'
import {
  extractVisualPayload,
  isCopyableVisual,
  visualClipboardHtml,
  visualFromPayload,
  visualToPayload,
} from '../src/renderer/visual-clipboard'

const anchor = {
  fromRow: 2,
  fromColumn: 3,
  fromRowOffset: 100,
  fromColumnOffset: 200,
  toRow: 12,
  toColumn: 9,
  toRowOffset: 300,
  toColumnOffset: 400,
}

const shape: WorkbookVisualObject = {
  id: 'drawing1-3',
  sheetId: 'sheet-1',
  kind: 'shape',
  anchor,
  shapeType: 'roundRect',
  fillColor: '#DDEBF7',
  text: 'Note ✓',
  drawingPath: 'xl/drawings/drawing1.xml',
  drawingIndex: 3,
}

const image: WorkbookVisualObject = {
  id: 'drawing1-1',
  sheetId: 'sheet-1',
  kind: 'image',
  anchor,
  mediaPath: 'xl/media/image1.png',
  mediaType: 'image/png',
  name: 'logo.png',
  drawingPath: 'xl/drawings/drawing1.xml',
  drawingIndex: 1,
}

const chart: WorkbookVisualObject = {
  id: 'added-chart-1',
  sheetId: 'sheet-1',
  kind: 'chart',
  anchor,
  chart: {
    chartTypes: ['barChart'],
    title: 'Sales',
    series: [
      {
        name: 'Q1',
        categories: ['a', 'b'],
        values: [1, 2],
        valuesRef: "'Sheet1'!$B$2:$B$3",
        categoriesRef: "'Sheet1'!$A$2:$A$3",
      },
    ],
  },
}

describe('visual clipboard payload', () => {
  it('drops the file identity of a shape and keeps its look and text', () => {
    const payload = visualToPayload(shape, 'Sheet1', 'book-1')!
    expect(payload.visual).toMatchObject({ kind: 'shape', shapeType: 'roundRect', text: 'Note ✓' })
    expect(payload.visual).not.toHaveProperty('drawingPath')
    expect(payload.visual).not.toHaveProperty('id')
    expect(payload.spanRows).toBe(10)
    expect(payload.spanColumns).toBe(6)
  })

  it('needs the bytes of a picture', () => {
    expect(visualToPayload(image, 'Sheet1', 'book-1')).toBeNull()
    const payload = visualToPayload(image, 'Sheet1', 'book-1', 'data:image/png;base64,AAAA')!
    expect(payload.visual).toMatchObject({
      kind: 'image',
      mediaDataUrl: 'data:image/png;base64,AAAA',
    })
    expect(payload.visual).not.toHaveProperty('mediaPath')
  })

  it('does not copy read-only embeds', () => {
    expect(isCopyableVisual({ ...shape, kind: 'ole' })).toBe(false)
    expect(visualToPayload({ ...shape, kind: 'slicer' }, 'Sheet1', 'book-1')).toBeNull()
  })

  it('survives the clipboard markup, including non-ASCII text', () => {
    const payload = visualToPayload(shape, 'Sheet1', 'book-1')!
    const html = visualClipboardHtml(payload, 'a <b> & c')
    expect(html).toContain('a &lt;b&gt; &amp; c')
    expect(extractVisualPayload(html)).toEqual(payload)
    expect(extractVisualPayload('<span>nothing</span>')).toBeNull()
    expect(extractVisualPayload(null)).toBeNull()
  })

  it('rebuilds the object at the active cell with its size, offsets and a new identity', () => {
    const payload = visualToPayload(chart, 'Sheet1', 'book-1')!
    const placed = visualFromPayload(payload, {
      id: 'added-chart-9',
      sheetId: 'sheet-2',
      row: 20,
      column: 1,
    })
    expect(placed.id).toBe('added-chart-9')
    expect(placed.sheetId).toBe('sheet-2')
    expect(placed.anchor).toMatchObject({
      fromRow: 20,
      fromColumn: 1,
      toRow: 30,
      toColumn: 7,
      fromRowOffset: 100,
      toColumnOffset: 400,
    })
    expect(placed.chart?.series[0]?.values).toEqual([1, 2])
  })

  it('preserves exact pixel width and height when pasted with target sheet geometry', () => {
    // 350px wide x 220px high picture
    const pixelWidth = 350
    const pixelHeight = 220
    const payload = visualToPayload(
      image,
      'Sheet1',
      'book-1',
      'data:image/png;base64,AAAA',
      pixelWidth,
      pixelHeight,
    )!
    expect(payload.pixelWidth).toBe(350)
    expect(payload.pixelHeight).toBe(220)

    // Destination sheet has column width 100px and row height 50px
    const geometry = {
      columnWidth: () => 100,
      rowHeight: () => 50,
      maxColumns: 100,
      maxRows: 100,
    }

    const placed = visualFromPayload(
      payload,
      {
        id: 'added-picture-9',
        sheetId: 'sheet-2',
        row: 5,
        column: 2,
      },
      geometry,
    )

    // With 350px starting at column 2 (index 2):
    // col 2: 100px (remaining 250px)
    // col 3: 100px (remaining 150px)
    // col 4: 100px (remaining 50px)
    // col 5: 50px offset -> toColumn: 5, toColumnOffset: 50 * EMU_PER_PIXEL (50 * 9525 = 476250)
    // With 220px starting at row 5 (index 5):
    // row 5: 50px (remaining 170px)
    // row 6: 50px (remaining 120px)
    // row 7: 50px (remaining 70px)
    // row 8: 50px (remaining 20px)
    // row 9: 20px offset -> toRow: 9, toRowOffset: 20 * EMU_PER_PIXEL (20 * 9525 = 190500)
    expect(placed.anchor.fromRow).toBe(5)
    expect(placed.anchor.fromColumn).toBe(2)
    expect(placed.anchor.toColumn).toBe(5)
    expect(placed.anchor.toColumnOffset).toBe(50 * 9525)
    expect(placed.anchor.toRow).toBe(9)
    expect(placed.anchor.toRowOffset).toBe(20 * 9525)
    expect(placed.frameWidth).toBe(350 * 9525)
    expect(placed.frameHeight).toBe(220 * 9525)
  })
})

