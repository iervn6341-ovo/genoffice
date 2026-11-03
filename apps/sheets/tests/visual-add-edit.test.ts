import { describe, expect, it } from 'vitest'

import {
  createEditJournal,
  journalSize,
  recordChartEdit,
  recordSheetRemove,
  recordVisualAdd,
  recordVisualEdit,
  removeVisualAdd,
  toSaveChartEdits,
  toSaveVisualEdits,
  updateVisualAdd,
  toSaveVisualAdds,
  shiftCellImagesForStructuralOp,
  shiftVisualAnchor,
  shiftVisualForStructuralOp,
} from '../src/renderer/edit-journal'
import {
  workbookChartEditSchema,
  workbookVisualAddSchema,
  type WorkbookVisualObject,
} from '../src/shared/desktop-api'

const textBox = (): WorkbookVisualObject => ({
  id: 'added-shape-abc-1',
  sheetId: 'sheet-1',
  kind: 'shape',
  anchor: {
    fromRow: 0,
    fromColumn: 0,
    fromRowOffset: 0,
    fromColumnOffset: 0,
    toRow: 4,
    toColumn: 4,
    toRowOffset: 0,
    toColumnOffset: 0,
  },
  shapeType: 'rect',
  name: 'TextBox',
  fillColor: '#FFFFFF',
  text: 'Text',
})

describe('updateVisualAdd', () => {
  it('moves, resizes, and retexts a session visual; the save payload follows', () => {
    const journal = createEditJournal()
    recordVisualAdd(journal, textBox())

    const moved = { ...textBox().anchor, fromRow: 2, toRow: 6, fromColumn: 1, toColumn: 5 }
    expect(updateVisualAdd(journal, 'added-shape-abc-1', { anchor: moved })).toBe(true)
    expect(updateVisualAdd(journal, 'added-shape-abc-1', { text: 'Hello' })).toBe(true)

    const visual = journal.visualAdds[0]
    expect(visual?.anchor).toEqual(moved)
    expect(visual?.text).toBe('Hello')
    expect(visual?.fillColor).toBe('#FFFFFF')

    const additions = toSaveVisualAdds(journal)
    expect(additions).toHaveLength(1)
    expect(additions[0]?.anchor).toEqual(moved)
    expect(additions[0]?.shape).toMatchObject({ text: 'Hello', isTextBox: true })
  })

  it('rejects unknown ids (file visuals go through recordVisualEdit instead)', () => {
    const journal = createEditJournal()
    recordVisualAdd(journal, textBox())
    expect(updateVisualAdd(journal, 'file-visual-1', { text: 'nope' })).toBe(false)
    expect(journal.visualAdds[0]?.text).toBe('Text')
  })

  it('drops a session visual outright on removal', () => {
    const journal = createEditJournal()
    recordVisualAdd(journal, textBox())
    expect(removeVisualAdd(journal, 'added-shape-abc-1')).toBe(true)
    expect(journal.visualAdds).toHaveLength(0)
    expect(removeVisualAdd(journal, 'added-shape-abc-1')).toBe(false)
  })

  it('projects baked chart state (legend/labels/axis titles/colors) into the save payload', () => {
    const journal = createEditJournal()
    recordVisualAdd(journal, {
      ...textBox(),
      id: 'added-chart-abc-1',
      kind: 'chart',
      chart: {
        chartTypes: ['pieChart'],
        title: 'Regional Share',
        legend: 'bottom',
        dataLabels: 'category-percent',
        dataLabelPosition: 'outside-end',
        dataLabelFormat: '0.0%',
        axisTitles: { category: 'Month', value: null },
        series: [
          {
            name: 'S1',
            categories: ['a', 'b'],
            values: [1, 2],
            color: '#4472C4',
            pointColors: [
              { index: 0, color: '#ED7D31' },
              { index: 1, color: '#70AD47' },
            ],
          },
        ],
      },
    })
    const chart = toSaveVisualAdds(journal)[0]?.chart
    expect(chart).toMatchObject({
      chartType: 'pie',
      legend: 'bottom',
      dataLabels: 'category-percent',
      dataLabelPosition: 'outside-end',
      dataLabelFormat: '0.0%',
      axisTitles: { category: 'Month' },
    })
    expect(chart?.axisTitles).not.toHaveProperty('value')
    expect(chart?.series[0]).toMatchObject({
      color: '#4472C4',
      pointColors: { '0': '#ED7D31', '1': '#70AD47' },
    })
  })
})

describe('save-wire string clamping', () => {
  const long = (length: number): string => 'x'.repeat(length)

  it('clamps cell-derived chart strings so a long cell never fails the save schema', () => {
    const journal = createEditJournal()
    recordVisualAdd(journal, {
      ...textBox(),
      id: 'added-chart-long-1',
      kind: 'chart',
      chart: {
        chartTypes: ['barChart'],
        barDirection: 'col',
        title: long(400),
        axisTitles: { category: long(400) },
        series: [{ name: long(300), categories: ['short', long(3_000)], values: [1, 2] }],
      },
    })
    const addition = toSaveVisualAdds(journal)[0]
    expect(() => workbookVisualAddSchema.parse(addition)).not.toThrow()
    expect(addition?.chart?.title).toHaveLength(255)
    expect(addition?.chart?.axisTitles?.category).toHaveLength(255)
    const series = addition?.chart?.series[0]
    expect(series?.name).toHaveLength(255)
    expect(series?.categories).toEqual(['short', long(1_024)])
  })

  it('clamps chart-edit series rewrites the same way', () => {
    const journal = createEditJournal()
    recordChartEdit(journal, 'xl/charts/chart1.xml', {
      title: long(400),
      series: [{ index: 0, name: long(300), categories: [long(3_000)], values: [1] }],
    })
    const edit = toSaveChartEdits(journal)[0]
    expect(() => workbookChartEditSchema.parse(edit)).not.toThrow()
    expect(edit?.title).toHaveLength(255)
    expect(edit?.series?.[0]?.name).toHaveLength(255)
    expect(edit?.series?.[0]?.categories?.[0]).toHaveLength(1_024)
  })
})

describe('recordVisualEdit', () => {
  const fileImage = (): WorkbookVisualObject => ({
    ...textBox(),
    id: 'visual-3',
    kind: 'image',
    drawingPath: 'xl/drawings/drawing1.xml',
    drawingIndex: 3,
  })

  it('merges moves and removals per visual, removal winning', () => {
    const journal = createEditJournal()
    const moved = { ...fileImage().anchor, fromRow: 8, toRow: 12 }
    expect(recordVisualEdit(journal, fileImage(), { anchor: moved })).toBe(true)
    expect(recordVisualEdit(journal, fileImage(), { remove: true })).toBe(true)
    expect(journalSize(journal)).toBe(1)
    expect(toSaveVisualEdits(journal)).toEqual([
      {
        drawingPath: 'xl/drawings/drawing1.xml',
        drawingIndex: 3,
        remove: true,
        anchor: moved,
      },
    ])
  })

  it('refuses visuals without a drawing locator', () => {
    const journal = createEditJournal()
    const visual = { ...fileImage(), drawingPath: undefined, drawingIndex: undefined }
    expect(recordVisualEdit(journal, visual, { remove: true })).toBe(false)
    expect(toSaveVisualEdits(journal)).toEqual([])
  })

  it('drops edits on removed sheets from the save payload', () => {
    const journal = createEditJournal()
    recordVisualEdit(journal, fileImage(), { remove: true })
    recordSheetRemove(journal, 'sheet-1')
    expect(toSaveVisualEdits(journal)).toEqual([])
  })
})

describe('shiftVisualAnchor with structural ops', () => {
  const baseAnchor = {
    fromRow: 2,
    fromColumn: 2,
    fromRowOffset: 0,
    fromColumnOffset: 0,
    toRow: 6,
    toColumn: 6,
    toRowOffset: 0,
    toColumnOffset: 0,
  }

  it('shifts image right when inserting columns to the left', () => {
    const shifted = shiftVisualAnchor(baseAnchor, {
      axis: 'column',
      removing: false,
      index: 1,
      count: 2,
    })
    expect(shifted.fromColumn).toBe(4)
    expect(shifted.toColumn).toBe(8)
  })

  it('does NOT expand image when inserting columns directly at right edge with zero offset', () => {
    const shifted = shiftVisualAnchor(baseAnchor, {
      axis: 'column',
      removing: false,
      index: 6,
      count: 1,
    })
    expect(shifted.fromColumn).toBe(2)
    expect(shifted.toColumn).toBe(6)
  })

  it('does NOT expand image when inserting columns to the right of image', () => {
    const shifted = shiftVisualAnchor(baseAnchor, {
      axis: 'column',
      removing: false,
      index: 7,
      count: 2,
    })
    expect(shifted.fromColumn).toBe(2)
    expect(shifted.toColumn).toBe(6)
  })

  it('expands image when inserting columns inside the image', () => {
    const shifted = shiftVisualAnchor(baseAnchor, {
      axis: 'column',
      removing: false,
      index: 4,
      count: 2,
    })
    expect(shifted.fromColumn).toBe(2)
    expect(shifted.toColumn).toBe(8)
  })

  it('shifts image left when deleting columns to the left', () => {
    const shifted = shiftVisualAnchor(baseAnchor, {
      axis: 'column',
      removing: true,
      index: 0,
      count: 1,
    })
    expect(shifted.fromColumn).toBe(1)
    expect(shifted.toColumn).toBe(5)
  })

  it('shifts rows appropriately when inserting rows above', () => {
    const shifted = shiftVisualAnchor(baseAnchor, {
      axis: 'row',
      removing: false,
      index: 1,
      count: 3,
    })
    expect(shifted.fromRow).toBe(5)
    expect(shifted.toRow).toBe(9)
  })
})

describe('shiftCellImagesForStructuralOp', () => {
  const images = [
    { id: 'img-1', row: 2, column: 2 },
    { id: 'img-2', row: 5, column: 5 },
  ]

  it('shifts image column when inserting column to the left', () => {
    const shifted = shiftCellImagesForStructuralOp(images, {
      kind: 'insert-cols',
      index: 1,
      count: 2,
    })
    expect(shifted).toEqual([
      { id: 'img-1', row: 2, column: 4 },
      { id: 'img-2', row: 5, column: 7 },
    ])
  })

  it('leaves column unchanged when inserting column to the right', () => {
    const shifted = shiftCellImagesForStructuralOp(images, {
      kind: 'insert-cols',
      index: 6,
      count: 2,
    })
    expect(shifted).toEqual(images)
  })

  it('drops image if its row is removed', () => {
    const shifted = shiftCellImagesForStructuralOp(images, {
      kind: 'remove-rows',
      index: 2,
      count: 1,
    })
    expect(shifted).toEqual([{ id: 'img-2', row: 4, column: 5 }])
  })
})

