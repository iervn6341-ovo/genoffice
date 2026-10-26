import { describe, expect, it } from 'vitest'

import {
  handleRibbonCommand,
  MULTI_RANGE_STYLE_COMMANDS,
  type RibbonCommandContext,
} from '../src/renderer/ribbon-actions'

interface Block {
  startRow: number
  endRow: number
  startColumn: number
  endColumn: number
}

/// Two-range selection A1:A3 + C1:C3; the active (primary) range is the last one, like Univer.
function harness(initialBold: boolean[], sizes: Record<string, number> = {}) {
  const blocks: Block[] = [
    { startRow: 0, endRow: 2, startColumn: 0, endColumn: 0 },
    { startRow: 0, endRow: 2, startColumn: 2, endColumn: 2 },
  ]
  const calls: string[] = []
  const makeRange = (index: number, b: Block) => ({
    getRange: () => b,
    getRow: () => b.startRow,
    getColumn: () => b.startColumn,
    getHeight: () => b.endRow - b.startRow + 1,
    getWidth: () => b.endColumn - b.startColumn + 1,
    getCellStyleData: () => ({ bl: initialBold[index] ? 1 : 0 }),
    setFontWeight: (value: string | null) => calls.push(`bold:${index}:${value}`),
    setFontSize: (value: number) => calls.push(`size:${index}:${value}`),
    setFontColor: (value: string) => calls.push(`color:${index}:${value}`),
  })
  const ranges = blocks.map((b, i) => makeRange(i, b))
  const worksheet = {
    getSheetId: () => 's',
    getMaxRows: () => 1000,
    getSheet: () => ({
      getComposedCellStyle: (r: number, c: number) => ({ fs: sizes[`${r},${c}`] ?? 11 }),
    }),
    getSelection: () => ({ getActiveRangeList: () => ranges }),
    getRange: (row: number, col: number, h: number, w: number) => ({
      setFontSize: (value: number) => calls.push(`block:${row},${col},${h},${w}:${value}`),
    }),
  }
  const workbook = {
    getActiveSheet: () => worksheet,
    getActiveRange: () => ranges[1],
    getId: () => 'wb',
  }
  const ctx = {
    univerRef: { current: { univerAPI: { getActiveWorkbook: () => workbook } } },
    setMessage: () => {},
    lazyWorkbookRef: { current: null },
  } as unknown as RibbonCommandContext
  return { ctx, calls }
}

describe('multi-range ribbon formatting', () => {
  it('bold reaches every range, not only the active one', () => {
    const { ctx, calls } = harness([false, false])
    handleRibbonCommand(ctx, 'bold')
    expect(calls).toEqual(['bold:0:bold', 'bold:1:bold'])
  })

  it('a bold toggle follows the active range for the whole selection', () => {
    // first range already bold, active range not: Excel makes everything bold
    const a = harness([true, false])
    handleRibbonCommand(a.ctx, 'bold')
    expect(a.calls).toEqual(['bold:0:bold', 'bold:1:bold'])
    // active range bold: everything is un-bolded
    const b = harness([false, true])
    handleRibbonCommand(b.ctx, 'bold')
    expect(b.calls).toEqual(['bold:0:null', 'bold:1:null'])
  })

  it('font colour reaches every range', () => {
    const { ctx, calls } = harness([false, false])
    handleRibbonCommand(ctx, 'font-color:#ff0000')
    expect(calls).toEqual(['color:0:#ff0000', 'color:1:#ff0000'])
  })

  it('font-size-step walks each cell from its own size across both ranges', () => {
    const { ctx, calls } = harness([false, false], { '0,0': 10, '1,0': 20, '0,2': 20 })
    handleRibbonCommand(ctx, 'font-size-step:1')
    expect(calls).toContain('block:0,0,1,1:11')
    expect(calls).toContain('block:1,0,1,1:22')
    expect(calls).toContain('block:0,2,1,1:22')
    // everything else is 11 -> 12
    expect(calls).toContain('block:1,2,2,1:12')
  })

  it('font-size-step down uses the ladder too', () => {
    const { ctx, calls } = harness([false, false], { '0,0': 24 })
    handleRibbonCommand(ctx, 'font-size-step:-1')
    expect(calls).toContain('block:0,0,1,1:22')
  })

  it('only formatting commands fan out over ranges', () => {
    expect(MULTI_RANGE_STYLE_COMMANDS.has('bold')).toBe(true)
    expect(MULTI_RANGE_STYLE_COMMANDS.has('font-size')).toBe(true)
    expect(MULTI_RANGE_STYLE_COMMANDS.has('sort')).toBe(false)
    expect(MULTI_RANGE_STYLE_COMMANDS.has('merge')).toBe(false)
  })
})

describe('in-cell editing: ribbon formats the selected text, not the cell', () => {
  it('maps character formats to the Univer set-range commands', async () => {
    const { inCellEditPlan } = await import('../src/renderer/ribbon-actions')
    expect(inCellEditPlan('bold')).toEqual({ id: 'sheet.command.set-range-bold' })
    expect(inCellEditPlan('italic')).toEqual({ id: 'sheet.command.set-range-italic' })
    expect(inCellEditPlan('underline')).toEqual({ id: 'sheet.command.set-range-underline' })
    expect(inCellEditPlan('strike')).toEqual({ id: 'sheet.command.set-range-stroke' })
    expect(inCellEditPlan('font-size:14')).toEqual({
      id: 'sheet.command.set-range-fontsize',
      params: { value: 14 },
    })
    expect(inCellEditPlan('font-family:Arial')).toEqual({
      id: 'sheet.command.set-range-font-family',
      params: { value: 'Arial' },
    })
    expect(inCellEditPlan('font-color:#ff0000')).toEqual({
      id: 'sheet.command.set-range-text-color',
      params: { value: '#ff0000' },
    })
    expect(inCellEditPlan('font-size-step:1')?.id).toBe('sheet.command.set-range-font-increase')
    expect(inCellEditPlan('font-size-step:-1')?.id).toBe('sheet.command.set-range-font-decrease')
  })

  it('leaves cell-level formats and absolute dialog forms alone', async () => {
    const { inCellEditPlan } = await import('../src/renderer/ribbon-actions')
    expect(inCellEditPlan('fill:#ff0000')).toBeNull()
    expect(inCellEditPlan('align:center')).toBeNull()
    expect(inCellEditPlan('bold:on')).toBeNull()
    expect(inCellEditPlan('underline:double')).toBeNull()
    expect(inCellEditPlan('font-color:auto')).toBeNull()
    expect(inCellEditPlan('font-size:0')).toBeNull()
  })

  it('bold while a cell is open runs the inline command and does not restyle the cell', () => {
    const executed: Array<{ id: string; params: unknown }> = []
    const { ctx, calls } = harness([false, false])
    const runtime = (
      ctx as unknown as { univerRef: { current: { univerAPI: Record<string, unknown> } } }
    ).univerRef.current
    const workbook = (runtime.univerAPI.getActiveWorkbook as () => Record<string, unknown>)()
    workbook.isCellEditing = () => true
    runtime.univerAPI.executeCommand = async (id: string, params: unknown) => {
      executed.push({ id, params })
      return true
    }
    handleRibbonCommand(ctx, 'bold')
    expect(executed).toEqual([{ id: 'sheet.command.set-range-bold', params: undefined }])
    expect(calls).toEqual([])
  })
})
