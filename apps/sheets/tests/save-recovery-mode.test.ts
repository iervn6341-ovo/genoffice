/**
 * A dirty workbook gets a crash-recovery copy through the
 * same save pipeline, but recovery mode must never touch the opened file, prompt,
 * or clear the journal — the whole point is that unsaved work survives a crash.
 */
import JSZip from 'jszip'
import { describe, expect, it, vi, beforeEach } from 'vitest'
import { applyCellEditsToXlsx } from '@genoffice/xlsx-gateway/gateway/xlsx-gateway'
import { handleSave, type SaveContext } from '../src/renderer/save-actions'
import { t } from '../src/renderer/i18n/locale'
import { createEditJournal, recordSetRangeValues } from '../src/renderer/edit-journal'
import { buildEditFixture } from './fixture-builder'
import { setToastEmitter, type ToastData } from '../src/renderer/toast-bus'

const saveWorkbookEdits = vi.fn()
const writeWorkbookRecovery = vi.fn()

beforeEach(() => {
  saveWorkbookEdits.mockReset().mockResolvedValue({ canceled: true })
  writeWorkbookRecovery.mockReset().mockResolvedValue({ ok: true })
  ;(globalThis as unknown as { window: unknown }).window = {
    desktopApi: { saveWorkbookEdits, writeWorkbookRecovery },
  }
})

function ctxWith(opts: { dirty: boolean; needsSaveAs?: boolean; restoredFromRecovery?: boolean }): {
  ctx: SaveContext
  messages: string[]
  journal: ReturnType<typeof createEditJournal>
  overlay: Map<string, Map<string, { v?: string | number | boolean | null }>>
} {
  const journal = createEditJournal()
  if (opts.dirty) recordSetRangeValues(journal, 'sheet-1', { 0: { 0: { v: 'edited' } } })
  const messages: string[] = []
  const overlay = new Map<string, Map<string, { v?: string | number | boolean | null }>>()
  return {
    messages,
    journal,
    overlay,
    ctx: {
      univerRef: { current: null },
      stashViewRestore: () => {},
      lazyWorkbookRef: {
        current: {
          editJournal: journal,
          recalc: {
            timer: null,
            generation: 0,
            failed: false,
            formulaCells: new Map(),
            overlay,
          },
          file: {
            sessionId: '11111111-1111-4111-8111-111111111111',
            needsSaveAs: !!opts.needsSaveAs,
            restoredFromRecovery: !!opts.restoredFromRecovery,
          },
        },
      } as never,
      setMessage: (m: string) => messages.push(m),
      openLazyWorkbook: () => {},
    },
  }
}

describe('handleSave recovery mode', () => {
  it('writes a recovery copy and never the opened file', async () => {
    const { ctx, messages } = ctxWith({ dirty: true })
    await handleSave(ctx, 'recovery')
    expect(writeWorkbookRecovery).toHaveBeenCalledTimes(1)
    expect(saveWorkbookEdits).not.toHaveBeenCalled()
    // silent: no status messages for a background copy
    expect(messages).toEqual([])
  })

  it('sends the same payload shape the save pipeline gets', async () => {
    const { ctx } = ctxWith({ dirty: true })
    await handleSave(ctx, 'recovery')
    const payload = writeWorkbookRecovery.mock.calls[0]![0] as Record<string, unknown>
    expect(payload.sessionId).toBe('11111111-1111-4111-8111-111111111111')
    // 'recovery' is a renderer-side mode; the main process still sees a normal save request
    expect(payload.mode).toBe('save')
    expect(Array.isArray(payload.edits)).toBe(true)
    expect((payload.edits as unknown[]).length).toBe(1)
    expect(payload).toHaveProperty('structuralOps')
    expect(payload).toHaveProperty('sparklineAdditions')
  })

  it('does nothing when there is nothing pending', async () => {
    const { ctx, messages } = ctxWith({ dirty: false })
    await handleSave(ctx, 'recovery')
    expect(writeWorkbookRecovery).not.toHaveBeenCalled()
    expect(messages).toEqual([])
  })

  it('a failed copy is swallowed (best-effort, never surfaces; resolves not-ok)', async () => {
    writeWorkbookRecovery.mockRejectedValue(new Error('disk full'))
    const { ctx, messages } = ctxWith({ dirty: true })
    // handleSave now reports an outcome (the MCP bridge reads it); recovery
    // mode still stays silent on failure
    await expect(handleSave(ctx, 'recovery')).resolves.toEqual({ ok: false })
    expect(messages).toEqual([])
  })
})

describe('handleSave formula cache overlay', () => {
  it('refuses a normal save when the imported-formula calculation has not applied', async () => {
    const { ctx, journal, messages } = ctxWith({ dirty: false })
    recordSetRangeValues(journal, 'sheet-1', { 0: { 2: { v: 8 } } })
    Object.assign(ctx.lazyWorkbookRef.current!, {
      formulaMode: true,
      flags: { preloadComplete: true, preloadRunning: false },
      closure: { status: 'idle', pinned: new Map() },
      formulaText: new Map([['sheet-1', new Map([['2:1', '=SUM(C1:C2)']])]]),
    })
    Object.assign(ctx.univerRef, {
      current: {
        univer: { __getInjector: () => ({ get: () => ({ beforeCommandExecuted: vi.fn() }) }) },
        univerAPI: {
          getActiveWorkbook: () => null,
          getFormula: () => ({
            executeCalculation: vi.fn(),
            onCalculationResultApplied: vi.fn().mockRejectedValue(new Error('timeout')),
          }),
        },
      },
    })
    Object.assign(ctx, { readCells: () => ({ B3: { value: 5, formula: '=SUM(C1:C2)' } }) })

    await expect(handleSave(ctx, 'save')).resolves.toEqual({ ok: false })
    expect(saveWorkbookEdits).not.toHaveBeenCalled()
    expect(messages.at(-1)).toBe(t('appSaveCalcPending'))
  })

  it.each(['preloaded', 'closure'] as const)(
    'refreshes an unedited imported dependent formula in %s mode',
    async (mode) => {
      const { ctx, journal } = ctxWith({ dirty: false })
      // The fixture's B3 has =SUM(C1:C2) and an old cached result of 5.
      recordSetRangeValues(journal, 'sheet-1', { 0: { 2: { v: 8 } } })
      const state = ctx.lazyWorkbookRef.current!
      Object.assign(state, {
        formulaMode: mode === 'preloaded',
        flags: { preloadComplete: mode === 'preloaded', preloadRunning: false },
        closure: { status: mode === 'closure' ? 'active' : 'idle', pinned: new Map() },
        formulaText: new Map([['sheet-1', new Map([['2:1', '=SUM(C1:C2)']])]]),
      })
      const executeCalculation = vi.fn()
      const onCalculationResultApplied = vi.fn().mockResolvedValue(undefined)
      Object.assign(ctx.univerRef, {
        current: {
          univer: {
            __getInjector: () => ({ get: () => ({ beforeCommandExecuted: vi.fn() }) }),
          },
          univerAPI: {
            getActiveWorkbook: () => null,
            getFormula: () => ({ executeCalculation, onCalculationResultApplied }),
          },
        },
      })
      Object.assign(ctx, {
        readCells: () => ({
          B3: { value: 8, rawValue: 8, formula: '=SUM(C1:C2)' },
        }),
      })

      await handleSave(ctx, 'save')

      expect(executeCalculation).toHaveBeenCalledTimes(1)
      expect(onCalculationResultApplied).toHaveBeenCalledTimes(1)
      const payload = saveWorkbookEdits.mock.calls[0]![0] as {
        edits: { row: number; column: number; writeValue: boolean; value: unknown }[]
        formulaValues: { sheetId: string; row: number; column: number; value: unknown }[]
      }
      expect(payload.formulaValues).toContainEqual({
        sheetId: 'sheet-1',
        row: 2,
        column: 1,
        value: 8,
      })
      const mutation = await applyCellEditsToXlsx(
        await buildEditFixture(),
        payload.edits.map((edit) => ({
          sheetName: 'Data',
          row: edit.row,
          column: edit.column,
          writeValue: edit.writeValue,
          cell: { value: edit.value as string | number | boolean | null },
        })),
        [],
        [],
        undefined,
        [],
        [],
        [],
        [],
        [],
        null,
        [],
        [],
        [
          {
            sheetName: 'Data',
            cells: payload.formulaValues.map(({ row, column, value }) => ({
              row,
              column,
              value: value as number,
            })),
          },
        ],
      )
      const zip = await JSZip.loadAsync(mutation.buffer)
      const xml = (await zip.file('xl/worksheets/sheet1.xml')?.async('text')) ?? ''
      expect(xml).toMatch(/<c r="B3"[^>]*><f>SUM\(C1:C2\)<\/f><v>8<\/v><\/c>/)
    },
  )

  it('never pairs an immediately saved replacement formula with the previous cached value', async () => {
    const { ctx, journal, overlay } = ctxWith({ dirty: false })
    recordSetRangeValues(journal, 'sheet-1', { 2: { 1: { f: '=2+2' } } })
    overlay.set(
      'sheet-1',
      new Map([
        // B3 used to be SUM(C1:C2), whose cached value was 5.
        ['2:1', { v: 5 }],
        ['0:2', { v: 7 }],
      ]),
    )
    let savedWorksheet = ''
    saveWorkbookEdits.mockImplementationOnce(async (payload) => {
      const request = payload as {
        edits: {
          row: number
          column: number
          writeValue: boolean
          value: string | number | boolean | null
          formula?: string
        }[]
        formulaValues: {
          sheetId: string
          row: number
          column: number
          value: string | number | boolean | null
        }[]
      }
      const mutation = await applyCellEditsToXlsx(
        await buildEditFixture(),
        request.edits.map((edit) => ({
          sheetName: 'Data',
          row: edit.row,
          column: edit.column,
          writeValue: edit.writeValue,
          cell: {
            value: edit.value,
            ...(edit.formula === undefined ? {} : { formula: edit.formula }),
          },
        })),
        [],
        [],
        undefined,
        [],
        [],
        [],
        [],
        [],
        null,
        [],
        [],
        [
          {
            sheetName: 'Data',
            cells: request.formulaValues.map(({ row, column, value }) => ({
              row,
              column,
              value,
            })),
          },
        ],
      )
      const zip = await JSZip.loadAsync(mutation.buffer)
      savedWorksheet = await zip.file('xl/worksheets/sheet1.xml')!.async('text')
      return { canceled: true }
    })

    await handleSave(ctx, 'save')

    const payload = saveWorkbookEdits.mock.calls[0]![0] as {
      edits: { row: number; column: number; formula?: string }[]
      formulaValues: { row: number; column: number; value: unknown }[]
    }
    expect(payload.edits).toContainEqual(
      expect.objectContaining({ row: 2, column: 1, formula: '=2+2' }),
    )
    expect(payload.formulaValues).toEqual([{ sheetId: 'sheet-1', row: 0, column: 2, value: 7 }])
    const savedB3 = /<c r="B3"[^>]*>[\s\S]*?<\/c>/.exec(savedWorksheet)?.[0]
    expect(savedB3).toContain('<f>2+2</f>')
    expect(savedB3).not.toContain('<v>')
  })

  it("never writes IronCalc's #ERROR! into a cached <v> (issue 235)", async () => {
    const { ctx, journal, overlay } = ctxWith({ dirty: false })
    recordSetRangeValues(journal, 'sheet-1', { 5: { 5: { v: 1 } } })
    overlay.set(
      'sheet-1',
      new Map([
        ['0:0', { v: '#ERROR!' }],
        ['0:1', { v: '#N/A' }],
        ['0:2', { v: 7 }],
      ]),
    )
    await handleSave(ctx, 'save')
    const payload = saveWorkbookEdits.mock.calls[0]![0] as {
      formulaValues: { row: number; column: number; value: unknown }[]
    }
    expect(payload.formulaValues).toEqual([
      { sheetId: 'sheet-1', row: 0, column: 1, value: '#N/A' },
      { sheetId: 'sheet-1', row: 0, column: 2, value: 7 },
    ])
  })
})

describe('handleSave restored-recovery write-back', () => {
  it('a clean restored session still saves: the workbook bytes are the change', async () => {
    const { ctx } = ctxWith({ dirty: false, restoredFromRecovery: true })
    await handleSave(ctx, 'save')
    expect(saveWorkbookEdits).toHaveBeenCalledTimes(1)
    const payload = saveWorkbookEdits.mock.calls[0]![0] as Record<string, unknown>
    expect(payload.mode).toBe('save')
    expect(payload.restoreWriteBack).toBe(true)
  })

  it('a clean ordinary session still refuses a pointless save', async () => {
    const { ctx, messages } = ctxWith({ dirty: false })
    await handleSave(ctx, 'save')
    expect(saveWorkbookEdits).not.toHaveBeenCalled()
    expect(messages.length).toBe(1)
  })

  it('never flags recovery-mode writes as a restore write-back', async () => {
    const { ctx } = ctxWith({ dirty: true, restoredFromRecovery: true })
    await handleSave(ctx, 'recovery')
    expect(writeWorkbookRecovery).toHaveBeenCalledTimes(1)
    const payload = writeWorkbookRecovery.mock.calls[0]![0] as Record<string, unknown>
    expect(payload.restoreWriteBack).toBeUndefined()
  })
})

describe('recovery copy failure is visible', () => {
  const toasts: ToastData[] = []
  beforeEach(() => {
    toasts.length = 0
    setToastEmitter((toast) => toasts.push(toast))
    vi.useFakeTimers()
    vi.setSystemTime(new Date('2026-10-15T10:00:00Z'))
  })

  const fail = (error: string): void => {
    writeWorkbookRecovery.mockResolvedValue({ ok: false, error })
  }

  it('shows an error toast with the reason, once, while the same failure repeats', async () => {
    fail('reason-A')
    const { ctx } = ctxWith({ dirty: true })
    await handleSave(ctx, 'recovery')
    await handleSave(ctx, 'recovery')
    expect(toasts).toHaveLength(1)
    expect(toasts[0]?.kind).toBe('error')
    expect(toasts[0]?.text).toContain('reason-A')
  })

  it('warns again when the reason changes, or after ten minutes', async () => {
    fail('reason-B')
    const { ctx } = ctxWith({ dirty: true })
    await handleSave(ctx, 'recovery')
    fail('reason-C')
    await handleSave(ctx, 'recovery')
    expect(toasts.map((toast) => toast.text.includes('reason-C'))).toEqual([false, true])
    vi.setSystemTime(new Date('2026-10-15T10:11:00Z'))
    await handleSave(ctx, 'recovery')
    expect(toasts).toHaveLength(3)
  })

  it('stays silent when the copy is simply skipped (no error reported)', async () => {
    writeWorkbookRecovery.mockResolvedValue({ ok: false })
    const { ctx } = ctxWith({ dirty: true })
    await handleSave(ctx, 'recovery')
    expect(toasts).toHaveLength(0)
  })
})
