import { describe, expect, it } from 'vitest'
import { ICommandService, IUndoRedoService } from '@univerjs/core'

import {
  alignmentDeltaOfCommand,
  protectionDeltaOfCommand,
  registerCellProtectionUndoRuntime,
  restoreProtectionSnapshot,
  withCellProtectionUndo,
  type ProtectionSnapshot,
} from '../src/renderer/cell-protection-undo'
import {
  createEditJournal,
  recordNeutralStyleEdit,
  toSaveEdits,
} from '../src/renderer/edit-journal'
import type { LazyWorkbookState } from '../src/renderer/univer-state'
import type { UniverRuntime } from '../src/renderer/univer-state'
import {
  captureUndoCarry,
  consumePendingUndoCarry,
  stashUndoCarry,
  type CarriedUndoItem,
} from '../src/renderer/undo-carry'

const state = (sessionId: string) =>
  ({ file: { sessionId }, editJournal: createEditJournal() }) as Pick<
    LazyWorkbookState,
    'file' | 'editJournal'
  >

describe('cell protection history', () => {
  it('parses only explicit flag changes', () => {
    expect(protectionDeltaOfCommand('cellprot:locked-off:')).toEqual({
      protectionLocked: false,
    })
    expect(protectionDeltaOfCommand('cellprot::hidden-on')).toEqual({
      protectionHidden: true,
    })
    expect(protectionDeltaOfCommand('cellprot::')).toBeNull()
    expect(protectionDeltaOfCommand('cellprot:locked-mixed:')).toBeNull()
    expect(alignmentDeltaOfCommand('align:centerContinuous')).toEqual({
      horizontalAlignment: 'centerContinuous',
    })
    expect(alignmentDeltaOfCommand('valign:distributed')).toEqual({
      verticalAlignment: 'distributed',
    })
    expect(alignmentDeltaOfCommand('rotate-ooxml:127')).toEqual({ textRotation: 127 })
    expect(alignmentDeltaOfCommand('shrink:off')).toEqual({ shrinkToFit: false })
    expect(alignmentDeltaOfCommand('rotate-ooxml:181')).toBeNull()
  })

  it('restores exact prior journal entries before save, including formula and other styles', () => {
    const current = state('original')
    const prior = {
      row: 0,
      column: 0,
      hasValue: true,
      value: 3,
      formula: '=1+2',
      style: { bold: true, protectionHidden: true },
    }
    current.editJournal.cells.set('sheet-1', new Map([['0:0', prior]]))
    recordNeutralStyleEdit(current.editJournal, 'sheet-1', 0, 0, { protectionLocked: false })
    recordNeutralStyleEdit(current.editJournal, 'sheet-1', 0, 1, { protectionLocked: false })
    const before: ProtectionSnapshot = {
      sessionId: 'original',
      cells: [
        { row: 0, column: 0, locked: true, hidden: true, entry: prior },
        { row: 0, column: 1, locked: true, hidden: false, entry: null },
      ],
    }
    restoreProtectionSnapshot(current, 'sheet-1', before)
    expect(current.editJournal.cells.get('sheet-1')?.get('0:0')).toEqual(prior)
    expect(current.editJournal.cells.get('sheet-1')?.has('0:1')).toBe(false)
  })

  it('restores only changed flags after save and leaves the opposite flag and content alone', () => {
    const current = state('saved')
    recordNeutralStyleEdit(current.editJournal, 'sheet-1', 0, 0, {
      italic: true,
      protectionHidden: true,
    })
    const before: ProtectionSnapshot = {
      sessionId: 'original',
      cells: [{ row: 0, column: 0, locked: true, hidden: false, entry: null }],
    }
    const after: ProtectionSnapshot = {
      sessionId: 'original',
      cells: [{ row: 0, column: 0, locked: false, hidden: false, entry: null }],
    }
    restoreProtectionSnapshot(current, 'sheet-1', before, after)
    expect(current.editJournal.cells.get('sheet-1')?.get('0:0')).toMatchObject({
      hasValue: false,
      style: { italic: true, protectionHidden: true, protectionLocked: true },
    })
  })

  it('cross-save undo restores only edited alignment fields', () => {
    const current = state('saved')
    recordNeutralStyleEdit(current.editJournal, 'sheet-1', 2, 1, { bold: true, shrinkToFit: true })
    const before: ProtectionSnapshot = {
      sessionId: 'original',
      cells: [
        {
          row: 2,
          column: 1,
          locked: true,
          hidden: false,
          horizontalAlignment: 'fill',
          verticalAlignment: 'justify',
          textRotation: 37,
          shrinkToFit: true,
          entry: null,
        },
      ],
    }
    const after: ProtectionSnapshot = {
      sessionId: 'original',
      cells: [
        {
          row: 2,
          column: 1,
          locked: true,
          hidden: false,
          horizontalAlignment: 'centerContinuous',
          verticalAlignment: 'justify',
          textRotation: 127,
          shrinkToFit: true,
          entry: null,
        },
      ],
    }
    restoreProtectionSnapshot(current, 'sheet-1', before, after)
    expect(current.editJournal.cells.get('sheet-1')?.get('2:1')).toMatchObject({
      hasValue: false,
      style: { bold: true, shrinkToFit: true, horizontalAlignment: 'fill', textRotation: 37 },
    })
    expect(
      current.editJournal.cells.get('sheet-1')?.get('2:1')?.style?.verticalAlignment,
    ).toBeUndefined()
  })

  it('cross-save undo and redo carry the original XF identity, including applyAlignment distinctions', () => {
    const current = state('saved')
    const before: ProtectionSnapshot = {
      sessionId: 'original',
      cells: [
        { row: 0, column: 0, styleReference: 1, locked: true, hidden: false, entry: null },
        { row: 0, column: 1, styleReference: 2, locked: true, hidden: false, entry: null },
        { row: 0, column: 2, styleReference: 3, locked: true, hidden: false, entry: null },
        { row: 0, column: 3, styleReference: null, locked: true, hidden: false, entry: null },
      ],
    }
    const after: ProtectionSnapshot = {
      sessionId: 'original',
      cells: before.cells.map((cell) => ({
        ...cell,
        verticalAlignment: 'distributed',
        entry: {
          row: cell.row,
          column: cell.column,
          hasValue: false,
          value: null,
          style: { verticalAlignment: 'distributed' },
        },
      })),
    }
    restoreProtectionSnapshot(current, 'sheet-1', before, after)
    expect(toSaveEdits(current.editJournal).map((edit) => edit.styleReference)).toEqual([
      1,
      2,
      3,
      null,
    ])
    expect(toSaveEdits(current.editJournal).every((edit) => !edit.writeValue && !edit.style)).toBe(
      true,
    )
    restoreProtectionSnapshot(current, 'sheet-1', after, before)
    expect(toSaveEdits(current.editJournal)[0]).toMatchObject({
      styleReference: 1,
      style: { verticalAlignment: 'distributed' },
      writeValue: false,
    })
  })

  it('registers the mutation on a fresh runtime before carried Undo and Redo replay', async () => {
    const oldUnit = 'file-oldsha'
    const newUnit = 'file-newsha'
    const makeRuntime = (unitId: string) => {
      const commands = new Map<
        string,
        { handler: (_accessor: unknown, params: unknown) => boolean }
      >()
      const undo = new Map<string, CarriedUndoItem[]>()
      const redo = new Map<string, CarriedUndoItem[]>()
      const service = {
        _undoStacks: undo,
        _redoStacks: redo,
        clearUndoRedo(id: string) {
          undo.set(id, [])
          redo.set(id, [])
        },
        pushUndoRedo(item: CarriedUndoItem) {
          undo.set(item.unitID, [...(undo.get(item.unitID) ?? []), item])
          redo.set(item.unitID, [])
        },
      }
      const registry = {
        registerCommand(command: {
          id: string
          handler: (_accessor: unknown, params: unknown) => boolean
        }) {
          commands.set(command.id, command)
        },
      }
      const runtime = {
        univer: {
          __getInjector: () => ({
            get: (token: unknown) =>
              token === ICommandService ? registry : token === IUndoRedoService ? service : null,
          }),
        },
        univerAPI: {
          getActiveWorkbook: () => ({ getId: () => unitId, getSnapshot: () => ({ styles: {} }) }),
        },
      } as unknown as UniverRuntime
      return { runtime, commands, service }
    }
    const old = makeRuntime(oldUnit)
    const next = makeRuntime(newUnit)
    const univerRef = { current: old.runtime }
    const lazyWorkbookRef = { current: state('original') as LazyWorkbookState }
    const ctx = {
      univerRef,
      lazyWorkbookRef,
      setPendingEdits: () => undefined,
      refreshAlignmentPreview: () => undefined,
    }
    await withCellProtectionUndo(
      ctx,
      {
        sheetId: 'sheet-1',
        range: { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 },
        locked: true,
        hidden: false,
        alignment: {
          horizontal: 'general',
          vertical: 'bottom',
          rotation: 0,
          shrinkToFit: false,
          merged: false,
        },
        cells: [
          {
            row: 0,
            column: 0,
            styleReference: 1,
            locked: true,
            hidden: false,
            horizontalAlignment: 'general',
            verticalAlignment: 'bottom',
            textRotation: 0,
            shrinkToFit: false,
          },
        ],
      },
      { verticalAlignment: 'distributed' },
    )
    const carry = captureUndoCarry(old.runtime, 'newsha')
    expect(carry?.items).toHaveLength(1)
    univerRef.current = next.runtime
    lazyWorkbookRef.current = state('saved') as LazyWorkbookState
    registerCellProtectionUndoRuntime(ctx)
    stashUndoCarry(carry)
    consumePendingUndoCarry(next.runtime, newUnit)
    const registered = next.commands.get('genoffice.mutation.set-cell-protection')
    expect(registered).toBeDefined()
    const item = next.service._undoStacks.get(newUnit)?.[0]
    expect(item).toBeDefined()
    expect(registered!.handler({}, item!.undoMutations[0]!.params)).toBe(true)
    expect(toSaveEdits(lazyWorkbookRef.current.editJournal)[0]).toMatchObject({
      styleReference: 1,
      writeValue: false,
    })
    expect(toSaveEdits(lazyWorkbookRef.current.editJournal)[0]?.style).toBeUndefined()
    expect(registered!.handler({}, item!.redoMutations[0]!.params)).toBe(true)
    expect(toSaveEdits(lazyWorkbookRef.current.editJournal)[0]).toMatchObject({
      styleReference: 1,
      style: { verticalAlignment: 'distributed' },
      writeValue: false,
    })
  })
})
