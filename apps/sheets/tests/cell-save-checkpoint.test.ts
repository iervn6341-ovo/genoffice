import { describe, expect, it } from 'vitest'
import {
  captureCellSaveCheckpoint,
  cellPersistenceSignature,
  reconcileCellSaveCheckpoint,
} from '../src/renderer/cell-save-checkpoint'
import { createEditJournal, journalSize, recordSetRangeValues } from '../src/renderer/edit-journal'

describe('cell save checkpoints', () => {
  it('removes reverted entries on Undo and records them again on Redo', () => {
    const journal = createEditJournal()
    const original = { v: 'Original', s: { bl: 1 } }
    captureCellSaveCheckpoint(journal, 's', 1, 1, original)
    recordSetRangeValues(journal, 's', { 1: { 1: { v: 'Changed' } } })
    expect(reconcileCellSaveCheckpoint(journal, 's', 1, 1, { ...original, v: 'Changed' })).toBe(
      false,
    )
    expect(journalSize(journal)).toBe(1)
    recordSetRangeValues(journal, 's', { 1: { 1: original } })
    expect(reconcileCellSaveCheckpoint(journal, 's', 1, 1, { ...original, p: null })).toBe(true)
    expect(journalSize(journal)).toBe(0)
    recordSetRangeValues(journal, 's', { 1: { 1: { v: 'Changed' } } })
    expect(reconcileCellSaveCheckpoint(journal, 's', 1, 1, { ...original, v: 'Changed' })).toBe(
      false,
    )
    expect(journalSize(journal)).toBe(1)
  })

  it('rebases on save and retains actual style differences and pre-existing edits', () => {
    const saved = createEditJournal()
    captureCellSaveCheckpoint(saved, 's', 1, 1, { v: 'Changed' })
    recordSetRangeValues(saved, 's', { 1: { 1: { v: 'Original' } } })
    expect(reconcileCellSaveCheckpoint(saved, 's', 1, 1, { v: 'Original' })).toBe(false)
    expect(reconcileCellSaveCheckpoint(saved, 's', 1, 1, { v: 'Changed', s: { bl: 1 } })).toBe(
      false,
    )
    expect(journalSize(saved)).toBe(1)
    expect(reconcileCellSaveCheckpoint(saved, 's', 1, 1, { v: 'Changed' })).toBe(true)
    expect(journalSize(saved)).toBe(0)
    recordSetRangeValues(saved, 's', { 1: { 1: { v: 'Pre-existing edit' } } })
    captureCellSaveCheckpoint(saved, 's', 2, 1, { v: 12 })
    recordSetRangeValues(saved, 's', { 2: { 1: { v: 20 } } })
    reconcileCellSaveCheckpoint(saved, 's', 2, 1, { v: 12 })
    expect(journalSize(saved)).toBe(1)
  })

  it('compares formulas independently of cached results and keeps format changes', () => {
    expect(cellPersistenceSignature({ f: '=A1+1', v: 2, t: 2, p: null })).toBe(
      cellPersistenceSignature({ f: '=A1+1', v: 3 }),
    )
    expect(cellPersistenceSignature({ f: '=A1+1' })).not.toBe(
      cellPersistenceSignature({ f: '=A1+2' }),
    )
    expect(cellPersistenceSignature({ v: 1, s: { bl: 1 } })).not.toBe(
      cellPersistenceSignature({ v: 1 }),
    )
  })
})
