import type { EditJournal, JournalEntry } from './edit-journal'
import { restoreJournalCellSnapshot } from './edit-journal'

interface Checkpoint {
  signature: string
  entry: JournalEntry | null
}
const checkpoints = new WeakMap<EditJournal, Map<string, Checkpoint>>()

/** Cached formula results and absent/null fields are not authored differences. */
export function cellPersistenceSignature(cell: unknown): string {
  const normalize = (value: unknown): unknown => {
    if (value === null || value === undefined) return undefined
    if (Array.isArray(value)) return value.map(normalize)
    if (typeof value !== 'object') return value
    const object = value as Record<string, unknown>
    const result: Record<string, unknown> = {}
    for (const key of Object.keys(object).sort()) {
      if (typeof object.f === 'string' && object.f && ['v', 't', 'si'].includes(key)) continue
      const item = normalize(object[key])
      if (
        item !== undefined &&
        !(
          typeof item === 'object' &&
          !Array.isArray(item) &&
          Object.keys(item as object).length === 0
        )
      )
        result[key] = item
    }
    return result
  }
  return JSON.stringify(normalize(cell) ?? {})
}

/** Captured before the first authored mutation; a save creates a new journal/checkpoint. */
export function captureCellSaveCheckpoint(
  journal: EditJournal,
  sheetId: string,
  row: number,
  column: number,
  cell: unknown,
): void {
  let map = checkpoints.get(journal)
  if (!map) {
    map = new Map()
    checkpoints.set(journal, map)
  }
  const key = `${sheetId}:${row}:${column}`
  if (map.has(key)) return
  const entry = journal.cells.get(sheetId)?.get(`${row}:${column}`) ?? null
  map.set(key, {
    signature: cellPersistenceSignature(cell),
    entry: entry ? structuredClone(entry) : null,
  })
}

export function reconcileCellSaveCheckpoint(
  journal: EditJournal,
  sheetId: string,
  row: number,
  column: number,
  cell: unknown,
): boolean {
  const saved = checkpoints.get(journal)?.get(`${sheetId}:${row}:${column}`)
  if (!saved || saved.signature !== cellPersistenceSignature(cell)) return false
  restoreJournalCellSnapshot(journal, sheetId, row, column, saved.entry)
  return true
}

/** Row/column moves change coordinate identity; never reuse an unrelated cell's checkpoint. */
export function invalidateCellSaveCheckpoints(journal: EditJournal): void {
  checkpoints.delete(journal)
}
