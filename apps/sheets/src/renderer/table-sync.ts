import type { WorkbookTableAdd } from '../shared/desktop-api'
import { recordTableAdd, removeTableAdd, type EditJournal } from './edit-journal'

/**
 * A table the user or the AI added lives twice: as a Univer table (drawn, on the ⌘Z stack) and as
 * a journal entry (what the save writes). ⌘Z only reached the first, so an undone table stayed in
 * the saved file and blocked the next Format as Table with an "overlaps" error. Undo and redo run
 * Univer's own delete-table / add-table mutations; these follow them in the journal.
 */
export const ADD_TABLE_MUTATION = 'sheet.mutation.add-table'
export const DELETE_TABLE_MUTATION = 'sheet.mutation.delete-table'

/** Univer table id → the journal entry it stands for (kept after a removal, for redo) */
const entries = new Map<string, WorkbookTableAdd>()
/** the ids whose table exists right now: a restyle leaves several ids under one name */
const live = new Set<string>()

/** A new workbook starts with no tracked tables (sheet ids and table names repeat across files). */
export function resetTableSync(): void {
  entries.clear()
  live.clear()
}

export function trackTable(tableId: string, entry: WorkbookTableAdd): void {
  entries.set(tableId, entry)
  live.add(tableId)
}

export function tableIdOf(sheetId: string, name: string): string | undefined {
  const lower = name.toLowerCase()
  for (const id of live) {
    const entry = entries.get(id)
    if (entry && entry.sheetId === sheetId && entry.name.toLowerCase() === lower) return id
  }
  return undefined
}

const inJournal = (journal: EditJournal, entry: WorkbookTableAdd) =>
  journal.tableAdds.find(
    (t) => t.sheetId === entry.sheetId && t.name.toLowerCase() === entry.name.toLowerCase(),
  )

/** True when the journal changed (the caller refreshes the pending-edit count). */
export function syncTableMutation(
  journal: EditJournal,
  mutationId: string,
  params: { tableId?: string } | undefined,
): boolean {
  const tableId = params?.tableId
  if (!tableId) return false
  const known = entries.get(tableId)
  if (!known) return false
  if (mutationId === DELETE_TABLE_MUTATION) {
    live.delete(tableId)
    const current = inJournal(journal, known)
    if (!current) return false
    entries.set(tableId, current) // area / column edits made since the add
    return removeTableAdd(journal, current.sheetId, current.name)
  }
  if (mutationId === ADD_TABLE_MUTATION) {
    live.add(tableId)
    if (inJournal(journal, known)) return false
    recordTableAdd(journal, known)
    return true
  }
  return false
}
