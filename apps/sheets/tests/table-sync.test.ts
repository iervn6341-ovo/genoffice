import { beforeEach, describe, expect, it } from 'vitest'

import type { WorkbookTableAdd } from '../src/shared/desktop-api'
import { createEditJournal, recordTableAdd } from '../src/renderer/edit-journal'
import {
  ADD_TABLE_MUTATION,
  DELETE_TABLE_MUTATION,
  resetTableSync,
  syncTableMutation,
  tableIdOf,
  trackTable,
} from '../src/renderer/table-sync'

const entry = (name: string, style: string): WorkbookTableAdd => ({
  sheetId: 's1',
  area: { startRow: 0, startColumn: 0, endRow: 5, endColumn: 3 },
  name,
  columnNames: ['a', 'b', 'c', 'd'],
  style,
  bandedRows: true,
})

describe('table journal follows Univer undo / redo of a table', () => {
  beforeEach(resetTableSync)

  it('drops the journal entry when ⌘Z deletes the table and restores it on redo', () => {
    const journal = createEditJournal()
    const table = entry('Table1', 'TableStyleLight1')
    recordTableAdd(journal, table)
    trackTable('t1', table)

    expect(syncTableMutation(journal, DELETE_TABLE_MUTATION, { tableId: 't1' })).toBe(true)
    expect(journal.tableAdds).toEqual([])
    expect(syncTableMutation(journal, ADD_TABLE_MUTATION, { tableId: 't1' })).toBe(true)
    expect(journal.tableAdds).toEqual([table])
    // a second add of the same table changes nothing
    expect(syncTableMutation(journal, ADD_TABLE_MUTATION, { tableId: 't1' })).toBe(false)
  })

  it('walks a restyle (same name, new table id) back and forth', () => {
    const journal = createEditJournal()
    const gray = entry('Table1', 'TableStyleLight1')
    const orange = entry('Table1', 'TableStyleMedium3')
    recordTableAdd(journal, gray)
    trackTable('t1', gray)
    expect(tableIdOf('s1', 'Table1')).toBe('t1')

    // the restyle: delete t1, add t2
    syncTableMutation(journal, DELETE_TABLE_MUTATION, { tableId: 't1' })
    recordTableAdd(journal, orange)
    trackTable('t2', orange)
    expect(tableIdOf('s1', 'table1')).toBe('t2')

    // ⌘Z: delete t2, add t1 back
    syncTableMutation(journal, DELETE_TABLE_MUTATION, { tableId: 't2' })
    syncTableMutation(journal, ADD_TABLE_MUTATION, { tableId: 't1' })
    expect(journal.tableAdds.map((t) => t.style)).toEqual(['TableStyleLight1'])
    expect(tableIdOf('s1', 'Table1')).toBe('t1')

    // ⇧⌘Z: delete t1, add t2
    syncTableMutation(journal, DELETE_TABLE_MUTATION, { tableId: 't1' })
    syncTableMutation(journal, ADD_TABLE_MUTATION, { tableId: 't2' })
    expect(journal.tableAdds.map((t) => t.style)).toEqual(['TableStyleMedium3'])
  })

  it('keeps the area a table grew to while it was alive', () => {
    const journal = createEditJournal()
    const table = entry('Table1', 'TableStyleLight1')
    recordTableAdd(journal, table)
    trackTable('t1', table)
    journal.tableAdds[0] = { ...table, area: { ...table.area, endRow: 9 } }
    syncTableMutation(journal, DELETE_TABLE_MUTATION, { tableId: 't1' })
    syncTableMutation(journal, ADD_TABLE_MUTATION, { tableId: 't1' })
    expect(journal.tableAdds[0]?.area.endRow).toBe(9)
  })

  it('ignores tables it never tracked (file tables, other workbooks)', () => {
    const journal = createEditJournal()
    const table = entry('Table1', 'TableStyleLight1')
    recordTableAdd(journal, table)
    expect(syncTableMutation(journal, DELETE_TABLE_MUTATION, { tableId: 'file-table-1' })).toBe(
      false,
    )
    expect(syncTableMutation(journal, DELETE_TABLE_MUTATION, undefined)).toBe(false)
    expect(journal.tableAdds).toEqual([table])
  })
})
