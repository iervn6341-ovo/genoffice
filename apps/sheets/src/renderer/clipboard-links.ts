/**
 * Hyperlinks of a block pasted inside one workbook window. Univer's paste moves values and
 * formats; the links are kept in the edit journal, so the paste has to copy them (and, after a
 * Cut, take them off the source) itself.
 */
export interface LocalLinkPaste {
  readonly sourceRows: readonly number[]
  readonly sourceColumns: readonly number[]
  readonly targetRows: readonly number[]
  readonly targetColumns: readonly number[]
  readonly cut: boolean
  readonly sourceSheetId: string
  readonly targetSheetId: string
  readonly sourceTarget: (row: number, column: number) => string | null
  readonly targetLink: (row: number, column: number) => string | null
}

export interface LinkEdit {
  readonly row: number
  readonly column: number
  readonly target: string | null
}

/// The link edits per sheet: every pasted cell takes the link of the source cell it came from
/// (the block repeats over a bigger selection); after a Cut the source cells lose theirs.
export function linkEditsForLocalPaste(input: LocalLinkPaste): Map<string, LinkEdit[]> {
  const edits = new Map<string, Map<string, LinkEdit>>()
  const put = (sheetId: string, edit: LinkEdit): void => {
    const sheet = edits.get(sheetId) ?? new Map<string, LinkEdit>()
    edits.set(sheetId, sheet)
    sheet.set(`${edit.row}:${edit.column}`, edit)
  }
  const rowCount = input.sourceRows.length
  const columnCount = input.sourceColumns.length
  if (rowCount === 0 || columnCount === 0) return new Map()
  const moved: { row: number; column: number }[] = []
  input.targetRows.forEach((row, i) => {
    const sourceRow = input.sourceRows[i % rowCount]!
    input.targetColumns.forEach((column, j) => {
      const sourceColumn = input.sourceColumns[j % columnCount]!
      const wanted = input.sourceTarget(sourceRow, sourceColumn)
      const have = input.targetLink(row, column)
      if (wanted !== have) put(input.targetSheetId, { row, column, target: wanted })
      if (input.cut && wanted !== null) moved.push({ row: sourceRow, column: sourceColumn })
    })
  })
  if (input.cut) {
    const inTarget = new Set([...(edits.get(input.targetSheetId)?.keys() ?? [])])
    for (const { row, column } of moved) {
      // a source cell that the paste itself landed on keeps the link it was just given
      if (input.sourceSheetId === input.targetSheetId && inTarget.has(`${row}:${column}`)) continue
      if (
        input.targetRows.includes(row) &&
        input.targetColumns.includes(column) &&
        input.sourceSheetId === input.targetSheetId
      )
        continue
      put(input.sourceSheetId, { row, column, target: null })
    }
  }
  return new Map([...edits].map(([sheetId, map]) => [sheetId, [...map.values()]]))
}
