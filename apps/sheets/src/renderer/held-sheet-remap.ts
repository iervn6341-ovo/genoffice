/**
 * A two-phase save writes the structure (new sheets, row/column shifts) first and the held
 * PivotTables / tables against the REOPENED file second. A pivot or table on a sheet created in the
 * same session names that sheet by the renderer's temporary id, which the reopened session does not
 * know (the gateway gives the new sheet its real id when it writes it). The sheet itself is
 * identifiable by its final name, so before the second phase the held items are re-addressed by
 * name onto the real ids of the reopened file.
 */

export interface SheetOpLike {
  readonly kind: string
  readonly sheetId?: string
  readonly name?: string
  readonly newName?: string
}

export interface SavedSheetLike {
  readonly id: string
  readonly name: string
}

/** temporary id -> the name the sheet carries once the first phase has written it */
export function addedSheetFinalNames(ops: readonly SheetOpLike[]): Map<string, string> {
  const names = new Map<string, string>()
  for (const op of ops) {
    if ((op.kind === 'add-sheet' || op.kind === 'duplicate-sheet') && op.sheetId && op.name)
      names.set(op.sheetId, op.name)
  }
  for (const op of ops) {
    if (op.kind === 'rename-sheet' && op.sheetId && op.newName && names.has(op.sheetId))
      names.set(op.sheetId, op.newName)
  }
  return names
}

/** Real id of the saved sheet called `name` (exact first, then case-insensitive as Excel treats names). */
function savedIdByName(saved: readonly SavedSheetLike[], name: string): string | undefined {
  return (
    saved.find((sheet) => sheet.name === name)?.id ??
    saved.find((sheet) => sheet.name.toLowerCase() === name.toLowerCase())?.id
  )
}

/**
 * Re-address `sheetId` (and `sourceSheetId` when present) of held items that point at sheets added
 * this session. Ids of sheets the file already had are left alone (they are stable across saves).
 * Throws when an added sheet cannot be found in the saved file, so the caller can report the
 * second phase as failed instead of writing to the wrong sheet.
 */
export function remapHeldToSavedSheets<T extends { sheetId: string; sourceSheetId?: string }>(
  items: readonly T[],
  finalNames: ReadonlyMap<string, string>,
  saved: readonly SavedSheetLike[],
): T[] {
  const resolve = (sheetId: string): string => {
    const finalName = finalNames.get(sheetId)
    if (finalName === undefined) return sheetId
    const id = savedIdByName(saved, finalName)
    if (id === undefined) throw new Error(`The new sheet "${finalName}" is not in the saved file.`)
    return id
  }
  return items.map((item) => ({
    ...item,
    sheetId: resolve(item.sheetId),
    ...(item.sourceSheetId === undefined ? {} : { sourceSheetId: resolve(item.sourceSheetId) }),
  }))
}
