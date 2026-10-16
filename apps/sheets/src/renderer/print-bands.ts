/**
 * Splits a print job that is too large to lay out in one piece into row bands. Each band is read
 * from the file, laid out and printed on its own (and released afterwards), so a sheet with
 * hundreds of thousands of cells prints without ever being fully loaded into the grid.
 */

/** Cells per band: small enough that rendering one band stays light (a 40k-cell band cost ~2.7GB). */
export const BAND_MAX_CELLS = 10_000
/** Cells a whole print job may cover; beyond it the user is asked to narrow the range. */
export const PRINT_JOB_MAX_CELLS = 5_000_000

export interface PrintArea {
  readonly startRow: number
  readonly startColumn: number
  readonly endRow: number
  readonly endColumn: number
}

export function areaCells(area: PrintArea): number {
  return (area.endRow - area.startRow + 1) * (area.endColumn - area.startColumn + 1)
}

export function totalCells(areas: readonly PrintArea[]): number {
  return areas.reduce((sum, area) => sum + Math.max(0, areaCells(area)), 0)
}

/** Row bands of one area, each at most `maxCells` (a band always holds at least one row). */
export function bandsOf(area: PrintArea, maxCells: number = BAND_MAX_CELLS): PrintArea[] {
  const columns = area.endColumn - area.startColumn + 1
  if (columns < 1 || area.endRow < area.startRow) return []
  const rowsPerBand = Math.max(1, Math.floor(maxCells / columns))
  const bands: PrintArea[] = []
  for (let row = area.startRow; row <= area.endRow; row += rowsPerBand) {
    bands.push({ ...area, startRow: row, endRow: Math.min(area.endRow, row + rowsPerBand - 1) })
  }
  return bands
}

/**
 * The print units of one sheet: a single unit holding every area while they fit one payload,
 * otherwise one unit per band (areas banded in order).
 */
export function unitsForAreas(
  areas: readonly PrintArea[],
  maxCells: number = BAND_MAX_CELLS,
): { readonly banded: boolean; readonly units: readonly (readonly PrintArea[])[] } {
  if (totalCells(areas) <= maxCells * 2) return { banded: false, units: [areas] }
  return { banded: true, units: areas.flatMap((area) => bandsOf(area, maxCells)).map((b) => [b]) }
}

export function exceedsJobLimit(areas: readonly PrintArea[]): boolean {
  return totalCells(areas) > PRINT_JOB_MAX_CELLS
}

/** Progress label source: how far a band-by-band job is. */
export function bandProgress(done: number, total: number): { done: number; total: number } {
  return { done: Math.min(Math.max(done, 0), total), total }
}
