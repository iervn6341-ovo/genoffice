/**
 * Univer skips borders on merged ranges when only the merge's main cell
 * carries them (the common file shape — Excel writes the style on the
 * top-left cell). `_setStylesCacheForOneCell` passes a `mergeRange` only for
 * covered cells (`isMerged`), never for the main cell (`isMergedMainCell`),
 * and the non-merge branch bails via `intersectMergeRange` — so the main
 * cell's borders never reach the render cache. Passing the merge range for
 * the main cell routes it through `_setMergeBorderProps`. That method stops
 * at the first absent covered cell, so we fill missing perimeter render
 * segments from the anchor style without creating cells or changing XLSX.
 */
import { getColorStyle, type IRange, type IStyleData, type Nullable } from '@univerjs/core'
import { SpreadsheetSkeleton } from '@univerjs/engine-render'

interface MergeInfoLike {
  isMergedMainCell?: boolean
  startRow: number
  startColumn: number
  endRow: number
  endColumn: number
}

interface SkeletonProtoLike {
  worksheet?: { getCellInfoInMergeData(row: number, col: number): MergeInfoLike }
  _stylesCache?: {
    border?: {
      getValue(row: number, col: number): Record<string, unknown> | undefined
      setValue(row: number, col: number, value: Record<string, unknown>): void
    }
  }
  _setBorderStylesCache(
    row: number,
    col: number,
    style: Nullable<IStyleData>,
    options: { mergeRange?: IRange; cacheItem?: unknown } | undefined,
  ): void
}

type BorderSide = 't' | 'b' | 'l' | 'r'

export function fillMissingMergedBorderSegments(
  cache: NonNullable<NonNullable<SkeletonProtoLike['_stylesCache']>['border']>,
  merge: IRange,
  style: IStyleData,
): void {
  // A full-column merge must not allocate a million render-cache entries.
  if (2 * (merge.endColumn - merge.startColumn + merge.endRow - merge.startRow + 2) > 10_000) return
  const segments: readonly { side: BorderSide; row: number; col: number }[] = [
    ...Array.from({ length: merge.endColumn - merge.startColumn + 1 }, (_, index) => ({
      side: 't' as const,
      row: merge.startRow,
      col: merge.startColumn + index,
    })),
    ...Array.from({ length: merge.endColumn - merge.startColumn + 1 }, (_, index) => ({
      side: 'b' as const,
      row: merge.endRow,
      col: merge.startColumn + index,
    })),
    ...Array.from({ length: merge.endRow - merge.startRow + 1 }, (_, index) => ({
      side: 'l' as const,
      row: merge.startRow + index,
      col: merge.startColumn,
    })),
    ...Array.from({ length: merge.endRow - merge.startRow + 1 }, (_, index) => ({
      side: 'r' as const,
      row: merge.startRow + index,
      col: merge.endColumn,
    })),
  ]
  for (const { side, row, col } of segments) {
    const props = style.bd?.[side]
    if (!props) continue
    const existing = cache.getValue(row, col)
    if (existing?.[side]) continue
    cache.setValue(row, col, {
      ...existing,
      [side]: {
        type: side,
        style: props.s,
        color: getColorStyle(props.cl) || 'rgb(0,0,0)',
      },
    })
  }
}

let installed = false

export function installMergeBorderFix(): void {
  if (installed) return
  installed = true
  const proto = SpreadsheetSkeleton.prototype as unknown as SkeletonProtoLike
  const original = proto._setBorderStylesCache
  if (typeof original !== 'function') return
  proto._setBorderStylesCache = function (
    this: SkeletonProtoLike,
    row: number,
    col: number,
    style: Nullable<IStyleData>,
    options: { mergeRange?: IRange; cacheItem?: unknown } | undefined,
  ): void {
    let mainMergeRange: IRange | undefined
    if (style?.bd) {
      const info = this.worksheet?.getCellInfoInMergeData(row, col)
      if (
        info?.isMergedMainCell &&
        (info.endRow > info.startRow || info.endColumn > info.startColumn)
      ) {
        mainMergeRange = options?.mergeRange ?? {
          startRow: info.startRow,
          startColumn: info.startColumn,
          endRow: info.endRow,
          endColumn: info.endColumn,
        }
        if (!options?.mergeRange) options = { ...options, mergeRange: mainMergeRange }
      }
    }
    original.call(this, row, col, style, options)
    const cache = this._stylesCache?.border
    if (mainMergeRange && cache && style)
      fillMissingMergedBorderSegments(cache, mainMergeRange, style)
  }
}
