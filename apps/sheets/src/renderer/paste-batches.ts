/// One set-range-values call writing more than about 40,000 cells loses its first cell inside
/// Univer's chunked execution of large writes; batches of whole rows below that size do not.
export const WRITE_BATCH_CELLS = 20_000

export function rowBatches<T extends Record<number, Record<number, unknown>>>(matrix: T): T[] {
  const batches: T[] = []
  let current: Record<number, Record<number, unknown>> = {}
  let count = 0
  for (const key of Object.keys(matrix)
    .map(Number)
    .sort((a, b) => a - b)) {
    const row = matrix[key]!
    const size = Object.keys(row).length
    if (count > 0 && count + size > WRITE_BATCH_CELLS) {
      batches.push(current as T)
      current = {}
      count = 0
    }
    current[key] = row
    count += size
  }
  if (count > 0) batches.push(current as T)
  return batches
}
