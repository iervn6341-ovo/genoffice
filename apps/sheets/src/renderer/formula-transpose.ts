/// Paste Special → Transpose rewrites the cell references of a pasted formula. Excel 16 (probed
/// through COM): `=A2+B2` in C2, copied with A1:C2 and pasted transposed at E1, lands in F3 as
/// `=F1+F2`; `=$A$1*2` copied with A1:B1 and pasted transposed at D1 becomes `=$D$1*2`. So:
///   - a relative reference keeps its offset from the formula cell with row and column swapped;
///   - an absolute reference to a cell inside the copied block follows that cell to its new spot;
///   - an absolute reference outside the block stays; mixed ($A1 / A$1) references stay as written.

export interface TransposeBlock {
  /// absolute row of each copied row, in order (a filtered-out row is not in the list)
  readonly rows: readonly number[]
  readonly columns: readonly number[]
  /// where the transposed block lands (its top-left cell)
  readonly anchorRow: number
  readonly anchorColumn: number
  /// the sheet the block was copied from: an absolute reference to ANOTHER sheet does not point into
  /// the block and stays as written (omitted = every reference is taken to be on that sheet)
  readonly sheetName?: string
}

const REFERENCE =
  /(?<![A-Za-z0-9_.$'"!])((?:'(?:[^']|'')+'|[A-Za-z_][A-Za-z0-9_.]*)!)?(\$?)([A-Z]{1,3})(\$?)(\d{1,7})(?::(\$?)([A-Z]{1,3})(\$?)(\d{1,7}))?(?![A-Za-z0-9_(.!])/gi

export function columnIndex(label: string): number {
  let index = 0
  for (const char of label.toUpperCase()) index = index * 26 + (char.charCodeAt(0) - 64)
  return index - 1
}

export function columnName(index: number): string {
  let name = ''
  let rest = index + 1
  while (rest > 0) {
    const part = (rest - 1) % 26
    name = String.fromCharCode(65 + part) + name
    rest = Math.floor((rest - 1) / 26)
  }
  return name
}

interface Corner {
  readonly colAbs: boolean
  readonly col: number
  readonly rowAbs: boolean
  readonly row: number
}

/// the corner moved to where the transposed paste puts it; null = leave it as written
function moveCorner(
  corner: Corner,
  source: { row: number; column: number },
  dest: { row: number; column: number },
  block: TransposeBlock,
  prefix: string | undefined,
): { row: number; col: number } | null | 'ref-error' {
  if (!corner.colAbs && !corner.rowAbs) {
    const row = dest.row + (corner.col - source.column)
    const col = dest.column + (corner.row - source.row)
    return row < 0 || col < 0 ? 'ref-error' : { row, col }
  }
  if (corner.colAbs && corner.rowAbs) {
    if (
      prefix !== undefined &&
      block.sheetName !== undefined &&
      sheetOf(prefix) !== block.sheetName
    ) {
      return null
    }
    const i = block.rows.indexOf(corner.row)
    const j = block.columns.indexOf(corner.col)
    if (i < 0 || j < 0) return null
    return { row: block.anchorRow + j, col: block.anchorColumn + i }
  }
  return null
}

/** "Sheet2!" / "'My Sheet'!" → the sheet's name */
function sheetOf(prefix: string): string {
  const name = prefix.replace(/!$/, '')
  return name.startsWith("'") ? name.slice(1, -1).replace(/''/g, "'") : name
}

/**
 * The formula cut into parts that are never references — "text", and [structured references]
 * such as Sales[[#This Row],[Q1]] (nesting counted) — and the parts between them, which are.
 */
function splitProtected(formula: string): { text: string; protectedPart: boolean }[] {
  const parts: { text: string; protectedPart: boolean }[] = []
  let plain = ''
  const flush = (): void => {
    if (plain) parts.push({ text: plain, protectedPart: false })
    plain = ''
  }
  for (let i = 0; i < formula.length; i += 1) {
    const char = formula[i]!
    if (char === '"') {
      let j = i + 1
      while (j < formula.length) {
        if (formula[j] === '"') {
          if (formula[j + 1] === '"') j += 2
          else break
        } else j += 1
      }
      flush()
      parts.push({ text: formula.slice(i, j + 1), protectedPart: true })
      i = j
    } else if (char === '[') {
      let depth = 0
      let j = i
      for (; j < formula.length; j += 1) {
        if (formula[j] === '[') depth += 1
        else if (formula[j] === ']') {
          depth -= 1
          if (depth === 0) break
        }
      }
      flush()
      parts.push({ text: formula.slice(i, j + 1), protectedPart: true })
      i = j
    } else {
      plain += char
    }
  }
  flush()
  return parts
}

/**
 * `formula` as it reads after its cell moved from `source` to `dest` in a transposed paste of
 * `block`. A reference that would fall off the sheet becomes `#REF!`, like a moved reference does.
 */
export function transposeFormula(
  formula: string,
  source: { row: number; column: number },
  dest: { row: number; column: number },
  block: TransposeBlock,
): string {
  // text inside "…" and structured references are never cell references
  return splitProtected(formula)
    .map(({ text: part, protectedPart }) =>
      protectedPart
        ? part
        : part.replace(
            REFERENCE,
            (
              match: string,
              sheet: string | undefined,
              c1: string,
              col1: string,
              r1: string,
              row1: string,
              c2: string | undefined,
              col2: string | undefined,
              r2: string | undefined,
              row2: string | undefined,
            ) => {
              const first: Corner = {
                colAbs: c1 === '$',
                col: columnIndex(col1),
                rowAbs: r1 === '$',
                row: Number(row1) - 1,
              }
              const last: Corner | null =
                col2 === undefined || row2 === undefined
                  ? null
                  : {
                      colAbs: c2 === '$',
                      col: columnIndex(col2),
                      rowAbs: r2 === '$',
                      row: Number(row2) - 1,
                    }
              const movedFirst = moveCorner(first, source, dest, block, sheet)
              const movedLast = last ? moveCorner(last, source, dest, block, sheet) : undefined
              if (movedFirst === 'ref-error' || movedLast === 'ref-error') return '#REF!'
              // a corner left as written leaves the whole reference as written
              if (movedFirst === null || movedLast === null) return match
              const render = (corner: Corner, at: { row: number; col: number }): string =>
                `${corner.colAbs ? '$' : ''}${columnName(at.col)}${corner.rowAbs ? '$' : ''}${at.row + 1}`
              const prefix = sheet ?? ''
              if (!last || movedLast === undefined) return `${prefix}${render(first, movedFirst)}`
              // the transposed range keeps its corners in order
              const top = Math.min(movedFirst.row, movedLast.row)
              const bottom = Math.max(movedFirst.row, movedLast.row)
              const left = Math.min(movedFirst.col, movedLast.col)
              const right = Math.max(movedFirst.col, movedLast.col)
              return `${prefix}${render(first, { row: top, col: left })}:${render(last, { row: bottom, col: right })}`
            },
          ),
    )
    .join('')
}
