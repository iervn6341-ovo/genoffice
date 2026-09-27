/**
 * Circular references, the way Excel reports them in its status bar
 * ("Circular References: C3"). Univer resolves a cycle quietly (the cells show
 * 0 or an iterated value) and exposes no list of the cells involved. Its
 * getAllDependencyTrees() is not a read — it dispatches a formula-engine
 * mutation, which raced workbook open and sheet focus when run per recalc —
 * so the graph is built here, passively, from the formula text.
 */

export interface DependencyTreeLike {
  readonly treeId: number
  readonly parents?: readonly number[]
  readonly subUnitId: string
  readonly row: number
  readonly column: number
}

export interface CircularCell {
  readonly subUnitId: string
  readonly row: number
  readonly column: number
}

/**
 * Formula cells that sit on a cycle: members of a strongly connected component
 * with more than one cell, or a cell that reads itself. Ordered by sheet, then
 * row-major, so the first entry is the top-left cycle cell — Excel names one cell.
 */
export function circularCells(trees: readonly DependencyTreeLike[]): CircularCell[] {
  const byId = new Map<number, DependencyTreeLike>()
  for (const tree of trees) byId.set(tree.treeId, tree)
  // Tarjan's SCC, iterative so a long dependency chain cannot overflow the stack
  const index = new Map<number, number>()
  const low = new Map<number, number>()
  const onStack = new Set<number>()
  const stack: number[] = []
  const out: CircularCell[] = []
  let next = 0
  for (const root of byId.keys()) {
    if (index.has(root)) continue
    const work: { id: number; edge: number }[] = [{ id: root, edge: 0 }]
    index.set(root, next)
    low.set(root, next++)
    stack.push(root)
    onStack.add(root)
    while (work.length) {
      const frame = work[work.length - 1]!
      const edges = (byId.get(frame.id)?.parents ?? []).filter((p) => byId.has(p))
      if (frame.edge < edges.length) {
        const to = edges[frame.edge++]!
        if (!index.has(to)) {
          index.set(to, next)
          low.set(to, next++)
          stack.push(to)
          onStack.add(to)
          work.push({ id: to, edge: 0 })
        } else if (onStack.has(to)) {
          low.set(frame.id, Math.min(low.get(frame.id)!, index.get(to)!))
        }
        continue
      }
      work.pop()
      const parent = work[work.length - 1]
      if (parent) low.set(parent.id, Math.min(low.get(parent.id)!, low.get(frame.id)!))
      if (low.get(frame.id) !== index.get(frame.id)) continue
      const component: number[] = []
      let member: number
      do {
        member = stack.pop()!
        onStack.delete(member)
        component.push(member)
      } while (member !== frame.id)
      const selfLoop = component.length === 1 && edges.includes(frame.id)
      if (component.length > 1 || selfLoop) {
        for (const id of component) {
          const t = byId.get(id)!
          out.push({ subUnitId: t.subUnitId, row: t.row, column: t.column })
        }
      }
    }
  }
  return out.sort((a, b) =>
    a.subUnitId === b.subUnitId
      ? a.row - b.row || a.column - b.column
      : a.subUnitId < b.subUnitId
        ? -1
        : 1,
  )
}

export interface FormulaCell {
  readonly sheetId: string
  readonly row: number
  readonly column: number
  /** formula text, with or without the leading '=' */
  readonly formula: string
}

interface RefRange {
  readonly sheet: string | null // sheet *name* as written, null = the formula's own sheet
  readonly r1: number
  readonly c1: number
  readonly r2: number
  readonly c2: number
}

const MAX_ROW = 1_048_575
const MAX_COL = 16_383

function colIndex(letters: string): number {
  let n = 0
  for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64)
  return n - 1
}

const SHEET = String.raw`(?:'((?:[^']|'')+)'|([A-Za-z_\u00C0-\uFFFF][\w.\u00C0-\uFFFF]*))!`
// A1 / A1:B2 / A:C / 1:3, optionally sheet-qualified; $ anchors allowed
const REF = new RegExp(
  String.raw`(?<![\w.$'])(?:${SHEET})?(?:` +
    String.raw`\$?([A-Za-z]{1,3})\$?(\d{1,7})(?::\$?([A-Za-z]{1,3})\$?(\d{1,7}))?` +
    String.raw`|\$?([A-Za-z]{1,3}):\$?([A-Za-z]{1,3})` +
    String.raw`|\$?(\d{1,7}):\$?(\d{1,7}))(?![\w(!])`,
  'g',
)

/** Cell / range / whole-column / whole-row references in a formula (string literals skipped) */
export function formulaRefs(formula: string): RefRange[] {
  const text = formula.replace(/"(?:[^"]|"")*"/g, '""')
  const out: RefRange[] = []
  for (const m of text.matchAll(REF)) {
    const sheet = m[1] != null ? m[1].replace(/''/g, "'") : (m[2] ?? null)
    if (m[3] != null) {
      const c1 = colIndex(m[3])
      const r1 = Number(m[4]) - 1
      const c2 = m[5] != null ? colIndex(m[5]) : c1
      const r2 = m[6] != null ? Number(m[6]) - 1 : r1
      out.push({
        sheet,
        r1: Math.min(r1, r2),
        c1: Math.min(c1, c2),
        r2: Math.max(r1, r2),
        c2: Math.max(c1, c2),
      })
    } else if (m[7] != null) {
      const a = colIndex(m[7])
      const b = colIndex(m[8]!)
      out.push({ sheet, r1: 0, c1: Math.min(a, b), r2: MAX_ROW, c2: Math.max(a, b) })
    } else if (m[9] != null) {
      const a = Number(m[9]) - 1
      const b = Number(m[10]) - 1
      out.push({ sheet, r1: Math.min(a, b), c1: 0, r2: Math.max(a, b), c2: MAX_COL })
    }
  }
  return out
}

/**
 * The formula graph as dependency trees: each formula cell reads the formula
 * cells inside the ranges it references. Names, INDIRECT/OFFSET and structured
 * references are not followed (Excel's indicator does see those — rare in cycles).
 */
export function formulaDependencyTrees(
  cells: readonly FormulaCell[],
  sheetIdByName: ReadonlyMap<string, string>,
): DependencyTreeLike[] {
  const bySheet = new Map<string, { id: number; row: number; column: number }[]>()
  cells.forEach((c, id) => {
    const list = bySheet.get(c.sheetId) ?? []
    list.push({ id, row: c.row, column: c.column })
    bySheet.set(c.sheetId, list)
  })
  const lowerNames = new Map([...sheetIdByName].map(([name, id]) => [name.toLowerCase(), id]))
  return cells.map((c, id) => {
    const parents = new Set<number>()
    for (const ref of formulaRefs(c.formula)) {
      const sheetId = ref.sheet == null ? c.sheetId : lowerNames.get(ref.sheet.toLowerCase())
      if (sheetId === undefined) continue
      for (const t of bySheet.get(sheetId) ?? []) {
        if (t.row >= ref.r1 && t.row <= ref.r2 && t.column >= ref.c1 && t.column <= ref.c2)
          parents.add(t.id)
      }
    }
    return { treeId: id, parents: [...parents], subUnitId: c.sheetId, row: c.row, column: c.column }
  })
}
