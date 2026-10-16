import type { Node as PmNode } from '@tiptap/pm/model'
import type { Command } from '@tiptap/pm/state'
import { TableMap } from '@tiptap/pm/tables'
import { selectedTable } from './table-properties'
import { numberOf } from './table-tools'

/**
 * Table Layout > Formula: Word's `{ =SUM(ABOVE) \# "0.00" }` field. The result text carries an
 * `instrField` mark (the same native complex-field representation every other field uses), so the
 * field survives save/reopen and Word can update it with F9. Supported: + - * / ( ), cell references
 * (A1, B2:C4), ABOVE / BELOW / LEFT / RIGHT, and SUM AVERAGE COUNT MAX MIN PRODUCT ABS INT ROUND MOD SIGN.
 */

export class FormulaError extends Error {}

interface GridCell {
  text: string
}

export interface FormulaGrid {
  rows: number
  columns: number
  cell(row: number, column: number): GridCell | null
}

type Direction = 'ABOVE' | 'BELOW' | 'LEFT' | 'RIGHT'

type Token =
  | { kind: 'number'; value: number }
  | { kind: 'name'; value: string }
  | { kind: 'op'; value: string }

function tokenize(source: string): Token[] {
  const tokens: Token[] = []
  const re = /\s*(?:(\d+\.?\d*|\.\d+)|([A-Za-z_][A-Za-z0-9_]*)|([-+*/(),;:%^]))/y
  let index = 0
  const trimmed = source.replace(/\s+$/, '')
  while (index < trimmed.length) {
    re.lastIndex = index
    const match = re.exec(trimmed)
    if (!match) throw new FormulaError('Syntax Error')
    if (match[1] !== undefined) tokens.push({ kind: 'number', value: Number(match[1]) })
    else if (match[2] !== undefined) tokens.push({ kind: 'name', value: match[2].toUpperCase() })
    else tokens.push({ kind: 'op', value: match[3] })
    index = re.lastIndex
  }
  return tokens
}

const columnIndex = (letters: string): number =>
  [...letters].reduce((sum, letter) => sum * 26 + (letter.charCodeAt(0) - 64), 0) - 1

const CELL_REF = /^([A-Z]{1,3})(\d{1,4})$/

function valuesOf(cells: Array<GridCell | null>): number[] {
  const out: number[] = []
  for (const cell of cells) {
    const value = cell ? numberOf(cell.text) : null
    if (value !== null) out.push(value)
  }
  return out
}

/** Cells in a direction from (row, column): stops at the first empty cell, ignores text cells (Word). */
function directional(
  grid: FormulaGrid,
  row: number,
  column: number,
  direction: Direction,
): GridCell[] {
  const step = {
    ABOVE: [-1, 0],
    BELOW: [1, 0],
    LEFT: [0, -1],
    RIGHT: [0, 1],
  }[direction]
  const out: GridCell[] = []
  let r = row + step[0]
  let c = column + step[1]
  while (r >= 0 && c >= 0 && r < grid.rows && c < grid.columns) {
    const cell = grid.cell(r, c)
    if (!cell || cell.text.trim() === '') break
    out.push(cell)
    r += step[0]
    c += step[1]
  }
  return out
}

/** Value list of one function argument. */
type Arg = number[]

class Parser {
  private index = 0
  constructor(
    private readonly tokens: Token[],
    private readonly grid: FormulaGrid,
    private readonly row: number,
    private readonly column: number,
  ) {}

  parse(): number {
    const value = this.expression()
    if (this.index < this.tokens.length) throw new FormulaError('Syntax Error')
    return value
  }

  private peek(): Token | undefined {
    return this.tokens[this.index]
  }

  private take(): Token {
    const token = this.tokens[this.index++]
    if (!token) throw new FormulaError('Syntax Error')
    return token
  }

  private isOp(value: string): boolean {
    const token = this.peek()
    return token?.kind === 'op' && token.value === value
  }

  private expression(): number {
    let left = this.term()
    while (this.isOp('+') || this.isOp('-')) {
      const op = (this.take() as { value: string }).value
      const right = this.term()
      left = op === '+' ? left + right : left - right
    }
    return left
  }

  private term(): number {
    let left = this.unary()
    while (this.isOp('*') || this.isOp('/')) {
      const op = (this.take() as { value: string }).value
      const right = this.unary()
      if (op === '/' && right === 0) throw new FormulaError('Division by zero')
      left = op === '*' ? left * right : left / right
    }
    return left
  }

  private unary(): number {
    if (this.isOp('-')) {
      this.take()
      return -this.unary()
    }
    if (this.isOp('+')) {
      this.take()
      return this.unary()
    }
    let value = this.primary()
    while (this.isOp('^')) {
      this.take()
      value = value ** this.unary()
    }
    if (this.isOp('%')) {
      this.take()
      value /= 100
    }
    return value
  }

  private primary(): number {
    const token = this.take()
    if (token.kind === 'number') return token.value
    if (token.kind === 'op') {
      if (token.value !== '(') throw new FormulaError('Syntax Error')
      const value = this.expression()
      if (!this.isOp(')')) throw new FormulaError('Syntax Error')
      this.take()
      return value
    }
    if (this.isOp('(')) return this.call(token.value)
    // bare cell reference used as a scalar
    const values = this.reference(token.value)
    if (values.length !== 1) throw new FormulaError('Syntax Error')
    return values[0]
  }

  private reference(name: string): number[] {
    if (name === 'ABOVE' || name === 'BELOW' || name === 'LEFT' || name === 'RIGHT')
      return valuesOf(directional(this.grid, this.row, this.column, name))
    const start = CELL_REF.exec(name)
    if (!start) throw new FormulaError(`Undefined bookmark, ${name}`)
    const r1 = Number(start[2]) - 1
    const c1 = columnIndex(start[1])
    let r2 = r1
    let c2 = c1
    if (this.isOp(':')) {
      this.take()
      const endToken = this.take()
      const end = endToken.kind === 'name' ? CELL_REF.exec(endToken.value) : null
      if (!end) throw new FormulaError('Syntax Error')
      r2 = Number(end[2]) - 1
      c2 = columnIndex(end[1])
    }
    const cells: Array<GridCell | null> = []
    for (let r = Math.min(r1, r2); r <= Math.max(r1, r2); r++)
      for (let c = Math.min(c1, c2); c <= Math.max(c1, c2); c++) cells.push(this.grid.cell(r, c))
    return cells.length === 1 && cells[0] === null
      ? (() => {
          throw new FormulaError('Syntax Error')
        })()
      : valuesOf(cells)
  }

  private args(): Arg[] {
    this.take() // (
    const args: Arg[] = []
    if (this.isOp(')')) {
      this.take()
      return args
    }
    for (;;) {
      const token = this.peek()
      if (
        token?.kind === 'name' &&
        (CELL_REF.test(token.value) || /^(ABOVE|BELOW|LEFT|RIGHT)$/.test(token.value))
      ) {
        const after = this.tokens[this.index + 1]
        const isRange = after?.kind === 'op' && after.value === ':'
        const isRef =
          after?.kind === 'op' &&
          (after.value === ',' || after.value === ';' || after.value === ')')
        if (isRange || isRef || /^(ABOVE|BELOW|LEFT|RIGHT)$/.test(token.value)) {
          this.take()
          args.push(this.reference(token.value))
        } else args.push([this.expression()])
      } else args.push([this.expression()])
      if (this.isOp(',') || this.isOp(';')) {
        this.take()
        continue
      }
      if (this.isOp(')')) {
        this.take()
        return args
      }
      throw new FormulaError('Syntax Error')
    }
  }

  private call(name: string): number {
    const args = this.args()
    const flat = args.flat()
    const scalar = (index: number): number => {
      const value = args[index]?.[0]
      if (value === undefined) throw new FormulaError('Syntax Error')
      return value
    }
    switch (name) {
      case 'SUM':
        return flat.reduce((sum, value) => sum + value, 0)
      case 'AVERAGE':
        if (flat.length === 0) throw new FormulaError('Division by zero')
        return flat.reduce((sum, value) => sum + value, 0) / flat.length
      case 'COUNT':
        return flat.length
      case 'MAX':
        return flat.length ? Math.max(...flat) : 0
      case 'MIN':
        return flat.length ? Math.min(...flat) : 0
      case 'PRODUCT':
        return flat.reduce((product, value) => product * value, 1)
      case 'ABS':
        return Math.abs(scalar(0))
      case 'INT':
        return Math.floor(scalar(0))
      case 'SIGN':
        return Math.sign(scalar(0))
      case 'MOD': {
        const divisor = scalar(1)
        if (divisor === 0) throw new FormulaError('Division by zero')
        return scalar(0) % divisor
      }
      case 'ROUND': {
        const digits = args.length > 1 ? scalar(1) : 0
        const factor = 10 ** digits
        return Math.round(scalar(0) * factor) / factor
      }
      default:
        throw new FormulaError(`Unknown function, ${name}`)
    }
  }
}

/** Evaluate `=SUM(ABOVE)` style text (leading `=` optional) at a grid position. */
export function evaluateFormula(
  formula: string,
  grid: FormulaGrid,
  row: number,
  column: number,
): number {
  const body = formula.trim().replace(/^=/, '')
  if (!body.trim()) throw new FormulaError('Syntax Error')
  return new Parser(tokenize(body), grid, row, column).parse()
}

/**
 * Word number-format picture (`\# "..."`): 0 and # digit placeholders, `,` thousands, `.` decimals,
 * `%` percent, and literal prefix/suffix text.
 */
export function formatWithPicture(value: number, picture: string | null | undefined): string {
  if (!picture || !picture.trim()) {
    const rounded = Math.round(value * 1e10) / 1e10
    return String(rounded)
  }
  const sections = picture.split(';')
  const negative = value < 0
  const section = negative && sections.length > 1 ? sections[1] : sections[0]
  const match = /^([^0#]*)([0#,]*)(?:\.([0#]*))?([^0#]*)$/.exec(section)
  if (!match) return String(value)
  const [, prefix, integerPart, decimalPart = '', suffix] = match
  const percent = /%/.test(section)
  const scaled = Math.abs(value) * (percent ? 100 : 1)
  const decimals = decimalPart.length
  const fixed = scaled.toFixed(decimals)
  let [whole, fraction = ''] = fixed.split('.')
  const minDigits = (integerPart.match(/0/g) ?? []).length
  whole = whole.padStart(minDigits, '0')
  if (whole === '0' && minDigits === 0 && decimals > 0) whole = ''
  if (integerPart.includes(',')) whole = whole.replace(/\B(?=(\d{3})+(?!\d))/g, ',')
  // trailing # decimals drop zeros
  const requiredFraction = (decimalPart.match(/0/g) ?? []).length
  while (fraction.length > requiredFraction && fraction.endsWith('0'))
    fraction = fraction.slice(0, -1)
  const sign = negative && sections.length === 1 ? '-' : ''
  return `${sign}${prefix}${whole}${fraction ? `.${fraction}` : ''}${suffix}`
}

/** Field instruction text Word writes for a table formula. */
export function formulaInstruction(formula: string, picture: string | null | undefined): string {
  const body = formula.trim().replace(/^=/, '')
  return `=${body}${picture?.trim() ? ` \\# "${picture.trim()}"` : ''}`
}

/** Parse ` =SUM(ABOVE) \# "0.00" ` back into its formula and picture. */
export function parseFormulaInstruction(
  instruction: string,
): { formula: string; picture: string | null } | null {
  const match = /^\s*=\s*(.*?)(?:\s+\\#\s+"([^"]*)")?\s*$/.exec(instruction)
  if (!match) return null
  return { formula: match[1], picture: match[2] ?? null }
}

/** FormulaGrid over a ProseMirror table (merged cells expose their text at every covered slot). */
export function tableFormulaGrid(table: PmNode): FormulaGrid {
  const map = TableMap.get(table)
  return {
    rows: map.height,
    columns: map.width,
    cell(row, column) {
      if (row < 0 || column < 0 || row >= map.height || column >= map.width) return null
      const node = table.nodeAt(map.map[row * map.width + column])
      return node ? { text: node.textContent } : null
    },
  }
}

export function formulaResultText(
  formula: string,
  picture: string | null | undefined,
  grid: FormulaGrid,
  row: number,
  column: number,
): string {
  return formatWithPicture(evaluateFormula(formula, grid, row, column), picture)
}

/** Default formula Word proposes: SUM over the numbers above, else to the left. */
export function defaultFormulaFor(grid: FormulaGrid, row: number, column: number): string {
  const above = directional(grid, row, column, 'ABOVE')
  if (valuesOf(above).length > 0) return '=SUM(ABOVE)'
  const left = directional(grid, row, column, 'LEFT')
  return valuesOf(left).length > 0 ? '=SUM(LEFT)' : '=SUM(ABOVE)'
}

/** Position of the caret's cell within the table grid. */
export function caretGridPosition(
  state: Parameters<Command>[0],
): { table: PmNode; row: number; column: number } | null {
  const selected = selectedTable(state)
  if (!selected) return null
  return { table: selected.node, row: selected.rect.top, column: selected.rect.left }
}

/** Table Layout > Formula: inserts the field and its calculated result at the caret. */
export function insertTableFormula(formula: string, picture: string | null): Command {
  return (state, dispatch) => {
    const here = caretGridPosition(state)
    const mark = state.schema.marks.instrField
    if (!here || !mark) return false
    let text: string
    try {
      text = formulaResultText(
        formula,
        picture,
        tableFormulaGrid(here.table),
        here.row,
        here.column,
      )
    } catch {
      return false
    }
    if (dispatch) {
      const node = state.schema.text(text || '0', [
        mark.create({ instr: formulaInstruction(formula, picture), dirty: false }),
      ])
      dispatch(state.tr.replaceSelectionWith(node, false).scrollIntoView())
    }
    return true
  }
}

/** Recalculate every `=` formula field of the caret's table (Word: select the table, F9). */
export function updateTableFormulas(): Command {
  return (state, dispatch) => {
    const selected = selectedTable(state)
    const mark = state.schema.marks.instrField
    if (!selected || !mark) return false
    const grid = tableFormulaGrid(selected.node)
    const map = TableMap.get(selected.node)
    const edits: { from: number; to: number; text: string; node: PmNode }[] = []
    selected.node.descendants((node, pos) => {
      if (!node.isText) return true
      const field = node.marks.find((m) => m.type === mark)
      const parsed = field ? parseFormulaInstruction(String(field.attrs.instr)) : null
      if (!field || !parsed) return true
      // locate the enclosing cell's grid slot
      const absolute = selected.tableStart + pos
      const $pos = state.doc.resolve(absolute)
      let cellStart = -1
      for (let depth = $pos.depth; depth > 0; depth--) {
        const role = $pos.node(depth).type.spec.tableRole
        if (role === 'cell' || role === 'header_cell') {
          cellStart = $pos.before(depth)
          break
        }
      }
      if (cellStart < 0) return true
      const rect = map.findCell(cellStart - selected.tableStart)
      try {
        const text = formulaResultText(parsed.formula, parsed.picture, grid, rect.top, rect.left)
        if (text !== node.text)
          edits.push({ from: absolute, to: absolute + node.nodeSize, text, node })
      } catch {
        // leave a field with an invalid formula untouched
      }
      return true
    })
    if (edits.length === 0) return false
    if (dispatch) {
      let tr = state.tr
      for (const edit of edits.sort((a, b) => b.from - a.from))
        tr = tr.replaceWith(edit.from, edit.to, state.schema.text(edit.text, edit.node.marks))
      dispatch(tr)
    }
    return true
  }
}
