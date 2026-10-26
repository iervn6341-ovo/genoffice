/**
 * Formula text on its way out of a workbook. Excel (and another workbook window) can only keep a
 * formula whose references still mean something where it lands: other sheets of the source
 * workbook and workbook-level names are not there.
 */
import { offsetFormulaRefs } from '@genoffice/xlsx-gateway/domain/formula-shift'

const CELL_REF = /^\$?[A-Za-z]{1,3}\$?\d+$/
const COLUMN_REF = /^\$?[A-Za-z]{1,3}$/
const RESERVED = new Set(['TRUE', 'FALSE'])

/// Defined names (and other bare identifiers that are not cell references, functions or
/// booleans) a formula uses, upper-cased.
export function nameReferences(formula: string): string[] {
  const text = formula
    .replace(/"(?:[^"]|"")*"/g, '""')
    .replace(/'(?:[^']|'')*'!/g, '')
    .replace(/\b[A-Za-z_][\w.]*!/g, '')
  const names = new Set<string>()
  const re = /(?<![\w.$])([A-Za-z_\\][\w.]*)(?![\w.]*\s*\()(?!\s*[(!])/g
  let match: RegExpExecArray | null
  while ((match = re.exec(text)) !== null) {
    const token = match[1]!
    const after = text[match.index + token.length]
    if (CELL_REF.test(token) || RESERVED.has(token.toUpperCase())) continue
    // whole-column spans (A:C) carry no name
    if (COLUMN_REF.test(token) && (after === ':' || text[match.index - 1] === ':')) continue
    // R1C1 style or structured-reference column parts are not names this check understands
    if (/^[RC]\d*$/i.test(token)) continue
    names.add(token.toUpperCase())
  }
  return [...names]
}

/// A formula as Excel reads it from the clipboard markup: Excel puts the pasted block's first
/// cell at A1 for the purpose of moving relative references, so the references are moved to
/// where they would point if the block started at A1. Null when a reference would leave the grid.
export function formulaForExcelClipboard(
  formula: string,
  firstRow: number,
  firstColumn: number,
): string | null {
  const text = formula.startsWith('=') ? formula : `=${formula}`
  const moved = offsetFormulaRefs(text, -firstRow, -firstColumn)
  return moved.includes('#REF!') && !text.includes('#REF!') ? null : moved
}
