import { Editor } from '@tiptap/core'
import { TextSelection } from '@tiptap/pm/state'
import { parseDocx, saveDocx } from '@genoffice/docx-engine'
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import {
  defaultFormulaFor,
  evaluateFormula,
  formatWithPicture,
  formulaInstruction,
  insertTableFormula,
  parseFormulaInstruction,
  tableFormulaGrid,
  updateTableFormulas,
  type FormulaGrid,
} from '../src/renderer/editor/table-formula'

const gridOf = (rows: string[][]): FormulaGrid => ({
  rows: rows.length,
  columns: rows[0].length,
  cell: (r, c) => (rows[r]?.[c] === undefined ? null : { text: rows[r][c] }),
})

describe('table formula evaluation', () => {
  const grid = gridOf([
    ['Item', 'Price', 'Qty'],
    ['a', '1,200.50', '2'],
    ['b', '$3', '4'],
    ['Total', '', ''],
  ])
  it('sums ABOVE, stopping at blanks and ignoring text', () => {
    expect(evaluateFormula('=SUM(ABOVE)', grid, 3, 1)).toBeCloseTo(1203.5)
    expect(evaluateFormula('SUM(ABOVE)', grid, 3, 2)).toBe(6)
    expect(evaluateFormula('=SUM(LEFT)', gridOf([['1', '2', '3', '']]), 0, 3)).toBe(6)
  })
  it('supports cell references, ranges, operators and functions', () => {
    expect(evaluateFormula('=B2*C2+B3*C3', grid, 3, 1)).toBeCloseTo(2413.0)
    expect(evaluateFormula('=AVERAGE(C2:C3)', grid, 3, 2)).toBe(3)
    expect(evaluateFormula('=MAX(B2:B3,5)', grid, 3, 1)).toBeCloseTo(1200.5)
    expect(evaluateFormula('=ROUND(10/3,2)', grid, 3, 2)).toBe(3.33)
    expect(evaluateFormula('=(2+3)*4-6/3', grid, 0, 0)).toBe(18)
    expect(evaluateFormula('=COUNT(B2:C3)', grid, 3, 1)).toBe(4)
    expect(evaluateFormula('=-5+MOD(7,4)', grid, 0, 0)).toBe(-2)
  })
  it('reports syntax and arithmetic errors', () => {
    expect(() => evaluateFormula('=1/0', grid, 0, 0)).toThrow('Division by zero')
    expect(() => evaluateFormula('=SUM(', grid, 0, 0)).toThrow()
    expect(() => evaluateFormula('=FOO(1)', grid, 0, 0)).toThrow('Unknown function')
    expect(() => evaluateFormula('', grid, 0, 0)).toThrow()
  })
  it('proposes SUM(ABOVE) or SUM(LEFT) like Word', () => {
    expect(defaultFormulaFor(grid, 3, 1)).toBe('=SUM(ABOVE)')
    expect(defaultFormulaFor(gridOf([['1', '2', '']]), 0, 2)).toBe('=SUM(LEFT)')
  })
})

describe('number format pictures', () => {
  it('formats decimals, thousands, percent and literals', () => {
    expect(formatWithPicture(1203.5, '0.00')).toBe('1203.50')
    expect(formatWithPicture(1234567.891, '#,##0.00')).toBe('1,234,567.89')
    expect(formatWithPicture(0.256, '0%')).toBe('26%')
    expect(formatWithPicture(1203.5, '$#,##0.00')).toBe('$1,203.50')
    expect(formatWithPicture(5, '0.0#')).toBe('5.0')
    expect(formatWithPicture(-5.5, '0.0')).toBe('-5.5')
    expect(formatWithPicture(-5.5, '0.0;(0.0)')).toBe('(5.5)')
    expect(formatWithPicture(2.5, null)).toBe('2.5')
  })
  it('round-trips the field instruction', () => {
    const instruction = formulaInstruction('=SUM(ABOVE)', '#,##0.00')
    expect(instruction).toBe('=SUM(ABOVE) \\# "#,##0.00"')
    expect(parseFormulaInstruction(instruction)).toEqual({
      formula: 'SUM(ABOVE)',
      picture: '#,##0.00',
    })
    expect(parseFormulaInstruction(' PAGE ')).toBeNull()
  })
})

const cell = (text: string) => `<w:tc><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`
const TABLE =
  '<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/></w:tblPr>' +
  '<w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>' +
  `<w:tr>${cell('x')}${cell('10')}</w:tr>` +
  `<w:tr>${cell('y')}${cell('32.5')}</w:tr>` +
  `<w:tr>${cell('Total')}${cell('')}</w:tr></w:tbl><w:p/>`

describe('formula field in a native table', () => {
  it('inserts a result with a field mark, saves a complex field, updates when cells change', async () => {
    const parsed = await parseDocx(await buildDocx({ bodyXml: TABLE }))
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })
    const cells: number[] = []
    editor.state.doc.descendants((node, pos) => {
      if (node.type.name === 'docTableCell' || node.type.name === 'docTableHeader') cells.push(pos)
    })
    editor.view.dispatch(
      editor.state.tr.setSelection(TextSelection.near(editor.state.doc.resolve(cells[5] + 1))),
    )
    expect(insertTableFormula('=SUM(ABOVE)', '0.00')(editor.state, editor.view.dispatch)).toBe(true)
    expect(editor.state.doc.nodeAt(cells[5])!.textContent).toBe('42.50')
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const saved = await saveDocx(parsed, plan.saveBlocks)
    const xml = await (await JSZip.loadAsync(saved)).file('word/document.xml')!.async('string')
    expect(xml).toContain('=SUM(ABOVE)')
    expect(xml).toMatch(/<w:t[^>]*>42\.50<\/w:t>/)
    // reopen: the cell still carries the field mark with the same instruction
    const reopened = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc((await parseDocx(saved)).blocks) as never,
    })
    let instruction = ''
    reopened.state.doc.descendants((node) => {
      const mark = node.marks.find((m) => m.type.name === 'instrField')
      if (mark) instruction = String(mark.attrs.instr)
    })
    expect(parseFormulaInstruction(instruction)).toEqual({ formula: 'SUM(ABOVE)', picture: '0.00' })
    reopened.destroy()
    // change a cell then update fields
    const secondCell = cells[3]
    editor.view.dispatch(
      editor.state.tr.setSelection(
        TextSelection.create(editor.state.doc, secondCell + 2, secondCell + 2 + 4),
      ),
    )
    editor.view.dispatch(editor.state.tr.insertText('100'))
    expect(updateTableFormulas()(editor.state, editor.view.dispatch)).toBe(true)
    expect(editor.state.doc.nodeAt(cells[5])!.textContent).toBe('110.00')
    expect(tableFormulaGrid(editor.state.doc.firstChild!).cell(2, 1)!.text).toBe('110.00')
    editor.destroy()
  })

  it('keeps an imported formula field a field when the table is regenerated by another edit', async () => {
    const field =
      '<w:tc><w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
      '<w:r><w:instrText xml:space="preserve"> =SUM(ABOVE) </w:instrText></w:r>' +
      '<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>42</w:t></w:r>' +
      '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p></w:tc>'
    const table =
      '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>' +
      `<w:tr>${cell('x')}${cell('42')}</w:tr><w:tr>${cell('Total')}${field}</w:tr></w:tbl><w:p/>`
    const parsed = await parseDocx(await buildDocx({ bodyXml: table }))
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })
    let first = -1
    editor.state.doc.descendants((node, pos) => {
      if (first < 0 && node.type.name.startsWith('docTable') && node.type.name !== 'docTable')
        first = pos
    })
    editor.view.dispatch(
      editor.state.tr.setSelection(TextSelection.near(editor.state.doc.resolve(first + 1))),
    )
    const { setTableAltText } = await import('../src/renderer/editor/table-tools')
    setTableAltText('t', '')(editor.state, editor.view.dispatch)
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const xml = await (
      await JSZip.loadAsync(await saveDocx(parsed, plan.saveBlocks))
    )
      .file('word/document.xml')!
      .async('string')
    expect(xml).toContain('<w:tblCaption w:val="t"/>')
    expect(xml).toContain('=SUM(ABOVE)')
    expect(xml).toMatch(/fldCharType="separate"/)
    editor.destroy()
  })
})
