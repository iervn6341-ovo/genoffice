import { Editor } from '@tiptap/core'
import { TextSelection } from '@tiptap/pm/state'
import { CellSelection } from '@tiptap/pm/tables'
import { parseDocx, saveDocx } from '@genoffice/docx-engine'
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { distributeSelectedColumns } from '../src/renderer/editor/table-sizing'
import {
  canSortTable,
  compareCellText,
  convertTableToText,
  convertTextToTable,
  selectInTable,
  setCellTextDirection,
  setRowsMayBreak,
  setTableAltText,
  setTableCellSpacing,
  sortTableRows,
} from '../src/renderer/editor/table-tools'

const cell = (text: string, extra = '') =>
  `<w:tc>${extra}<w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`
const row = (...cells: string[]) => `<w:tr>${cells.join('')}</w:tr>`
const TABLE =
  '<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/>' +
  '<w:tblLook w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>' +
  '</w:tblPr><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>' +
  row(cell('Name'), cell('Qty')) +
  row(cell('pear'), cell('10')) +
  row(cell('apple'), cell('2')) +
  row(cell('fig'), cell('1,500')) +
  '</w:tbl>'

async function open(body: string, afterTable = true) {
  const source = await buildDocx({ bodyXml: afterTable ? `${body}<w:p/>` : body })
  const parsed = await parseDocx(source)
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
  return { editor, parsed }
}

function cellPositions(editor: Editor): number[] {
  const out: number[] = []
  editor.state.doc.descendants((node, pos) => {
    if (node.type.name === 'docTableCell' || node.type.name === 'docTableHeader') out.push(pos)
  })
  return out
}

function caretInCell(editor: Editor, index: number) {
  const pos = cellPositions(editor)[index]
  editor.view.dispatch(
    editor.state.tr.setSelection(TextSelection.near(editor.state.doc.resolve(pos + 1))),
  )
}

async function saveAndReparse(editor: Editor, parsed: Awaited<ReturnType<typeof parseDocx>>) {
  const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
  const saved = await saveDocx(parsed, plan.saveBlocks)
  const zip = await JSZip.loadAsync(saved)
  return {
    plan,
    document: await zip.file('word/document.xml')!.async('string'),
    reparsed: await parseDocx(saved),
  }
}

const rowTexts = (reparsed: Awaited<ReturnType<typeof parseDocx>>) =>
  reparsed.blocks
    .find((block) => block.table)!
    .table!.rows.map((cells) => cells.map((c) => c.paras.join('|')))

describe('table alt text, cell spacing, row break rule and text direction', () => {
  it('saves alt text in schema order and reloads it; clearing removes the elements', async () => {
    const { editor, parsed } = await open(TABLE)
    caretInCell(editor, 0)
    expect(
      setTableAltText('Fruit stock', 'Quantity per fruit')(editor.state, editor.view.dispatch),
    ).toBe(true)
    const { document, reparsed } = await saveAndReparse(editor, parsed)
    expect(document).toContain(
      '/><w:tblCaption w:val="Fruit stock"/><w:tblDescription w:val="Quantity per fruit"/></w:tblPr>',
    )
    const model = reparsed.blocks.find((block) => block.table)!.table!
    expect(model).toMatchObject({ altTitle: 'Fruit stock', altDescription: 'Quantity per fruit' })
    // Undo restores the original and the untouched table stays byte-identical
    editor.commands.undo()
    expect(pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks).changedCount).toBe(0)
    editor.destroy()
  })

  it('writes cell spacing and the cannot-split row rule, then reads them back', async () => {
    const { editor, parsed } = await open(TABLE)
    caretInCell(editor, 2)
    setTableCellSpacing(57)(editor.state, editor.view.dispatch)
    setRowsMayBreak(false)(editor.state, editor.view.dispatch)
    const { document, reparsed } = await saveAndReparse(editor, parsed)
    expect(document).toContain('<w:tblCellSpacing w:w="57" w:type="dxa"/>')
    // schema order: tblW, tblCellSpacing, ..., tblLook
    expect(document.indexOf('<w:tblW')).toBeLessThan(document.indexOf('<w:tblCellSpacing'))
    expect(document.indexOf('<w:tblCellSpacing')).toBeLessThan(document.indexOf('<w:tblLook'))
    const model = reparsed.blocks.find((block) => block.table)!.table!
    expect(model.cellSpacingTwips).toBe(57)
    expect(model.cantSplitRows).toEqual([false, true, false, false])
    // turning the rule off again drops the element
    setRowsMayBreak(true)(editor.state, editor.view.dispatch)
    setTableCellSpacing(null)(editor.state, editor.view.dispatch)
    const second = await saveAndReparse(editor, parsed)
    expect(second.document).not.toContain('<w:cantSplit')
    expect(second.document).not.toContain('<w:tblCellSpacing')
    editor.destroy()
  })

  it('cycles the text direction of the selected cells and writes w:textDirection', async () => {
    const { editor, parsed } = await open(TABLE)
    const cells = cellPositions(editor)
    editor.view.dispatch(
      editor.state.tr.setSelection(CellSelection.create(editor.state.doc, cells[2], cells[3])),
    )
    setCellTextDirection()(editor.state, editor.view.dispatch)
    let result = await saveAndReparse(editor, parsed)
    let model = result.reparsed.blocks.find((block) => block.table)!.table!
    expect(model.rows[1][0].textDirection).toBe('tbRl')
    expect(model.rows[1][1].textDirection).toBe('tbRl')
    expect(model.rows[0][0].textDirection).toBeUndefined()
    setCellTextDirection()(editor.state, editor.view.dispatch)
    result = await saveAndReparse(editor, parsed)
    model = result.reparsed.blocks.find((block) => block.table)!.table!
    expect(model.rows[1][0].textDirection).toBe('btLr')
    setCellTextDirection()(editor.state, editor.view.dispatch)
    result = await saveAndReparse(editor, parsed)
    expect(result.document).not.toContain('w:textDirection')
    editor.destroy()
  })
})

describe('table sort', () => {
  it('orders numbers with a kept header row, writes the new order and undoes in one step', async () => {
    const { editor, parsed } = await open(TABLE)
    caretInCell(editor, 3)
    expect(canSortTable(editor.state)).toBe(true)
    const run = sortTableRows({
      keys: [{ column: 1, kind: 'number', descending: true }],
      header: true,
    })
    expect(run(editor.state, editor.view.dispatch)).toBe(true)
    const { reparsed } = await saveAndReparse(editor, parsed)
    expect(rowTexts(reparsed)).toEqual([
      ['Name', 'Qty'],
      ['fig', '1,500'],
      ['pear', '10'],
      ['apple', '2'],
    ])
    editor.commands.undo()
    expect(pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks).changedCount).toBe(0)
    editor.destroy()
  })

  it('sorts text ascending including the first row and refuses merged tables', async () => {
    const { editor } = await open(TABLE)
    caretInCell(editor, 0)
    sortTableRows({ keys: [{ column: 0, kind: 'text', descending: false }], header: false })(
      editor.state,
      editor.view.dispatch,
    )
    const texts: string[] = []
    editor.state.doc.descendants((node) => {
      if (node.type.name === 'docParagraph' && node.textContent) texts.push(node.textContent)
    })
    expect(texts.filter((_, index) => index % 2 === 0)).toEqual(['apple', 'fig', 'Name', 'pear'])
    editor.destroy()

    const merged = await open(
      TABLE.replace(cell('pear'), cell('pear', '<w:tcPr><w:gridSpan w:val="2"/></w:tcPr>')).replace(
        cell('10'),
        '',
      ),
    )
    caretInCell(merged.editor, 0)
    expect(canSortTable(merged.editor.state)).toBe(false)
    expect(
      sortTableRows({ keys: [{ column: 0, kind: 'text', descending: false }], header: true })(
        merged.editor.state,
      ),
    ).toBe(false)
    merged.editor.destroy()
  })

  it('compares values like Word: blanks first, numbers numerically, dates chronologically', () => {
    expect(compareCellText('', '5', 'number')).toBeLessThan(0)
    expect(compareCellText('1,500', '20', 'number')).toBeGreaterThan(0)
    expect(compareCellText('(3)', '2', 'number')).toBeLessThan(0)
    expect(compareCellText('2026-01-05', '2025-12-31', 'date')).toBeGreaterThan(0)
    expect(compareCellText('Apple', 'apple', 'text')).toBe(0)
  })
})

describe('convert table <-> text', () => {
  it('converts a table to tab-separated paragraphs that save without any table', async () => {
    const { editor, parsed } = await open(TABLE)
    caretInCell(editor, 0)
    expect(convertTableToText('tab')(editor.state, editor.view.dispatch)).toBe(true)
    const { reparsed, document } = await saveAndReparse(editor, parsed)
    expect(document).not.toContain('<w:tbl>')
    const texts = reparsed.blocks
      .filter((b) => b.type === 'paragraph')
      .map((b) => (b.runs ?? []).map((r) => r.text).join(''))
    expect(texts.slice(0, 4)).toEqual(['Name\tQty', 'pear\t10', 'apple\t2', 'fig\t1,500'])
    editor.commands.undo()
    expect(editor.state.doc.firstChild!.type.name).toBe('docTable')
    editor.destroy()
  })

  it('converts text paragraphs split by commas into a table, keeps marks, and round-trips', async () => {
    const { editor, parsed } = await open(
      '<w:p><w:r><w:t>a,b,c</w:t></w:r></w:p>' +
        '<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>d,e</w:t></w:r></w:p>',
      false,
    )
    editor.view.dispatch(
      editor.state.tr.setSelection(
        TextSelection.create(editor.state.doc, 1, editor.state.doc.content.size - 1),
      ),
    )
    expect(convertTextToTable({ separator: 'comma' })(editor.state, editor.view.dispatch)).toBe(
      true,
    )
    expect(editor.state.doc.firstChild!.type.name).toBe('docTable')
    const { reparsed } = await saveAndReparse(editor, parsed)
    expect(rowTexts(reparsed)).toEqual([
      ['a', 'b', 'c'],
      ['d', 'e', ''],
    ])
    const bold = reparsed.blocks.find((b) => b.table)!.table!.rows[1][0].richParas![0].runs[0]
    expect(bold).toMatchObject({ text: 'd', bold: true })
    editor.destroy()
  })

  it('turns paragraphs into cells row by row for the paragraph separator', async () => {
    const { editor, parsed } = await open(
      ['one', 'two', 'three'].map((t) => `<w:p><w:r><w:t>${t}</w:t></w:r></w:p>`).join(''),
      false,
    )
    editor.view.dispatch(
      editor.state.tr.setSelection(
        TextSelection.create(editor.state.doc, 1, editor.state.doc.content.size - 1),
      ),
    )
    convertTextToTable({ separator: 'paragraph', columns: 2 })(editor.state, editor.view.dispatch)
    const { reparsed } = await saveAndReparse(editor, parsed)
    expect(rowTexts(reparsed)).toEqual([
      ['one', 'two'],
      ['three', ''],
    ])
    editor.destroy()
  })
})

describe('selection and column distribution', () => {
  it('selects the row, column and table as cell selections', async () => {
    const { editor } = await open(TABLE)
    caretInCell(editor, 3)
    selectInTable('row')(editor.state, editor.view.dispatch)
    let selection = editor.state.selection as CellSelection
    expect(selection.$anchorCell.pos).toBe(cellPositions(editor)[2])
    expect(selection.$headCell.pos).toBe(cellPositions(editor)[3])
    selectInTable('table')(editor.state, editor.view.dispatch)
    selection = editor.state.selection as CellSelection
    expect(selection.$anchorCell.pos).toBe(cellPositions(editor)[0])
    expect(selection.$headCell.pos).toBe(cellPositions(editor).at(-1))
    editor.destroy()
  })

  it('distributes unequal columns to the average width', async () => {
    const { editor } = await open(
      TABLE.replace('w:w="3000"/><w:gridCol w:w="3000"', 'w:w="1500"/><w:gridCol w:w="4500"'),
    )
    caretInCell(editor, 0)
    expect(distributeSelectedColumns(600)(editor.state, editor.view.dispatch)).toBe(true)
    const first = editor.state.doc.nodeAt(cellPositions(editor)[0])!
    const second = editor.state.doc.nodeAt(cellPositions(editor)[1])!
    expect(first.attrs.colwidth[0]).toBeCloseTo(second.attrs.colwidth[0], 1)
    editor.destroy()
  })
})
