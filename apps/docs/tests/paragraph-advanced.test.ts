import JSZip from 'jszip'
import { Editor } from '@tiptap/core'
import { TextSelection } from '@tiptap/pm/state'
import { createElement } from 'react'
import { act } from 'react'
import { createRoot } from 'react-dom/client'
import { describe, expect, it, vi } from 'vitest'
import { parseDocx, saveDocx } from '@genoffice/docx-engine'
import {
  buildDocx,
  IMAGE_PARAGRAPH_XML,
  MATH_PARAGRAPH_XML,
} from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { LAYOUT_ATTR, type LayoutSnapshot } from '../src/renderer/editor/layout-history'
import { ParagraphDialog } from '../src/renderer/components/ContextMenu'
import { LocaleProvider, setModuleLang } from '../src/renderer/i18n/locale'
import {
  applyParagraphChanges,
  commonParagraphValue,
  selectedParagraphs,
} from '../src/renderer/editor/paragraph-advanced'
import { sameBodyAfterManagedRPrRebase } from '../src/renderer/editor/save-rebase'

const importedPPr =
  '<w:pPr><w:keepNext w:val="0"/><w:widowControl w:val="0"/>' +
  '<w:tabs><w:tab w:val="right" w:pos="1440" w:leader="heavy"/></w:tabs>' +
  '<w:spacing w:after="200"/><w:contextualSpacing/><w:rPr><w:lang w:val="ja-JP"/></w:rPr></w:pPr>'

async function open() {
  const bytes = await buildDocx({
    withImage: true,
    bodyXml:
      `<w:p>${importedPPr}<w:r><w:t>Alpha </w:t></w:r>` +
      '<w:fldSimple w:instr=" DATE "><w:r><w:t>October</w:t></w:r></w:fldSimple>' +
      '<w:r><w:t> tail</w:t></w:r></w:p>' +
      '<w:p><w:pPr><w:keepLines/></w:pPr><w:r><w:t>Second</w:t></w:r></w:p>' +
      IMAGE_PARAGRAPH_XML +
      MATH_PARAGRAPH_XML,
  })
  const parsed = await parseDocx(bytes)
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
  const save = async () => {
    const plan = pmDocToSavePlan(editor.state.doc.toJSON() as PmNode, parsed.blocks)
    const result = await saveDocx(parsed, plan.saveBlocks)
    return (await JSZip.loadAsync(result)).file('word/document.xml')!.async('string')
  }
  return { editor, parsed, save }
}

describe('Paragraph Advanced authoring', () => {
  it('applies mixed pagination and Asian controls in one undo step without changing fields, math, or pictures', async () => {
    const { editor, parsed, save } = await open()
    const paragraphs = selectedParagraphs(editor, 1, editor.state.doc.content.size - 1)
    expect(paragraphs.length).toBeGreaterThanOrEqual(2)
    expect(commonParagraphValue(paragraphs.slice(0, 2), (p) => p.node.attrs.keepNext).mixed).toBe(
      true,
    )
    editor.view.dispatch(
      editor.state.tr.setSelection(
        TextSelection.create(editor.state.doc, paragraphs[0].pos + 1, paragraphs[1].pos + 2),
      ),
    )
    const original = await save()
    expect(
      applyParagraphChanges(editor, editor.state.selection.from, editor.state.selection.to, {
        keepNext: true,
        suppressAutoHyphens: true,
        kinsoku: false,
        topLinePunct: true,
        textAlignment: 'center',
      }),
    ).toBe(true)
    const changed = await save()
    expect(changed).toContain('<w:keepNext/>')
    expect(changed).toContain('<w:suppressAutoHyphens/>')
    expect(changed).toContain('<w:kinsoku w:val="0"/>')
    expect(changed).toContain('<w:topLinePunct/>')
    expect(changed).toContain('<w:textAlignment w:val="center"/>')
    for (const xml of [
      '<w:fldSimple w:instr=" DATE "><w:r><w:t>October</w:t></w:r></w:fldSimple>',
      '<w:contextualSpacing/>',
      '<w:rPr><w:lang w:val="ja-JP"/></w:rPr>',
      '<w:drawing>',
      '<m:oMath>',
    ])
      expect(changed).toContain(xml)
    const plan = pmDocToSavePlan(editor.state.doc.toJSON() as PmNode, parsed.blocks)
    const reopened = await parseDocx(await saveDocx(parsed, plan.saveBlocks))
    expect(
      sameBodyAfterManagedRPrRebase(
        editor.state.doc,
        editor.schema.nodeFromJSON(blocksToPmDoc(reopened.blocks)),
      ),
    ).toBe(true)
    editor.commands.undo()
    expect(await save()).toBe(original)
    editor.commands.redo()
    expect(await save()).toBe(changed)
    editor.destroy()
  })

  it('clears only direct tab stops and keeps an empty override through save/reopen rebase', async () => {
    const { editor, parsed, save } = await open()
    const first = selectedParagraphs(editor, 1, 1)[0]
    expect(first.node.attrs.tabStops).toContain('1440')
    applyParagraphChanges(editor, first.pos + 1, first.pos + 1, { tabStops: '[]' })
    const changed = await save()
    expect(changed).not.toContain('<w:tab w:val="right"')
    const plan = pmDocToSavePlan(editor.state.doc.toJSON() as PmNode, parsed.blocks)
    const reopened = await parseDocx(await saveDocx(parsed, plan.saveBlocks))
    expect(
      sameBodyAfterManagedRPrRebase(
        editor.state.doc,
        editor.schema.nodeFromJSON(blocksToPmDoc(reopened.blocks)),
      ),
    ).toBe(true)
    editor.commands.undo()
    expect(await save()).toContain('<w:tab w:val="right" w:pos="1440" w:leader="heavy"/>')
    editor.destroy()
  })

  it('removes newly authored native pPr children on Undo after Save and retains Redo', async () => {
    const { editor, parsed } = await open()
    const saveAgainst = async (baseline: typeof parsed) => {
      const plan = pmDocToSavePlan(editor.state.doc.toJSON() as PmNode, baseline.blocks)
      const bytes = await saveDocx(baseline, plan.saveBlocks)
      const body = await (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string')
      return { parsed: await parseDocx(bytes), body }
    }
    applyParagraphChanges(editor, 1, 1, {
      suppressAutoHyphens: true,
      kinsoku: false,
      topLinePunct: true,
      textAlignment: 'center',
      tabStops: JSON.stringify([{ pos: 1800, val: 'decimal', leader: 'dot' }]),
    })
    const first = await saveAgainst(parsed)
    expect(first.body).toContain('<w:suppressAutoHyphens/>')
    editor.commands.undo()
    const undone = await saveAgainst(first.parsed)
    expect(undone.body).not.toContain('<w:suppressAutoHyphens')
    expect(undone.body).not.toContain('<w:kinsoku')
    expect(undone.body).not.toContain('<w:topLinePunct')
    expect(undone.body).not.toContain('<w:textAlignment')
    expect(undone.body).toContain('<w:fldSimple w:instr=" DATE ">')
    const reopened = editor.schema.nodeFromJSON(blocksToPmDoc(undone.parsed.blocks))
    const comparable = editor.schema.nodeFromJSON({
      ...reopened.toJSON(),
      attrs: {
        ...reopened.attrs,
        [LAYOUT_ATTR]: editor.state.doc.attrs[LAYOUT_ATTR],
      },
    })
    expect(sameBodyAfterManagedRPrRebase(editor.state.doc, comparable)).toBe(true)
    expect(editor.commands.redo()).toBe(true)
    const redone = await saveAgainst(undone.parsed)
    expect(redone.body).toContain('<w:suppressAutoHyphens/>')
    expect(redone.body).toContain('<w:fldSimple w:instr=" DATE ">')
    editor.destroy()
  })

  it('groups paragraph and document-wide tab/compression edits into one Undo/Redo transaction', async () => {
    const { editor } = await open()
    const before = {
      sections: [],
      sectionsDirty: [],
      titlePg: false,
      titlePgDirty: false,
      evenOddHf: false,
      evenOddHfDirty: false,
      pgNumEdit: null,
      pgNumDirtySections: [],
      section: null,
      sectionDirty: false,
      design: {
        docFlags: {
          autoHyphenation: false,
          mirrorMargins: false,
          characterSpacingControl: null,
          defaultTabStopTwips: null,
          paragraphSpacing: null,
        },
        docFlagsDirty: false,
      },
    } as unknown as LayoutSnapshot
    const after = {
      ...before,
      design: {
        ...before.design,
        docFlags: {
          ...before.design!.docFlags!,
          characterSpacingControl: 'compressPunctuationAndJapaneseKana' as const,
          defaultTabStopTwips: 480,
        },
        docFlagsDirty: true,
      },
    } as LayoutSnapshot
    expect(
      applyParagraphChanges(editor, 1, 1, { keepNext: true }, undefined, { before, after }),
    ).toBe(true)
    expect(editor.state.doc.firstChild?.attrs.keepNext).toBe(true)
    expect(editor.state.doc.attrs[LAYOUT_ATTR].design.docFlags.defaultTabStopTwips).toBe(480)
    editor.commands.undo()
    expect(editor.state.doc.firstChild?.attrs.keepNext).toBe(false)
    expect(editor.state.doc.attrs[LAYOUT_ATTR].design.docFlags.defaultTabStopTwips).toBeNull()
    editor.commands.redo()
    expect(editor.state.doc.firstChild?.attrs.keepNext).toBe(true)
    expect(editor.state.doc.attrs[LAYOUT_ATTR].design.docFlags.characterSpacingControl).toBe(
      'compressPunctuationAndJapaneseKana',
    )
    editor.destroy()
  })

  it('stages Asian settings until OK and leaves the document unchanged on Cancel', async () => {
    Object.assign(window, { desktop: { onLanguageChanged: () => () => undefined } })
    setModuleLang('en')
    const { editor } = await open()
    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)
    const onApplyAdvanced = vi.fn()
    const onClose = vi.fn()
    const render = () =>
      act(() =>
        root.render(
          createElement(LocaleProvider, {
            initial: 'en',
            children: createElement(ParagraphDialog, {
              editor,
              onClose,
              compression: 'compressPunctuation',
              defaultTabStopTwips: 720,
              onApplyAdvanced,
            }),
          }),
        ),
      )
    render()
    act(() =>
      host.querySelector<HTMLButtonElement>('[data-testid="paragraph-pane-asian"]')!.click(),
    )
    const compression = [...host.querySelectorAll<HTMLButtonElement>('.gs-dd-btn')].find(
      (button) => button.getAttribute('aria-label') === 'Character compression',
    )!
    act(() => compression.click())
    act(() =>
      host
        .querySelector<HTMLButtonElement>(
          '.gs-dd-item[data-value="compressPunctuationAndJapaneseKana"]',
        )!
        .click(),
    )
    expect(onApplyAdvanced).not.toHaveBeenCalled()
    act(() =>
      [...host.querySelectorAll<HTMLButtonElement>('button')]
        .find((button) => button.textContent === 'Cancel')!
        .click(),
    )
    expect(onClose).toHaveBeenCalledOnce()
    expect(onApplyAdvanced).not.toHaveBeenCalled()
    act(() => root.unmount())
    host.remove()
    editor.destroy()
  })
})
