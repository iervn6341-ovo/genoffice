/**
 * Cross-document clipboard harness: opens docx bytes into a real editor (same
 * extensions as the app), performs a copy through the editor's own copy
 * handler into a fake DataTransfer, and pastes that payload into another
 * editor through ProseMirror's own paste pipeline.
 */
import { Editor } from '@tiptap/core'
import { AllSelection, NodeSelection, TextSelection } from '@tiptap/pm/state'
import JSZip from 'jszip'
import { parseDocx, saveDocx, type ParsedDoc, type StyleUpsert } from '@genoffice/docx-engine'
import {
  blocksToPmDoc,
  pmDocOptions,
  pmDocToSavePlan,
  type PmNode,
} from '../../src/renderer/editor/convert'
import { editorExtensions } from '../../src/renderer/editor/extensions'
import { cleanPastedHtml } from '../../src/renderer/editor/paste-clean'
import {
  bumpClipDocGeneration,
  currentClipDocKey,
} from '../../src/renderer/editor/clipboard-transfer'
import type { ClipHost } from '../../src/renderer/editor/clipboard-extension'

export interface OpenedDoc {
  editor: Editor
  parsed: Awaited<ReturnType<typeof parseDocx>>
  docKey: string
  /** styles / numbering imported by pastes into this document */
  imported: {
    styles: Array<{ id: string; xml?: string }>
    numbering: string[]
    notes: string[]
    newDefs: Array<{
      numId: string
      kind: 'bullet' | 'ordered'
      levels: Array<{
        numFmt: string
        lvlText: string
        indentLeft: number
        hanging?: number
        start?: number
        font?: string
      }>
    }>
  }
  footnotes: Array<{ id: string; text: string }>
  styleUpserts: StyleUpsert[]
  /** save the document as the app would (patch engine) and open the resulting package */
  save(): Promise<{ zip: JSZip; documentXml: string; stylesXml: string; numberingXml: string }>
  destroy(): void
}

export async function openDocx(bytes: Uint8Array): Promise<OpenedDoc> {
  const parsed = await parseDocx(bytes)
  bumpClipDocGeneration()
  const docKey = currentClipDocKey()
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks, undefined, pmDocOptions(parsed)) as never,
    editorProps: { transformPastedHTML: cleanPastedHtml },
  })
  editor.storage.listNumbering.styles = parsed.styles
  editor.storage.listNumbering.defs = parsed.numbering
  const doc: OpenedDoc = {
    editor,
    parsed,
    docKey,
    imported: { styles: [], numbering: [], notes: [], newDefs: [] },
    footnotes: parsed.footnotes.map((n) => ({ id: n.id, text: n.text })),
    styleUpserts: [],
    save: async () => {
      const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
      const bytes = await saveDocx(parsed, plan.saveBlocks, {
        ...(doc.styleUpserts.length ? { styleUpserts: doc.styleUpserts } : {}),
        ...(doc.imported.newDefs.length ? { numbering: { newDefs: doc.imported.newDefs } } : {}),
      } as never)
      const zip = await JSZip.loadAsync(bytes)
      const part = async (p: string) => (await zip.file(p)?.async('string')) ?? ''
      return {
        zip,
        documentXml: await part('word/document.xml'),
        stylesXml: await part('word/styles.xml'),
        numberingXml: await part('word/numbering.xml'),
      }
    },
    destroy: () => editor.destroy(),
  }
  const host: ClipHost = {
    source: () => ({
      docKey,
      styles: parsed.styles,
      styleXml: (id) => parsed.extras?.styleXml?.[id],
      numbering: parsed.numbering,
      footnotes: parsed.footnotes,
      endnotes: parsed.endnotes,
      blockXml: (i) => parsed.blocks.find((b) => b.docxIndex === i)?.originalXml ?? undefined,
    }),
    dest: () => ({
      docKey,
      styles: parsed.styles,
      importStyle: (id, def) => {
        parsed.styles.set(id, def.info)
        doc.imported.styles.push({ id, xml: def.xml })
        doc.styleUpserts.push(
          def.xml
            ? { styleId: id, type: def.info.type, rawXml: def.xml }
            : { styleId: id, type: def.info.type, name: def.info.name },
        )
        return id
      },
      numbering: parsed.numbering,
      importNumbering: (def) => {
        const numId = String(100 + doc.imported.numbering.length)
        parsed.numbering.set(numId, { ...def, numId })
        editor.storage.listNumbering.defs = new Map(parsed.numbering)
        doc.imported.numbering.push(numId)
        const lv = Object.keys(def.levels)
          .map(Number)
          .sort((a, b) => a - b)
          .map((i) => ({
            numFmt: def.levels[i]!.numFmt,
            lvlText: def.levels[i]!.lvlText,
            indentLeft: def.levels[i]!.indentLeft ?? 720 * (i + 1),
            hanging: def.levels[i]!.hanging ?? 360,
            start: def.levels[i]!.start ?? 1,
            ...(def.levels[i]!.font ? { font: def.levels[i]!.font } : {}),
          }))
        doc.imported.newDefs.push({
          numId,
          kind: lv[0]?.numFmt === 'bullet' ? 'bullet' : 'ordered',
          levels: lv,
        })
        return numId
      },
      importNote: (kind, note) => {
        const id = String(50 + doc.imported.notes.length)
        doc.imported.notes.push(`${kind}:${id}:${note.text}`)
        doc.footnotes.push({ id, text: note.text })
        return id
      },
    }),
  }
  editor.storage.clipboardTransfer.host = host
  return doc
}

export interface Clip {
  html: string
  text: string
}

function fakeData(init?: Clip) {
  const store = new Map<string, string>()
  if (init) {
    store.set('text/html', init.html)
    store.set('text/plain', init.text)
  }
  return {
    store,
    clearData: () => store.clear(),
    setData: (type: string, value: string) => void store.set(type, value),
    getData: (type: string) => store.get(type) ?? '',
    get types() {
      return [...store.keys()]
    },
    items: [] as unknown[],
    files: [] as unknown[],
  }
}

function fire(editor: Editor, type: 'copy' | 'cut' | 'paste', data: ReturnType<typeof fakeData>) {
  const ev = new Event(type, { bubbles: true, cancelable: true })
  Object.defineProperty(ev, 'clipboardData', { value: data })
  editor.view.dom.dispatchEvent(ev)
  return ev
}

export function copy(editor: Editor, cut = false): Clip {
  const data = fakeData()
  const ev = fire(editor, cut ? 'cut' : 'copy', data)
  if (!ev.defaultPrevented) throw new Error('copy handler did not run')
  return { html: data.getData('text/html'), text: data.getData('text/plain') }
}

export function paste(editor: Editor, clip: Clip): void {
  fire(editor, 'paste', fakeData(clip))
}

export function selectAll(editor: Editor): void {
  editor.view.dispatch(editor.state.tr.setSelection(new AllSelection(editor.state.doc)))
}
export function selectNode(editor: Editor, pos: number): void {
  editor.view.dispatch(editor.state.tr.setSelection(NodeSelection.create(editor.state.doc, pos)))
}
export function selectText(editor: Editor, from: number, to: number): void {
  editor.view.dispatch(
    editor.state.tr.setSelection(TextSelection.create(editor.state.doc, from, to)),
  )
}
export function caretAtEnd(editor: Editor): void {
  editor.view.dispatch(editor.state.tr.setSelection(TextSelection.atEnd(editor.state.doc)))
}

/** top-level blocks of the editor document as JSON */
export function blocks(editor: Editor): PmNode[] {
  return (editor.getJSON().content ?? []) as PmNode[]
}

export type { ParsedDoc }
