/**
 * Copy / cut / paste plumbing for the Docs editor (see clipboard-transfer.ts
 * for the payload). The copy handler writes
 *   text/html   rich fragment for other programs + the GenOffice slice payload
 *   text/plain  Word-style text (tables as tab-separated rows for Excel)
 * and the clipboard parser rebuilds the exact slice from the payload when the
 * pasted HTML carries one, falling back to the schema's parse rules for
 * everything else (Word, web pages, other editors).
 */
import { Extension } from '@tiptap/core'
import {
  DOMParser as PmDOMParser,
  type Fragment,
  type Node as PmNode,
  type Slice,
} from '@tiptap/pm/model'
import { Plugin, PluginKey, type Transaction } from '@tiptap/pm/state'
import type { EditorView } from '@tiptap/pm/view'
import {
  CLIP_SLICE_ATTR,
  buildClipPayload,
  currentClipDocKey,
  encodeClipPayload,
  payloadFromDom,
  payloadToSlice,
  type ClipDest,
  type ClipSource,
} from './clipboard-transfer'
import { enrichClipboardDom, type HostMounter, type ListItemDecor } from './clipboard-html'

/** what the host application (App.tsx) provides for the current document */
export interface ClipHost {
  source(): ClipSource
  dest(): ClipDest
  /** a paste landed: follow-up edits (renumbering imported note markers), appended to the paste's undo step */
  afterPaste?(tr: Transaction): Transaction | null
}

export interface ClipboardTransferStorage {
  host: ClipHost | null
}

declare module '@tiptap/core' {
  interface Storage {
    clipboardTransfer: ClipboardTransferStorage
  }
}

const KEY = new PluginKey('docsClipboardTransfer')

/** minimal host for editors without App state (tests, headless) */
export function fallbackHost(): ClipHost {
  return {
    source: () => ({ docKey: currentClipDocKey() }),
    dest: () => ({ docKey: currentClipDocKey() }),
  }
}

// ── plain text flavor ────────────────────────────────────────────────────

function cellText(cell: PmNode): string {
  const paras: string[] = []
  cell.forEach((child) => {
    paras.push(
      child.type.name === 'docNestedTable' || child.type.name === 'docCellBoxes'
        ? ''
        : child.textBetween(0, child.content.size, '\n', (leaf) =>
            leaf.type.name === 'hardBreak' ? '\n' : '',
          ),
    )
  })
  return paras.join('\n')
}

/** Excel-style quoting for a cell that contains tabs, line breaks or quotes */
function tsvField(text: string): string {
  return /[\t\n\r"]/.test(text) ? `"${text.replace(/"/g, '""')}"` : text
}

function tableText(table: PmNode): string {
  const rows: string[] = []
  table.forEach((row) => {
    const cells: string[] = []
    row.forEach((cell) => cells.push(tsvField(cellText(cell))))
    rows.push(cells.join('\t'))
  })
  return rows.join('\n')
}

/** Word-like plain text: one line break between paragraphs, tables as TSV */
export function clipboardPlainText(content: Fragment): string {
  const blocks: string[] = []
  content.forEach((node) => {
    if (node.type.name === 'docTable') blocks.push(tableText(node))
    else if (node.isTextblock) {
      blocks.push(
        node.textBetween(0, node.content.size, '\n', (leaf) => {
          if (leaf.type.name === 'hardBreak') return leaf.attrs.pageBreak ? '\f' : '\n'
          if (leaf.type.name === 'docInlineMath') return String(leaf.attrs.text ?? '')
          return ''
        }),
      )
    } else if (node.isAtom) {
      const label = String(node.attrs.previewText ?? '')
      if (label) blocks.push(label)
    } else {
      const inner = clipboardPlainText(node.content)
      if (inner) blocks.push(inner)
    }
  })
  return blocks.join('\n')
}

/**
 * A scratch element placed beside the editor, wearing the editor's classes and
 * inline custom properties so the document stylesheet resolves inside it. Off
 * screen but laid out (computed widths need a layout box).
 */
function mounterFor(view: EditorView): HostMounter | undefined {
  const root = view.dom as HTMLElement
  const parent = root.parentElement
  if (!parent || !root.isConnected) return undefined
  return (host) => {
    host.className = root.className.replace(/\b(?:ProseMirror-focused|page-dark|show-marks)\b/g, '')
    host.setAttribute('style', root.getAttribute('style') ?? '')
    for (const attr of Array.from(root.attributes)) {
      if (attr.name.startsWith('data-')) host.setAttribute(attr.name, attr.value)
    }
    host.style.position = 'absolute'
    host.style.left = '-100000px'
    host.style.top = '0'
    host.style.width = `${root.clientWidth || 816}px`
    host.style.pointerEvents = 'none'
    host.setAttribute('aria-hidden', 'true')
    parent.appendChild(host)
    return () => host.remove()
  }
}

/** marker decorations of the list items inside [from, to], in document order */
function liveListDecor(view: EditorView, from: number, to: number): ListItemDecor[] {
  const out: ListItemDecor[] = []
  view.state.doc.nodesBetween(from, to, (node, pos) => {
    if (node.type.name !== 'docListItem') return true
    const el = view.nodeDOM(pos)
    const attrs: Record<string, string> = {}
    const vars: string[] = []
    if (el instanceof HTMLElement) {
      for (const name of ['data-marker', 'data-marker-pic', 'data-marker-flow']) {
        const v = el.getAttribute(name)
        if (v !== null) attrs[name] = v
      }
      for (let i = 0; i < el.style.length; i++) {
        const prop = el.style.item(i)
        if (prop.startsWith('--li-')) vars.push(`${prop}:${el.style.getPropertyValue(prop)}`)
      }
    }
    out.push({ attrs, vars })
    return false
  })
  return out
}

// ── payload parsing ──────────────────────────────────────────────────────

/** the editor's clipboard HTML parser: GenOffice payload first, schema rules otherwise */
export function parseClipboardDom(
  view: EditorView,
  dom: Node,
  host: ClipHost | null,
  options: Parameters<PmDOMParser['parseSlice']>[1],
): Slice {
  const schema = view.state.schema
  if (host && 'querySelector' in dom) {
    const payload = payloadFromDom(dom as ParentNode)
    if (payload) {
      try {
        return payloadToSlice(schema, payload, host.dest())
      } catch {
        /* unreadable payload: parse the HTML below */
      }
    }
  }
  return PmDOMParser.fromSchema(schema).parseSlice(dom, options)
}

/** the exact content a GenOffice payload in `html` stands for, or null when `html` carries none
 *  (the wholesale empty-paragraph paste lane re-parses the HTML itself) */
export function payloadFragment(view: EditorView, html: string): Fragment | null {
  const host = (view as unknown as { __clipHost?: ClipHost }).__clipHost ?? null
  if (!host || !html.includes(CLIP_SLICE_ATTR)) return null
  const doc = new window.DOMParser().parseFromString(html, 'text/html')
  const payload = payloadFromDom(doc.body)
  if (!payload) return null
  try {
    const content = payloadToSlice(view.state.schema, payload, host.dest()).content
    let allBlocks = content.childCount > 0
    content.forEach((node) => {
      if (!node.isBlock) allBlocks = false
    })
    return allBlocks ? content : null
  } catch {
    return null
  }
}

// ── the extension ────────────────────────────────────────────────────────

export const ClipboardTransferExtension = Extension.create<object, ClipboardTransferStorage>({
  name: 'clipboardTransfer',
  addStorage() {
    return { host: null }
  },
  addProseMirrorPlugins() {
    const getHost = (): ClipHost | null => this.storage.host
    const copyHandler = (view: EditorView, event: ClipboardEvent, cut: boolean): boolean => {
      const sel = view.state.selection
      const data = event.clipboardData
      if (sel.empty || !data) return false
      const slice = sel.content()
      const { dom, text } = view.serializeForClipboard(slice)
      const host = getHost()
      let html: string
      try {
        if (host) {
          const encoded = encodeClipPayload(buildClipPayload(slice, host.source()))
          let first = dom.firstElementChild
          if (!first && dom.firstChild) {
            // bare text selection: wrap so the payload has an element to ride on
            const span = dom.ownerDocument.createElement('span')
            while (dom.firstChild) span.appendChild(dom.firstChild)
            dom.appendChild(span)
            first = span
          }
          if (encoded && first) first.setAttribute(CLIP_SLICE_ATTR, encoded)
        }
        html = enrichClipboardDom(
          dom as HTMLElement,
          mounterFor(view),
          liveListDecor(view, sel.from, sel.to),
        )
      } catch {
        html = (dom as HTMLElement).innerHTML
      }
      event.preventDefault()
      data.clearData()
      data.setData('text/html', html)
      data.setData('text/plain', text)
      if (cut) {
        view.dispatch(view.state.tr.deleteSelection().scrollIntoView().setMeta('uiEvent', 'cut'))
      }
      return true
    }
    const editor = this.editor
    return [
      new Plugin({
        key: KEY,
        view: (view) => {
          // the paste-lane helpers only get the view; expose the host through it
          Object.defineProperty(view, '__clipHost', {
            configurable: true,
            get: () => getHost(),
          })
          return {}
        },
        appendTransaction: (trs, _old, state) => {
          const host = getHost()
          if (!host?.afterPaste || !trs.some((tr) => tr.getMeta('uiEvent') === 'paste')) return null
          return host.afterPaste(state.tr)
        },
        props: {
          handleDOMEvents: {
            copy: (view, event) => copyHandler(view, event as ClipboardEvent, false),
            cut: (view, event) =>
              view.editable ? copyHandler(view, event as ClipboardEvent, true) : false,
          },
          clipboardTextSerializer: (slice) => clipboardPlainText(slice.content),
          // parseSlice has no view argument: the editor is the extension's own
          clipboardParser: {
            parseSlice: (dom: Node, options: Parameters<PmDOMParser['parseSlice']>[1]) =>
              parseClipboardDom(editor.view, dom, getHost(), options),
          } as unknown as PmDOMParser,
        },
      }),
    ]
  },
})
