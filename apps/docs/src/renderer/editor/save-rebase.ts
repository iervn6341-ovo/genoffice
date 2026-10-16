import type { Editor } from '@tiptap/core'
import type { Node as PmNode } from '@tiptap/pm/model'
import { splitXmlChildren, type Block } from '@genoffice/docx-engine'

// These rPr children have equivalent typed ProseMirror attrs or marks. The
// writer may insert or reorder them without changing the user's live edit.
const MODELED_RPR_CHILDREN = new Set([
  'w:rFonts',
  'w:b',
  'w:bCs',
  'w:i',
  'w:iCs',
  'w:strike',
  'w:color',
  'w:sz',
  'w:szCs',
  'w:u',
  'w:vertAlign',
])

function withoutModeledRPr(raw: unknown): unknown {
  if (typeof raw !== 'string') return raw
  const open = /^<w:rPr(?:\s[^>]*)?>/.exec(raw)?.[0]
  if (!open || !raw.endsWith('</w:rPr>')) return raw
  const inner = raw.slice(open.length, -'</w:rPr>'.length)
  const preserved = splitXmlChildren(inner)
    .filter((child) => !MODELED_RPR_CHILDREN.has(child.name))
    .map((child) => child.xml)
    .join('')
  return `${open}${preserved}</w:rPr>`
}

/** Keep Undo history when save only canonicalized typed run-property XML. */
export function sameBodyAfterManagedRPrRebase(live: PmNode, reopened: PmNode): boolean {
  const normalize = (node: PmNode): PmNode => {
    const json = node.toJSON()
    const visit = (value: typeof json): void => {
      if (value.attrs?.tabStops === '[]') value.attrs.tabStops = null
      // Each-page footnote numbering is calculated from the measured page
      // slices. The parser only knows body order, so its cached display number
      // can differ after save without a change to the native reference.
      if (value.type === 'docNoteRef' && value.attrs) {
        value.attrs.num = null
        // Reopening keeps the reference run's own rPr (superscript is modeled);
        // a freshly inserted reference carries none
        const stripped = withoutModeledRPr(value.attrs.rawRPr)
        value.attrs.rawRPr = stripped === '<w:rPr></w:rPr>' || stripped === '' ? null : stripped
      }
      for (const mark of value.marks ?? []) {
        if (mark.type === 'docTextStyle' && mark.attrs)
          mark.attrs.rawRPr = withoutModeledRPr(mark.attrs.rawRPr)
      }
      for (const child of value.content ?? []) visit(child)
    }
    visit(json)
    return node.type.schema.nodeFromJSON(json)
  }
  return normalize(live).eq(normalize(reopened))
}

/** A TOC result can be byte-identical on disk while its generated PM labels and anchors differ. */
export function sameBodyAfterTocRebase(
  live: PmNode,
  reopened: PmNode,
  reopenedBlocks: readonly Block[],
): boolean {
  if (live.childCount !== reopened.childCount) return false
  const json = live.toJSON()
  const saved = reopened.toJSON()
  let sawGeneratedToc = false
  for (let i = 0; i < live.childCount; i++) {
    const current = json.content?.[i]
    const parsed = saved.content?.[i]
    if (!current || !parsed || current.type !== parsed.type) return false
    const currentAttrs = current.attrs ?? {}
    const parsedAttrs = parsed.attrs ?? {}
    const xml = currentAttrs.genXml
    if (typeof xml === 'string') {
      const kind = currentAttrs.fieldDisplay?.kind
      if (current.type !== 'docProtected' || (kind !== 'tocLine' && kind !== 'pageBreak'))
        return false
      const original = reopenedBlocks.find((block) => block.docxIndex === parsedAttrs.docxIndex)
      const raw = original?.originalXml
      const open = original?.sdtShell?.openXml ?? ''
      const close = original?.sdtShell?.closeXml ?? ''
      if (raw !== xml && raw !== `${open}${xml}${close}`) return false
      current.attrs = parsedAttrs
      sawGeneratedToc = true
    } else {
      currentAttrs.docxIndex = parsedAttrs.docxIndex
      current.attrs = currentAttrs
    }
  }
  return (
    sawGeneratedToc && sameBodyAfterManagedRPrRebase(live.type.schema.nodeFromJSON(json), reopened)
  )
}

/** Re-anchor the current document without inserting a user-visible history step. */
export function rebaseTocAnchors(editor: Editor, reopened: PmNode): void {
  const tr = editor.state.tr
  editor.state.doc.forEach((node, pos, index) => {
    // Keep generated TOC anchors in the history's coordinate system. Reusing a
    // newer index would collide with a later body block after Undo changes the
    // number of TOC entries.
    if (node.type.name === 'docProtected' && node.attrs.fieldDisplay?.kind === 'tocLine') return
    const nextIndex = reopened.child(index)?.attrs.docxIndex
    if (nextIndex !== node.attrs.docxIndex)
      tr.setNodeMarkup(pos, undefined, { ...node.attrs, docxIndex: nextIndex ?? null })
  })
  if (tr.docChanged)
    editor.view.dispatch(tr.setMeta('addToHistory', false).setMeta('preventUpdate', true))
}
