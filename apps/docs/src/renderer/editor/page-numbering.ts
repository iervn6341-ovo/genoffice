import type { Editor } from '@tiptap/core'
import type { Node as PmNode } from '@tiptap/pm/model'
import type { Transaction } from '@tiptap/pm/state'
import type { HeaderFooter, SectionInfo, Run } from '@genoffice/docx-engine'

/** Split at the measured first line of a physical page, retaining the source anchor on its tail. */
export function numberingSectionBreak(
  editor: Editor,
  pos: number,
  sections: SectionInfo[],
  owner: number,
  sectPr: string,
): {
  sections: SectionInfo[]
  target: number
  mutate: (tr: Transaction) => Transaction
} {
  const doc = editor.state.doc
  const $pos = doc.resolve(pos)
  if ($pos.depth > 1 || ($pos.depth === 1 && !$pos.parent.isTextblock))
    throw new Error('pageNumberInsideTable')
  const before = $pos.depth ? $pos.before(1) : pos
  const node = doc.nodeAt(before)
  if (!node) throw new Error('pageNumberBoundaryUnavailable')
  const offset = $pos.depth ? $pos.parentOffset : 0
  let previous = sections[owner].firstBlockIndex - 1
  let next = sections[owner].lastBlockIndex
  doc.forEach((n, at) => {
    const idx = n.attrs.docxIndex
    if (typeof idx !== 'number') return
    if (at < before) previous = Math.max(previous, idx)
    else if (at === before) next = idx
  })
  const boundaryIndex = previous + (next - previous) / 2
  const prefix = {
    ...sections[owner],
    lastBlockIndex: boundaryIndex,
    sectPrXml: sectPr,
    pendingBreak: true as const,
  }
  const suffix = {
    ...sections[owner],
    firstBlockIndex: boundaryIndex + Number.EPSILON,
    startType: 'nextPage' as const,
  }
  const updated = [...sections.slice(0, owner), prefix, suffix, ...sections.slice(owner + 1)]
  const breakNode = editor.schema.nodes.docProtected.create({
    docxIndex: boundaryIndex,
    blockType: 'passthrough',
    label: 'Section break paragraph',
    previewText: '',
    genXml: `<w:p><w:pPr>${sectPr}</w:pPr></w:p>`,
    breakStartType: 'nextPage',
  })
  return {
    sections: updated,
    target: owner + 1,
    mutate: (tr) => {
      const nodes: PmNode[] = []
      if (offset > 0) {
        let content = node.content.cut(0, offset)
        const last = content.lastChild
        if (
          last?.type.name === 'hardBreak' &&
          (last.attrs.type === 'page' || last.attrs.breakType === 'page')
        )
          content = content.cut(0, content.size - last.nodeSize)
        nodes.push(
          node.type.create(
            { ...node.attrs, docxIndex: previous + (boundaryIndex - previous) / 2 },
            content,
            node.marks,
          ),
        )
      }
      nodes.push(breakNode)
      if (offset < node.content.size || offset === 0) {
        const attrs: Record<string, unknown> = { ...node.attrs, pageBreakBefore: false }
        if (attrs.docxIndex == null) attrs.docxIndex = boundaryIndex + (next - boundaryIndex) / 2
        nodes.push(
          node.type.create(attrs, offset ? node.content.cut(offset) : node.content, node.marks),
        )
      }
      return tr.replaceWith(before, before + node.nodeSize, nodes)
    },
  }
}

/** PAGE stays a field marker; other footer content and run formatting remain intact. */
const withoutPageRun = (r: Run): Run => ({
  ...r,
  ...(r.text.includes('\uE001') ? { hfField: undefined } : {}),
  text: r.text.replace(/\uE001/g, ''),
})
export function withoutPageNumber(hf: HeaderFooter | null): HeaderFooter {
  if (!hf) return { text: '', pageNumber: false }
  return {
    ...hf,
    suppressPageNumber: true,
    text: hf.text.replace(/\uE001/g, ''),
    pageNumber: false,
    ...(hf.paras
      ? {
          paras: hf.paras.map((p) => ({
            ...p,
            runs: p.runs.map(withoutPageRun),
            ...(p.cells
              ? {
                  cells: p.cells.map((c) => ({
                    ...c,
                    paras: c.paras.map((rs) => rs.map(withoutPageRun)),
                  })),
                }
              : {}),
          })),
        }
      : {}),
  }
}
export function withPageNumber(hf: HeaderFooter | null): HeaderFooter {
  if (
    hf?.pageNumber ||
    hf?.text.includes('\uE001') ||
    hf?.paras?.some((p) => p.runs.some((r) => r.text.includes('\uE001')))
  )
    return hf
  const para = { align: 'center' as const, runs: [{ text: '\uE001' }] }
  return {
    ...(hf ?? {}),
    suppressPageNumber: false,
    text: [hf?.text, '\uE001'].filter(Boolean).join('\n'),
    pageNumber: true,
    paras: [...(hf?.paras ?? (hf?.text ? [{ runs: [{ text: hf.text }] }] : [])), para],
  }
}
