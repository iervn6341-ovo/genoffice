import type { Editor } from '@tiptap/core'
import type { Node as PmNode } from '@tiptap/pm/model'
import { closeHistory } from '@tiptap/pm/history'
import type { Block, StyleInfo, TocEntry, TocOptions } from '@genoffice/docx-engine'
import {
  decodeEntities,
  patchTocCachedPageNumber,
  tocInstructionFromXml,
} from '@genoffice/docx-engine'

export interface TocSourceEntry extends TocEntry {
  pos: number
}

/** Original bytes or the latest generated fragment for a protected block. */
export function tocNodeXml(node: { attrs: Record<string, unknown> }, blocks: Block[]): string {
  if (node.attrs.genXml) return String(node.attrs.genXml)
  const idx = node.attrs.docxIndex
  if (idx === null || idx === undefined) return ''
  return blocks.find((block) => block.docxIndex === idx)?.originalXml ?? ''
}

function tcEntry(xml: string): { text: string; level: number } | null {
  const complex = [...xml.matchAll(/<w:instrText\b[^>]*>([\s\S]*?)<\/w:instrText>/g)]
    .map((part) => decodeEntities(part[1] ?? ''))
    .join('')
  const simple = /<w:fldSimple\b[^>]*\bw:instr="([^"]*)"/.exec(xml)?.[1]
  const instruction = complex || decodeEntities(simple ?? '')
  const match = /(?:^|\s)TC\s+"([^"]+)"/i.exec(instruction)
  if (!match) return null
  const level = Number(/\\l\s+([1-9])\b/i.exec(instruction)?.[1] ?? '1')
  return { text: match[1]!, level }
}

/** Current document entries from selected sources, preserving document order. */
export function collectTocSources(
  editor: Editor,
  blocks: Block[],
  styles: ReadonlyMap<string, StyleInfo> | undefined,
  options: TocOptions,
): TocSourceEntry[] {
  const entries: TocSourceEntry[] = []
  editor.state.doc.forEach((node, pos) => {
    const styleId = typeof node.attrs.styleId === 'string' ? node.attrs.styleId : ''
    const name = styles?.get(styleId)?.name ?? styleId
    const mapped = options.styleLevels[name] ?? options.styleLevels[styleId]
    const heading = node.type.name === 'docHeading'
    const text = node.textContent.trim()
    if (text && (heading || node.type.name === 'docParagraph')) {
      const headingLevel = Number(node.attrs.level) || styles?.get(styleId)?.headingLevel || 0
      const sourceXml = tocNodeXml(node as never, blocks)
      const directOutline = heading && !styleId && /<w:outlineLvl\b/.test(sourceXml)
      const nativeLevel =
        heading &&
        ((options.useHeadingStyles && !directOutline) ||
          (options.useOutlineLevels && directOutline))
          ? headingLevel
          : 0
      const level = mapped ?? nativeLevel
      if (level >= 1 && level <= options.levels) entries.push({ level, text, pos })
    }
    if (options.useTcFields) {
      const tc = tcEntry(tocNodeXml(node as never, blocks))
      if (tc && tc.level <= options.levels) entries.push({ ...tc, pos })
    }
  })
  return entries
}

export interface TocFieldRegion {
  from: number
  to: number
  nodes: Array<{ pos: number; node: PmNode; xml: string }>
  instruction: string
  keepPageBreak: boolean
}

/** Reuse the TOC field's original boundary anchors so its surrounding SDT stays balanced. */
export function anchorTocReplacement(
  nodes: Array<Record<string, unknown>>,
  region: TocFieldRegion,
  blocks: Block[],
): boolean {
  const original = region.nodes
    .map(({ node }) => node.attrs.docxIndex)
    .filter(
      (index): index is number =>
        typeof index === 'number' && Number.isInteger(index) && index >= 0,
    )
  if (original.length === 0 || nodes.length === 0) return true
  const groups = new Set(
    original
      .map((index) => blocks.find((block) => block.docxIndex === index)?.sdtShell?.group)
      .filter((group) => group !== undefined),
  )
  if (original.length === 1 && nodes.length > 1 && groups.size > 0) return false
  const assign = (nodeIndex: number, originalIndex: number) => {
    const attrs = nodes[nodeIndex]!.attrs as Record<string, unknown>
    attrs.docxIndex = original[originalIndex]
  }
  assign(0, 0)
  if (nodes.length > 1 && original.length > 1) assign(nodes.length - 1, original.length - 1)
  const middle = Math.min(nodes.length - 2, original.length - 2)
  for (let i = 0; i < middle; i++) assign(i + 1, i + 1)
  return true
}

/** Locate the first complete native TOC field; nested page-reference fields do not end it early. */
export function findTocField(editor: Editor, blocks: Block[]): TocFieldRegion | null {
  const nodes: TocFieldRegion['nodes'] = []
  let from = -1
  let depth = 0
  let instruction = ''
  let found: TocFieldRegion | null = null
  editor.state.doc.forEach((node, pos) => {
    if (found) return
    const xml = tocNodeXml(node as never, blocks)
    if (from < 0) {
      const code = tocInstructionFromXml(xml)
      if (!code) return
      instruction = code
      from = pos
    }
    nodes.push({ pos, node, xml })
    depth += (xml.match(/w:fldCharType="begin"/g) ?? []).length
    depth -= (xml.match(/w:fldCharType="end"/g) ?? []).length
    if (depth <= 0) {
      found = {
        from,
        to: pos + node.nodeSize,
        nodes,
        instruction,
        keepPageBreak: /<w:br\s[^>]*w:type="page"/.test(xml),
      }
    }
  })
  return found
}

/** Keep the pre-update field bytes in Undo history after Save rebases its source blocks. */
function innerTocXml(xml: string, docxIndex: unknown, blocks: Block[]): string {
  const block = blocks.find((candidate) => candidate.docxIndex === docxIndex)
  const open = block?.sdtShell?.openXml ?? ''
  const close = block?.sdtShell?.closeXml ?? ''
  let inner = xml
  if (open && inner.startsWith(open)) inner = inner.slice(open.length)
  if (close && inner.endsWith(close)) inner = inner.slice(0, -close.length)
  return inner
}

export function preserveTocHistorySource(
  editor: Editor,
  region: TocFieldRegion,
  blocks: Block[],
): void {
  const tr = editor.state.tr
  for (const { pos, node, xml } of region.nodes) {
    const inner = innerTocXml(xml, node.attrs.docxIndex, blocks)
    if (!inner || node.attrs.genXml === inner) continue
    tr.setNodeMarkup(pos, undefined, { ...node.attrs, genXml: inner })
  }
  if (tr.docChanged)
    editor.view.dispatch(tr.setMeta('addToHistory', false).setMeta('preventUpdate', true))
}

/** Page-only update keeps every paragraph, field instruction, and result style intact. */
export function updateTocPageCaches(
  editor: Editor,
  region: TocFieldRegion,
  pages: readonly string[],
  blocks: Block[],
): boolean {
  const candidates = region.nodes.filter(({ node }) => node.attrs.fieldDisplay?.kind === 'tocLine')
  if (candidates.length !== pages.length) return false
  const patched = candidates.map(({ node, xml }, index) => {
    const field = node.attrs.fieldDisplay as { right?: string }
    const next = patchTocCachedPageNumber(xml, field.right ?? '', pages[index]!)
    return next === null ? null : { xml: next, right: pages[index]! }
  })
  if (patched.some((item) => item === null)) return false
  preserveTocHistorySource(editor, region, blocks)
  editor.view.dispatch(closeHistory(editor.state.tr))
  const tr = editor.state.tr
  candidates.forEach(({ node, pos }, index) => {
    const next = patched[index]!
    if (next.xml === region.nodes.find((candidate) => candidate.pos === pos)!.xml) return
    tr.setNodeMarkup(pos, undefined, {
      ...node.attrs,
      genXml: next.xml,
      fieldDisplay: { ...node.attrs.fieldDisplay, right: next.right },
    })
  })
  if (tr.docChanged) {
    editor.view.dispatch(tr)
    editor.view.dispatch(closeHistory(editor.state.tr))
  }
  return true
}
