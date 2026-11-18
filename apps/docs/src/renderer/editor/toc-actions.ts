import type { Editor } from '@tiptap/core'
import { carrySavedTocHistorySource } from '../doc-dirty'
import type { Node as PmNode } from '@tiptap/pm/model'
import { closeHistory } from '@tiptap/pm/history'
import type {
  Block,
  FieldDisplay,
  NumberingDef,
  StyleInfo,
  TocEntry,
  TocOptions,
} from '@genoffice/docx-engine'
import {
  computeListMarkerInfos,
  decodeEntities,
  patchTocCachedPageNumber,
  tocInstructionFromXml,
  type ListItemRef,
} from '@genoffice/docx-engine'

export interface TocSourceEntry extends TocEntry {
  pos: number
  anchor?: string
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

  // Precompute list markers for docListItems so numbered headings retain their markers (e.g. "第1章")
  const refs: ListItemRef[] = []
  const listItems: Array<{ pos: number; node: PmNode }> = []
  editor.state.doc.descendants((node, pos) => {
    if (node.type.name === 'docListItem') {
      const deleted =
        !!node.attrs.paraMarkDel ||
        (node.attrs.blockRevision as { kind?: string } | null)?.kind === 'del'
      refs.push({
        numId: (node.attrs.numId as string | null) ?? null,
        ilvl: Number(node.attrs.ilvl) || 0,
        ...(deleted ? { deleted: true } : {}),
      })
      listItems.push({ pos, node })
      return false
    }
    return true
  })
  const numberingDefs =
    (editor.storage.listNumbering as { defs?: Map<string, NumberingDef> } | undefined)?.defs
  const markerMap = new Map<number, string>()
  if (refs.length > 0 && numberingDefs && numberingDefs.size > 0) {
    const markers = computeListMarkerInfos(refs, numberingDefs)
    markers.forEach((m, idx) => {
      if (m?.text) markerMap.set(listItems[idx].pos, m.text)
    })
  }

  editor.state.doc.forEach((node, pos) => {
    const styleId = typeof node.attrs.styleId === 'string' ? node.attrs.styleId : ''
    const styleInfo = styles?.get(styleId)
    const name = styleInfo?.name ?? styleId
    const mapped = options.styleLevels[name] ?? options.styleLevels[styleId]
    const heading = node.type.name === 'docHeading'
    const isListItem = node.type.name === 'docListItem'
    const isParagraph = node.type.name === 'docParagraph'
    const text = node.textContent.trim()
    if (text && (heading || isListItem || isParagraph)) {
      const headingLevel = Number(node.attrs.level) || styleInfo?.headingLevel || 0
      const sourceXml = tocNodeXml(node as never, blocks)
      const directOutlineMatch = /<w:outlineLvl\s+w:val="([0-9])"/.exec(sourceXml)
      const directOutline = directOutlineMatch ? Number(directOutlineMatch[1]) + 1 : undefined
      let nativeLevel = 0
      if (options.useHeadingStyles && headingLevel > 0) {
        nativeLevel = headingLevel
      } else if (options.useOutlineLevels && directOutline !== undefined) {
        nativeLevel = directOutline
      } else if (heading && headingLevel > 0) {
        nativeLevel = headingLevel
      }
      const level = mapped ?? nativeLevel
      if (level >= 1 && level <= options.levels) {
        const num = isListItem ? markerMap.get(pos) : undefined
        const hiddenBms = Array.isArray(node.attrs.hiddenBookmarks)
          ? (node.attrs.hiddenBookmarks as string[])
          : []
        const regularBms = Array.isArray(node.attrs.bookmarks)
          ? (node.attrs.bookmarks as string[])
          : []
        let anchor = [...hiddenBms, ...regularBms].find((b) => b.startsWith('_Toc'))
        if (!anchor && (hiddenBms.length > 0 || regularBms.length > 0)) {
          anchor = hiddenBms[0] ?? regularBms[0]
        }
        if (!anchor) {
          anchor = `_Toc${Math.floor(100000000 + Math.random() * 900000000)}`
          const nextHidden = [...hiddenBms, anchor]
          try {
            editor.view.dispatch(
              editor.state.tr
                .setNodeMarkup(pos, undefined, {
                  ...node.attrs,
                  hiddenBookmarks: nextHidden,
                })
                .setMeta('addToHistory', false)
                .setMeta('preventUpdate', true),
            )
          } catch {
            // ignore dispatch failure
          }
        }
        entries.push({ level, text, num, pos, ...(anchor ? { anchor } : {}) })
      }
    }
    if (options.captionLabel) {
      const xml = tocNodeXml(node as never, blocks)
      const seqPattern = new RegExp(`(?:^|\\s)SEQ\\s+${options.captionLabel}(?:\\s|\\\\|$)`, 'i')
      if (seqPattern.test(xml)) {
        const hiddenBms = Array.isArray(node.attrs.hiddenBookmarks)
          ? (node.attrs.hiddenBookmarks as string[])
          : []
        const regularBms = Array.isArray(node.attrs.bookmarks)
          ? (node.attrs.bookmarks as string[])
          : []
        let anchor = [...hiddenBms, ...regularBms].find((b) => b.startsWith('_Toc'))
        if (!anchor && (hiddenBms.length > 0 || regularBms.length > 0)) {
          anchor = hiddenBms[0] ?? regularBms[0]
        }
        entries.push({ level: 1, text, pos, ...(anchor ? { anchor } : {}) })
      }
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

/** Locate the complete native TOC field (preferring one around targetPos, if given). */
export function findTocField(
  editor: Editor,
  blocks: Block[],
  targetPos?: number,
): TocFieldRegion | null {
  const regions: TocFieldRegion[] = []
  let nodes: TocFieldRegion['nodes'] = []
  let from = -1
  let depth = 0
  let instruction = ''
  let targetSdtGroup: number | undefined = undefined
  let inToc = false
  let keepPageBreak = false

  editor.state.doc.forEach((node, pos) => {
    const xml = tocNodeXml(node as never, blocks)
    const code = tocInstructionFromXml(xml)
    const block = blocks.find((b) => b.docxIndex === node.attrs.docxIndex)
    const sdtGroup = block?.sdtShell?.group
    const hasCloseXml = !!block?.sdtShell?.closeXml
    const isTocLine = (node.attrs.fieldDisplay as FieldDisplay | undefined)?.kind === 'tocLine'
    const begins = (xml.match(/w:fldCharType="begin"/g) ?? []).length
    const ends = (xml.match(/w:fldCharType="end"/g) ?? []).length

    if (!inToc) {
      if (code) {
        inToc = true
        from = pos
        instruction = code
        depth = begins - ends
        targetSdtGroup = sdtGroup
        nodes = [{ pos, node, xml }]
        keepPageBreak = /<w:br\s[^>]*w:type="page"/.test(xml)
      }
    } else {
      depth += begins - ends
      const inSameSdt = targetSdtGroup !== undefined && sdtGroup === targetSdtGroup
      const isContinuation = depth > 0 || inSameSdt || isTocLine
      if (isContinuation) {
        nodes.push({ pos, node, xml })
        if (/<w:br\s[^>]*w:type="page"/.test(xml)) keepPageBreak = true
        if (targetSdtGroup !== undefined && hasCloseXml) {
          regions.push({
            from,
            to: pos + node.nodeSize,
            nodes,
            instruction,
            keepPageBreak,
          })
          inToc = false
          from = -1
          depth = 0
          instruction = ''
          targetSdtGroup = undefined
          nodes = []
          keepPageBreak = false
        }
      } else {
        const lastNode = nodes[nodes.length - 1]
        regions.push({
          from,
          to: lastNode ? lastNode.pos + lastNode.node.nodeSize : pos,
          nodes,
          instruction,
          keepPageBreak,
        })
        inToc = false
        from = -1
        depth = 0
        instruction = ''
        targetSdtGroup = undefined
        nodes = []
        keepPageBreak = false

        if (code) {
          inToc = true
          from = pos
          instruction = code
          depth = begins - ends
          targetSdtGroup = sdtGroup
          nodes = [{ pos, node, xml }]
          keepPageBreak = /<w:br\s[^>]*w:type="page"/.test(xml)
        }
      }
    }
  })

  if (inToc && nodes.length > 0) {
    const lastNode = nodes[nodes.length - 1]
    regions.push({
      from,
      to: lastNode ? lastNode.pos + lastNode.node.nodeSize : from,
      nodes,
      instruction,
      keepPageBreak,
    })
  }

  if (regions.length === 0) return null
  if (targetPos !== undefined) {
    const matching = regions.find((r) => targetPos >= r.from && targetPos <= r.to)
    if (matching) return matching
  }
  return regions[0]
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
    carrySavedTocHistorySource(editor, node, inner)
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
