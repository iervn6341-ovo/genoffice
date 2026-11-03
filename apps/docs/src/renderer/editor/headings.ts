import type { Node as PmNode } from '@tiptap/pm/model'

export interface HeadingRef {
  text: string
  level: number
  /** top-level position of the docHeading node */
  pos: number
}

/** Single heading predicate shared by TOC, nav pane and TOC page backfill (document order) */
export function collectHeadings(doc: PmNode): HeadingRef[] {
  const out: HeadingRef[] = []
  doc.forEach((node, offset) => {
    if (node.type.name === 'docHeading' && node.textContent.trim()) {
      out.push({ text: node.textContent, level: Number(node.attrs.level) || 1, pos: offset })
    }
  })
  return out
}

export interface OutlineItem {
  id: string
  text: string
  page?: string
  level: number
  pos: number
  targetPos?: number
  anchor?: string
  isHeading: boolean
  isToc: boolean
  index: number
}

/** Collects outline items by prioritizing document Table of Contents (TOC) entries, falling back to headings */
export function collectOutlineItems(doc: PmNode): {
  items: OutlineItem[]
} {
  const bookmarkToPos = new Map<string, number>()
  const titleToPos = new Map<string, number>()

  // Pre-index bookmarks and titles across the document to resolve TOC jump targets
  doc.forEach((node, offset) => {
    if (Array.isArray(node.attrs?.bookmarks)) {
      for (const b of node.attrs.bookmarks) {
        if (typeof b === 'string') bookmarkToPos.set(b, offset)
      }
    }
    if (Array.isArray(node.attrs?.hiddenBookmarks)) {
      for (const b of node.attrs.hiddenBookmarks) {
        if (typeof b === 'string') bookmarkToPos.set(b, offset)
      }
    }
    const clean = node.textContent.trim().replace(/\s+/g, '')
    if (clean && !titleToPos.has(clean)) {
      titleToPos.set(clean, offset)
    }
  })

  // 1. Try collecting Table of Contents items from TOC fields or TOC styles
  const tocItems: OutlineItem[] = []
  doc.forEach((node, offset) => {
    const field = node.attrs?.fieldDisplay
    const isTocLine = field?.kind === 'tocLine'
    const styleId = String(node.attrs?.styleId || '')
    const isTocStyle =
      /^(toc\s*\d+|目錄\s*\d+|TOC\d+|1[1-9]|2[1-9]|3[1-9])$/i.test(styleId) && !isTocLine

    if (isTocLine) {
      const baseText = String(field.left ?? node.textContent).trim()
      const numPrefix = field.num ? String(field.num).trim() : ''
      const title = numPrefix ? `${numPrefix} ${baseText}` : baseText
      if (!title) return

      const anchor = field.anchor ? String(field.anchor) : undefined
      let targetPos: number | undefined = undefined
      if (anchor && bookmarkToPos.has(anchor)) {
        targetPos = bookmarkToPos.get(anchor)
      } else {
        const cleanTitle = baseText.replace(/\s+/g, '')
        targetPos = titleToPos.get(cleanTitle)
      }

      tocItems.push({
        id: `toc-${offset}`,
        text: title,
        page: field.right ? String(field.right).trim() : undefined,
        level: Math.min(Math.max(Number(field.level) || 1, 1), 4),
        pos: offset,
        targetPos: targetPos ?? offset,
        anchor,
        isHeading: true,
        isToc: true,
        index: tocItems.length + 1,
      })
      return
    }

    if (isTocStyle) {
      const text = node.textContent.trim()
      if (!text) return
      // Match optional trailing page number like "... 1" or tab-delimited numbers
      const pageMatch = text.match(/[\.\s\t]+(\d+|[ivxlcdm]+)$/i)
      const cleanText = pageMatch ? text.slice(0, pageMatch.index).trim() : text
      const page = pageMatch ? pageMatch[1] : undefined

      const levelMatch = styleId.match(/\d+/)
      const rawLevel = levelMatch ? parseInt(levelMatch[0], 10) : 1
      const level = rawLevel > 10 ? Math.floor(rawLevel / 10) : rawLevel

      tocItems.push({
        id: `toc-${offset}`,
        text: cleanText,
        page,
        level: Math.min(Math.max(level, 1), 4),
        pos: offset,
        targetPos: titleToPos.get(cleanText.replace(/\s+/g, '')) ?? offset,
        isHeading: true,
        isToc: true,
        index: tocItems.length + 1,
      })
    }
  })

  if (tocItems.length > 0) {
    return { items: tocItems }
  }

  // 2. Fallback to headings in document body if no TOC exists
  const headings: OutlineItem[] = []
  doc.forEach((node, offset) => {
    const rawText = node.textContent.trim()
    if (!rawText) return

    if (node.type.name === 'docHeading') {
      const level = Number(node.attrs.level) || 1
      headings.push({
        id: `h-${offset}`,
        text: rawText,
        level: Math.min(Math.max(level, 1), 4),
        pos: offset,
        targetPos: offset,
        isHeading: true,
        isToc: false,
        index: headings.length + 1,
      })
      return
    }

    if (node.type.name === 'docParagraph' || node.type.name === 'docListItem') {
      const styleId = String(node.attrs?.styleId || '')
      const isHeadingStyle = /heading|標題|title|subtitle/i.test(styleId)
      const numMatch = rawText.match(
        /^(?:(\d+(\.\d+)*)[\.、\s]|[一二三四五六七八九十]+[、\.]|第[一二三四五六七八九十\d]+[章節點]|Chapter\s+\d+|Section\s+\d+)/i,
      )
      const isShortHeadingCandidate = Boolean(numMatch && rawText.length <= 120)

      let level = 1
      if (node.attrs?.ilvl !== undefined && node.attrs?.ilvl !== null) {
        level = Number(node.attrs.ilvl) + 1
      } else if (numMatch && numMatch[1]) {
        level = numMatch[1].split('.').length
      }

      if (isHeadingStyle || isShortHeadingCandidate) {
        headings.push({
          id: `h-${offset}`,
          text: rawText,
          level: Math.min(Math.max(level, 1), 4),
          pos: offset,
          targetPos: offset,
          isHeading: true,
          isToc: false,
          index: headings.length + 1,
        })
      }
    }
  })

  return { items: headings }
}
