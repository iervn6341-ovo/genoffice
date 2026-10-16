// Header/footer variant resolution and page-number sequences / formats.
import {
  formatNumber,
  type HeaderFooter,
  type HfPartInfo,
  type SectionInfo,
} from '@genoffice/docx-engine'

import type { BlockBox, PageSlice } from './pagination-types'

export type HfVariant = 'default' | 'first' | 'even'

export interface SectionHfRefs {
  header: Partial<Record<HfVariant, string>>
  footer: Partial<Record<HfVariant, string>>
}

/** Effective header/footer refs per section: undefined variants inherit from earlier sections */
export function effectiveHfRefs(sections: SectionInfo[]): SectionHfRefs[] {
  const out: SectionHfRefs[] = []
  let prev: SectionHfRefs = { header: {}, footer: {} }
  for (const s of sections) {
    const cur: SectionHfRefs = {
      header: { ...prev.header, ...s.headerRefs },
      footer: { ...prev.footer, ...s.footerRefs },
    }
    out.push(cur)
    prev = cur
  }
  return out
}

function hfHasContent(hf: HeaderFooter | HfPartInfo | null | undefined): boolean {
  if (!hf) return false
  if ((hf as HeaderFooter).pageNumber || (hf as HfPartInfo).hasPageNumber) return true
  if (hf.text.trim()) return true
  if ((hf as HfPartInfo).images?.length) return true
  return (hf.paras ?? []).some((p) => p.runs.some((r) => r.text.trim()))
}

/**
 * Direct (no-preview) PDF export prints the edit canvas, where the header/footer exists
 * once per document instead of once per page — so any printable header/footer must force
 * the preview-merge export path. Empty parts don't count.
 */
export function hasPrintableHeaderFooter(input: {
  /** local edit state: global header/footer, active variants, per-section edits */
  edited: Array<HeaderFooter | null | undefined>
  sections: SectionInfo[]
  hfParts?: Record<string, HfPartInfo>
  evenOddHf?: boolean
}): boolean {
  if (input.edited.some(hfHasContent)) return true
  const refs = effectiveHfRefs(input.sections)
  return refs.some((ref, i) => {
    const variants: HfVariant[] = ['default']
    if (input.sections[i]?.titlePg) variants.push('first')
    if (input.evenOddHf) variants.push('even')
    return variants.some((v) => {
      const h = ref.header[v]
      const f = ref.footer[v]
      return (
        hfHasContent(h ? input.hfParts?.[h] : null) || hfHasContent(f ? input.hfParts?.[f] : null)
      )
    })
  })
}

/** Header/footer variant a page renders: the titlePg first page of its section
 *  takes 'first', an even page under evenAndOddHeaders takes 'even' */
export function hfVariantOf(
  titlePg: boolean | undefined,
  firstOfSection: boolean,
  evenOddHf: boolean,
  pageNo: number,
): 'first' | 'even' | 'default' {
  if (titlePg && firstOfSection) return 'first'
  if (evenOddHf && pageNo % 2 === 0) return 'even'
  return 'default'
}

/** Displayed page number per page: restart at the section's pgNumType w:start, otherwise continue;
 *  evenPage/oddPage section breaks skip a number to fix parity (when not restarting) */
export function pageNumbers(slices: PageSlice[], sections: SectionInfo[]): number[] {
  const nums: number[] = []
  let n = 0
  let prevSection = -1
  for (const slice of slices) {
    if (slice.section !== prevSection) {
      const sec = sections[slice.section]
      const start = sec?.pageNumberStart
      n = start ?? n + 1
      if (start === undefined && prevSection !== -1) {
        if (sec?.startType === 'evenPage' && n % 2 === 1) n += 1
        if (sec?.startType === 'oddPage' && n % 2 === 0) n += 1
      }
      prevSection = slice.section
    } else {
      n += 1
    }
    nums.push(n)
  }
  return nums
}

/** Display the page number in the section's number format (w:pgNumType w:fmt) */
export function formatPageNumber(
  n: number,
  fmt?: string,
  chapter?: string | null,
  chapterSep?: string,
): string {
  const page = formatNumber(n, fmt ?? 'decimal')
  if (!chapter) return page
  const separator: Record<string, string> = {
    hyphen: '-',
    period: '.',
    colon: ':',
    emDash: '—',
    enDash: '–',
  }
  return `${chapter}${separator[chapterSep ?? 'hyphen'] ?? '-'}${page}`
}

/** The latest numbered Heading at each physical page, for native chapter PAGE display. */
export function chapterNumbersForPages(
  slices: PageSlice[],
  blocks: BlockBox[],
  sections: SectionInfo[],
  headingStyleIds: Map<number, string>,
): Array<string | null> {
  const styles = new Map([...headingStyleIds].map(([level, id]) => [id, level]))
  const latest = new Map<number, string>()
  const out: Array<string | null> = []
  let blockIndex = 0
  for (const slice of slices) {
    while (blockIndex < blocks.length && blocks[blockIndex].top < slice.end) {
      const el = blocks[blockIndex++].el
      const level = el ? styles.get(el.getAttribute('data-style') ?? '') : undefined
      const marker = el?.getAttribute('data-marker')?.replace(/[.\s、．]+$/, '')
      if (level !== undefined && marker) latest.set(level, marker)
    }
    const level = sections[slice.section]?.pageNumberChapterStyle
    out.push(level === undefined ? null : (latest.get(level) ?? null))
  }
  return out
}
