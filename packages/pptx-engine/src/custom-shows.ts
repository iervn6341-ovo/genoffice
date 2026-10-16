/** Native presentation.xml custom shows. Slide references are presentation relationship IDs. */
import type { OpenedPptx } from './index'
import { resolveTarget } from './zip'
import { escapeXmlAttr } from './xml-utils'
import { unescapeXml } from './notes'

export interface CustomShowInfo {
  /** OOXML unsigned integer, kept as a string to preserve imported identity. */
  id: string
  name: string
  /** Deck indices in playback order; duplicates are permitted by the native list. */
  slideIndices: number[]
}

const PRES_PATH = 'ppt/presentation.xml'
const LIST_RE = /<p:custShowLst\b[^>]*>([\s\S]*?)<\/p:custShowLst>/
const SHOW_RE = /<p:custShow\b([^>]*?)>([\s\S]*?)<\/p:custShow>/g

function replaceSlideRefs(inner: string, refs: string): string {
  const first = /<p:sld\b[^>]*\/>/.exec(inner)
  return first
    ? inner.slice(0, first.index) + refs + inner.slice(first.index).replace(/<p:sld\b[^>]*\/>/g, '')
    : refs + inner
}

function attr(xml: string, name: string): string | null {
  const match = new RegExp(`(?:^|\\s)${name}="([^"]*)"`).exec(xml)
  return match ? unescapeXml(match[1]!) : null
}

function slideRids(opened: OpenedPptx): Map<number, string> {
  const pres = opened.archive.readText(PRES_PATH) ?? ''
  const list = /<p:sldIdLst\b[^>]*>([\s\S]*?)<\/p:sldIdLst>/.exec(pres)?.[1] ?? ''
  const rels = opened.archive.readRels(PRES_PATH)
  const byPath = new Map(opened.deck.slides.map((slide, index) => [slide.path, index]))
  const out = new Map<number, string>()
  for (const match of list.matchAll(/<p:sldId\b[^>]*\/>/g)) {
    const rid = attr(match[0], 'r:id')
    const rel = rid ? rels.get(rid) : null
    const path = rel ? resolveTarget(PRES_PATH, rel.target) : null
    const index = path ? byPath.get(path) : undefined
    if (index != null && rid) out.set(index, rid)
  }
  return out
}

function parseShow(raw: string, byRid: Map<string, number>): CustomShowInfo {
  const opening = /^<p:custShow\b([^>]*)>/.exec(raw)?.[1] ?? ''
  const slideIndices: number[] = []
  for (const match of raw.matchAll(/<p:sld\b[^>]*\/>/g)) {
    const index = byRid.get(attr(match[0], 'r:id') ?? '')
    if (index != null) slideIndices.push(index)
  }
  return { id: attr(opening, 'id') ?? '', name: attr(opening, 'name') ?? '', slideIndices }
}

export function getCustomShows(opened: OpenedPptx): CustomShowInfo[] {
  const pres = opened.archive.readText(PRES_PATH) ?? ''
  const list = LIST_RE.exec(pres)?.[1] ?? ''
  const byRid = new Map([...slideRids(opened)].map(([index, rid]) => [rid, index]))
  return [...list.matchAll(SHOW_RE)].map((match) => parseShow(match[0], byRid))
}

function validShow(show: CustomShowInfo, slideCount: number): boolean {
  return (
    show != null &&
    typeof show === 'object' &&
    typeof show.id === 'string' &&
    /^(?:0|[1-9]\d*)$/.test(show.id) &&
    Number(show.id) <= 0xffffffff &&
    typeof show.name === 'string' &&
    show.name.length > 0 &&
    Array.isArray(show.slideIndices) &&
    show.slideIndices.every((index) => Number.isInteger(index) && index >= 0 && index < slideCount)
  )
}

/** Validate the native ID range and every slide reference before starting a transaction. */
export function canSetCustomShows(opened: OpenedPptx, shows: CustomShowInfo[]): boolean {
  if (!Array.isArray(shows)) return false
  const ids = new Set<string>()
  for (const show of shows) {
    if (!validShow(show, opened.deck.slides.length) || ids.has(show.id)) return false
    ids.add(show.id)
  }
  if (!opened.archive.readText(PRES_PATH)) return false
  const ridByIndex = slideRids(opened)
  if (shows.some((show) => show.slideIndices.some((index) => !ridByIndex.has(index)))) return false
  return true
}

/** Returns false for an invalid edit; leaves every non-custom-show XML part unchanged. */
export function setCustomShows(opened: OpenedPptx, shows: CustomShowInfo[]): boolean {
  if (!canSetCustomShows(opened, shows)) return false
  const pres = opened.archive.readText(PRES_PATH)!
  const ridByIndex = slideRids(opened)

  const oldList = LIST_RE.exec(pres)
  const byRid = new Map([...ridByIndex].map(([index, rid]) => [rid, index]))
  const oldById = new Map<string, string>()
  for (const match of (oldList?.[1] ?? '').matchAll(SHOW_RE))
    oldById.set(attr(match[1] ?? '', 'id') ?? '', match[0])
  const showXml = shows
    .map((show) => {
      const old = oldById.get(show.id)
      // An untouched imported show retains its exact XML, including unknown children and dangling refs.
      if (old && JSON.stringify(parseShow(old, byRid)) === JSON.stringify(show)) return old
      const slides = show.slideIndices
        .map((index) => `<p:sld r:id="${escapeXmlAttr(ridByIndex.get(index)!)}"/>`)
        .join('')
      if (old) {
        const opening = /^<p:custShow\b[^>]*>/.exec(old)?.[0] ?? '<p:custShow>'
        let updated = opening.replace(/\bname="[^"]*"/, `name="${escapeXmlAttr(show.name)}"`)
        if (!/\bname=/.test(opening))
          updated = updated.replace(/>$/, ` name="${escapeXmlAttr(show.name)}">`)
        const inner = old.slice(opening.length, -'</p:custShow>'.length)
        const nextInner = /<p:sldLst\b[^>]*>[\s\S]*?<\/p:sldLst>/.test(inner)
          ? inner.replace(
              /(<p:sldLst\b[^>]*>)([\s\S]*?)(<\/p:sldLst>)/,
              (_all, open: string, body: string, close: string) =>
                `${open}${replaceSlideRefs(body, slides)}${close}`,
            )
          : /<p:sldLst\b[^>]*\/>/.test(inner)
            ? inner.replace(/<p:sldLst\b[^>]*\/>/, `<p:sldLst>${slides}</p:sldLst>`)
            : `<p:sldLst>${slides}</p:sldLst>${inner}`
        return `${updated}${nextInner}</p:custShow>`
      }
      return `<p:custShow name="${escapeXmlAttr(show.name)}" id="${show.id}"><p:sldLst>${slides}</p:sldLst></p:custShow>`
    })
    .join('')

  let next: string
  if (oldList) {
    const openTag = /^<p:custShowLst\b[^>]*>/.exec(oldList[0])![0]
    const oldInner = oldList[1]!
    // Keep unknown list-level XML rather than erasing extensions on a list rewrite.
    const extras = oldInner.replace(SHOW_RE, '')
    const replacement =
      shows.length === 0 && extras.trim() === ''
        ? ''
        : `${openTag}${showXml}${extras}</p:custShowLst>`
    next = pres.replace(oldList[0], () => replacement)
  } else if (/<p:custShowLst\b[^>]*\/>/.test(pres)) {
    next = pres.replace(
      /<p:custShowLst\b([^>]*)\/>/,
      (_all, attrs: string) => `<p:custShowLst${attrs}>${showXml}</p:custShowLst>`,
    )
  } else if (shows.length) {
    const insertion = `<p:custShowLst>${showXml}</p:custShowLst>`
    // ECMA-376 presentation child order: after embedded fonts, before photo album/show properties.
    const before =
      /<p:(?:photoAlbum|custDataLst|kinsoku|defaultTextStyle|modifyVerifier|showPr|extLst)\b|<\/p:presentation>/.exec(
        pres,
      )
    if (!before) return false
    next = pres.slice(0, before.index) + insertion + pres.slice(before.index)
  } else return true
  if (next !== pres) opened.archive.entries.set(PRES_PATH, Buffer.from(next, 'utf8'))
  return true
}
