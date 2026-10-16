/** Native presentation-wide slide-show settings in ppt/presProps.xml. */
import type { OpenedPptx } from './index'
import { getCustomShows } from './custom-shows'
import { addContentTypeOverride, appendRelationship } from './notes'
import { resolveTarget } from './zip'

export type ShowMode = 'speaker' | 'browse' | 'kiosk'
export type ShowSelection =
  { kind: 'all' } | { kind: 'range'; start: number; end: number } | { kind: 'custom'; id: string }

export interface ShowSetup {
  mode: ShowMode
  loop: boolean
  showNarration: boolean
  showAnimation: boolean
  useTimings: boolean
  slides: ShowSelection
}

const PRESENTATION = 'ppt/presentation.xml'
const PROPS = 'ppt/presProps.xml'
const REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps'
const CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.presentationml.presProps+xml'
const NS_P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
const DEFAULT: ShowSetup = {
  mode: 'speaker',
  loop: false,
  showNarration: true,
  showAnimation: true,
  useTimings: true,
  slides: { kind: 'all' },
}

function propsPath(opened: OpenedPptx): string | null {
  for (const rel of opened.archive.readRels(PRESENTATION).values()) {
    if (rel.type === REL_TYPE) return resolveTarget(PRESENTATION, rel.target)
  }
  return opened.archive.has(PROPS) ? PROPS : null
}

function boolAttr(tag: string, name: string, fallback: boolean): boolean {
  const raw = new RegExp(`(?:^|\\s)${name}="([^"]*)"`).exec(tag)?.[1]
  return raw == null ? fallback : raw !== '0' && raw !== 'false'
}

function attr(tag: string, name: string): string | null {
  return new RegExp(`(?:^|\\s)${name}="([^"]*)"`).exec(tag)?.[1] ?? null
}

export function getShowSetup(opened: OpenedPptx): ShowSetup {
  const path = propsPath(opened)
  const xml = path ? opened.archive.readText(path) : null
  const raw = xml?.match(/<p:showPr\b[^>]*(?:\/>|>[\s\S]*?<\/p:showPr>)/)?.[0]
  if (!raw) return { ...DEFAULT, slides: { kind: 'all' } }
  const open = /^<p:showPr\b[^>]*>/.exec(raw)?.[0] ?? raw
  const mode: ShowMode = /<p:kiosk\b/.test(raw)
    ? 'kiosk'
    : /<p:browse\b/.test(raw)
      ? 'browse'
      : 'speaker'
  const range = /<p:sldRg\b[^>]*\/>/.exec(raw)?.[0]
  const custom = /<p:custShow\b[^>]*\/>/.exec(raw)?.[0]
  let slides: ShowSelection = { kind: 'all' }
  if (range) {
    const start = Number(attr(range, 'st'))
    const end = Number(attr(range, 'end'))
    if (Number.isInteger(start) && Number.isInteger(end) && start >= 1 && end >= start)
      slides = { kind: 'range', start, end }
  } else if (custom) {
    const id = attr(custom, 'id')
    if (id != null) slides = { kind: 'custom', id }
  }
  return {
    mode,
    loop: boolAttr(open, 'loop', false),
    showNarration: boolAttr(open, 'showNarration', true),
    showAnimation: boolAttr(open, 'showAnimation', true),
    useTimings: boolAttr(open, 'useTimings', true),
    slides,
  }
}

export function canSetShowSetup(opened: OpenedPptx, setup: ShowSetup): boolean {
  if (!setup || !['speaker', 'browse', 'kiosk'].includes(setup.mode)) return false
  if (
    !['loop', 'showNarration', 'showAnimation', 'useTimings'].every(
      (key) => typeof setup[key as keyof ShowSetup] === 'boolean',
    )
  )
    return false
  const slides = setup.slides
  if (!slides) return false
  if (slides.kind === 'all') return true
  if (slides.kind === 'range')
    return (
      Number.isInteger(slides.start) &&
      Number.isInteger(slides.end) &&
      slides.start >= 1 &&
      slides.end >= slides.start &&
      slides.end <= opened.deck.slides.length
    )
  if (slides.kind === 'custom') return getCustomShows(opened).some((show) => show.id === slides.id)
  return false
}

function setAttr(tag: string, name: string, value: string): string {
  const re = new RegExp(`(\\s${name}=")[^"]*(")`)
  return re.test(tag)
    ? tag.replace(re, (_all, before: string, after: string) => `${before}${value}${after}`)
    : tag.replace(/\s*\/?>(?=$)/, (end) => ` ${name}="${value}"${end}`)
}

function patchShowXml(raw: string, setup: ShowSetup): string {
  const selfClosing = /\/>$/.test(raw)
  const opening = /^<p:showPr\b[^>]*>/.exec(raw)?.[0] ?? raw
  let open = opening
  for (const name of ['loop', 'showNarration', 'showAnimation', 'useTimings'] as const)
    open = setAttr(open, name, setup[name] ? '1' : '0')
  if (selfClosing) open = open.replace(/\/>$/, '>')
  let inner = selfClosing ? '' : raw.slice(opening.length, -'</p:showPr>'.length)
  // Replace only the choice elements; pen color, extension lists and unknown children survive.
  inner = inner.replace(/<p:(?:present|browse|kiosk|sldAll|sldRg|custShow)\b[^>]*\/>/g, '')
  const mode =
    setup.mode === 'speaker'
      ? '<p:present/>'
      : setup.mode === 'browse'
        ? '<p:browse/>'
        : '<p:kiosk/>'
  const selection =
    setup.slides.kind === 'range'
      ? `<p:sldRg st="${setup.slides.start}" end="${setup.slides.end}"/>`
      : setup.slides.kind === 'custom'
        ? `<p:custShow id="${setup.slides.id}"/>`
        : '<p:sldAll/>'
  return `${open}${mode}${selection}${inner}</p:showPr>`
}

/** Changes only recognized show controls, retaining unrelated presentation properties. */
export function setShowSetup(opened: OpenedPptx, setup: ShowSetup): boolean {
  if (!canSetShowSetup(opened, setup)) return false
  let path = propsPath(opened)
  if (!path) {
    path = PROPS
    opened.archive.entries.set(
      path,
      Buffer.from(`<p:presentationPr xmlns:p="${NS_P}"></p:presentationPr>`, 'utf8'),
    )
    addContentTypeOverride(opened.archive, path, CONTENT_TYPE)
    appendRelationship(opened.archive, PRESENTATION, REL_TYPE, 'presProps.xml')
  }
  const xml = opened.archive.readText(path)
  if (!xml || !/<p:presentationPr\b/.test(xml)) return false
  const match = /<p:showPr\b[^>]*(?:\/>|>[\s\S]*?<\/p:showPr>)/.exec(xml)
  const newShow = patchShowXml('<p:showPr/>', setup)
  let next: string
  if (match) next = xml.replace(match[0], () => patchShowXml(match[0], setup))
  else if (/<p:presentationPr\b[^>]*\/>/.test(xml))
    next = xml.replace(
      /<p:presentationPr\b[^>]*\/>/,
      (root) => `${root.slice(0, -2)}>${newShow}</p:presentationPr>`,
    )
  else if (/<\/p:presentationPr>/.test(xml))
    next = xml.replace(/<p:extLst\b|<\/p:presentationPr>/, (token) => `${newShow}${token}`)
  else return false
  if (next === xml) return match != null
  opened.archive.entries.set(path, Buffer.from(next, 'utf8'))
  return true
}
