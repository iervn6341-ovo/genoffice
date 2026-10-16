/**
 * Picture accessibility and lock metadata (p:cNvPr descr / decorative flag and
 * a:picLocks noChangeAspect). Edits are sparse string patches of the element
 * fragment: every other attribute and child, including a16:creationId inside
 * the same extLst, stays untouched.
 */
import type { Slide } from './types'
import { patchedElementXml } from './index'

const DECORATIVE_URI = '{C183D7F6-B498-43B3-948B-1728B52AA6E4}'
const DECORATIVE_NS = 'http://schemas.microsoft.com/office/drawing/2017/decorative'
const INK_PREFIX = 'aislides-ink'

function escapeAttr(s: string): string {
  return s
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/\r?\n/g, '&#10;')
}

const DECORATIVE_EXT = `<a:ext uri="${DECORATIVE_URI}"><adec:decorative xmlns:adec="${DECORATIVE_NS}" val="1"/></a:ext>`
const DECORATIVE_RE = new RegExp(
  `<a:ext\\b[^>]*uri="${DECORATIVE_URI.replace(/[{}]/g, '\\$&')}"[^>]*>[\\s\\S]*?</a:ext>`,
)

export interface PictureAltTextState {
  text: string
  decorative: boolean
}

/** Alt text of a picture fragment (decoded), for tests and the UI. */
export function readPictureAltText(xml: string): PictureAltTextState {
  const cNvPr = /<p:cNvPr\b[^>]*>/.exec(xml)?.[0] ?? ''
  const raw = /\bdescr="([^"]*)"/.exec(cNvPr)?.[1] ?? ''
  const text = raw
    .replace(/&#10;|&#xA;/gi, '\n')
    .replace(/&quot;/g, '"')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&amp;/g, '&')
  return { text, decorative: DECORATIVE_RE.test(xml) }
}

function commit(slide: Slide, el: Slide['elements'][number], xml: string): void {
  el.dirty = el.dirtyTransform = el.dirtyFill = el.dirtyStroke = false
  el.dirtyPPr = undefined
  el.anchor.originalXml = xml
  slide.structureDirty = true
}

function editablePicture(slide: Slide, elementId: string) {
  const el = slide.elements.find((e) => e.id === elementId)
  // own ink strokes keep their vector payload in descr; media poster frames are not alt-text targets here
  if (!el || el.type !== 'picture' || el.name?.startsWith(INK_PREFIX)) return null
  return el
}

/**
 * Set picture alt text (null/empty clears) and the decorative flag. A decorative
 * picture carries no description, matching PowerPoint's "Mark as decorative".
 */
export function setPictureAltText(
  slide: Slide,
  elementId: string,
  text: string | null,
  decorative = false,
): boolean {
  const el = editablePicture(slide, elementId)
  if (!el) return false
  let xml = patchedElementXml(el)
  const open = /<p:cNvPr\b((?:"[^"]*"|'[^']*'|[^"'>])*?)(\/?)>/.exec(xml)
  if (!open) return false
  let attrs = open[1]!.replace(/\sdescr="[^"]*"/, '')
  const value = decorative ? '' : (text ?? '')
  if (value) attrs += ` descr="${escapeAttr(value)}"`
  const selfClosing = open[2] === '/'
  const afterOpen = open.index + open[0].length
  let rest = xml.slice(afterOpen)
  let body = ''
  if (selfClosing) {
    if (decorative) body = `<a:extLst>${DECORATIVE_EXT}</a:extLst>`
    xml =
      xml.slice(0, open.index) +
      (body ? `<p:cNvPr${attrs}>${body}</p:cNvPr>` : `<p:cNvPr${attrs}/>`) +
      rest
  } else {
    const close = rest.indexOf('</p:cNvPr>')
    if (close < 0) return false
    let inner = rest.slice(0, close).replace(DECORATIVE_RE, '')
    inner = inner.replace(/<a:extLst>\s*<\/a:extLst>/, '')
    if (decorative) {
      inner = /<a:extLst>/.test(inner)
        ? inner.replace('</a:extLst>', `${DECORATIVE_EXT}</a:extLst>`)
        : `${inner}<a:extLst>${DECORATIVE_EXT}</a:extLst>`
    }
    rest = inner + rest.slice(close)
    xml = xml.slice(0, open.index) + `<p:cNvPr${attrs}>` + rest
  }
  commit(slide, el, xml)
  const model = el as { descr?: string }
  if (value) model.descr = value
  else delete model.descr
  return true
}

/** Persist a picture's "Lock aspect ratio" (a:picLocks noChangeAspect). */
export function setPictureAspectLock(slide: Slide, elementId: string, locked: boolean): boolean {
  const el = editablePicture(slide, elementId)
  if (!el) return false
  let xml = patchedElementXml(el)
  const locks = /<a:picLocks\b([^>]*?)(\/?)>/.exec(xml)
  if (locks) {
    // an explicit "0" keeps an unlocked picture unlocked (an absent attribute reads as locked in the UI)
    const attrs = ` noChangeAspect="${locked ? 1 : 0}"${locks[1]!.replace(/\snoChangeAspect="[^"]*"/, '')}`
    xml =
      xml.slice(0, locks.index) +
      `<a:picLocks${attrs}${locks[2]}>` +
      xml.slice(locks.index + locks[0].length)
  } else {
    const lock = `<a:picLocks noChangeAspect="${locked ? 1 : 0}"/>`
    const pr = /<p:cNvPicPr\b([^>]*?)(\/?)>/.exec(xml)
    if (!pr) return false
    if (pr[2] === '/') {
      xml =
        xml.slice(0, pr.index) +
        `<p:cNvPicPr${pr[1]}>${lock}</p:cNvPicPr>` +
        xml.slice(pr.index + pr[0].length)
    } else {
      const at = pr.index + pr[0].length
      xml = xml.slice(0, at) + lock + xml.slice(at)
    }
  }
  commit(slide, el, xml)
  ;(el as { lockAspect?: boolean }).lockAspect = locked
  return true
}
