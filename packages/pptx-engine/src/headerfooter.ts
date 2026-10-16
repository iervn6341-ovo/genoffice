/** Slide header/footer placeholders. Existing shapes are edited in place so their
 * geometry, style, relationship data, and animation targets survive. */
import type { Paragraph, Slide, TextElement } from './types'
import { randomUUID } from 'node:crypto'
import { posix } from 'node:path'
import { generateParagraphXml } from './generate'
import { cNvPrIdsInXml, pruneTimingForSpids } from './animation'
import { materializeSlide, type OpenedPptx } from './index'
import { nextCNvPrId } from './insert'
import { escapeXmlAttr, escapeXmlText } from './xml-utils'
import { relsPathFor, resolveTarget } from './zip'
import { getFirstSlideNumber, setFirstSlideNumber } from './numbering'

export interface HeaderFooterOptions {
  footer?: string | null
  slideNum?: boolean
  date?: string | null
  dateAuto?: boolean
  /** PowerPoint dynamic date field type. */
  dateFormat?: string
  /** Clear footer-family placeholders on slides using a title layout. */
  dontShowOnTitleSlide?: boolean
  firstSlideNum?: number
}

export interface HeaderFooterState {
  footer: string | null
  slideNum: boolean
  date: string | null
  dateAuto: boolean
  dateFormat: string
  dontShowOnTitleSlide: boolean
  firstSlideNum: number
}

const FIELD_RE = /^datetime[\w.-]*$/

function fieldType(value: string | undefined): string {
  return value && FIELD_RE.test(value) ? value : 'datetime1'
}

function hfSpXml(
  slide: Slide,
  type: 'dt' | 'ftr' | 'sldNum',
  idx: number,
  box: { x: number; y: number; cx: number; cy: number },
  para: Paragraph,
): string {
  const id = nextCNvPrId(slide)
  const nameMap = {
    dt: 'Date Placeholder',
    ftr: 'Footer Placeholder',
    sldNum: 'Slide Number Placeholder',
  }
  return (
    `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${nameMap[type]} ${id}"/>` +
    '<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>' +
    `<p:nvPr><p:ph type="${type}" sz="${type === 'dt' ? 'half' : 'quarter'}" idx="${idx}"/></p:nvPr></p:nvSpPr>` +
    `<p:spPr><a:xfrm><a:off x="${box.x}" y="${box.y}"/><a:ext cx="${box.cx}" cy="${box.cy}"/></a:xfrm>` +
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>' +
    '<p:txBody><a:bodyPr vert="horz" lIns="0" tIns="0" rIns="0" bIns="0" rtlCol="0" anchor="ctr"/><a:lstStyle/>' +
    generateParagraphXml(para) +
    '</p:txBody></p:sp>'
  )
}

/** Replace only placeholder text/field markup; leave the shape and first run's rPr untouched. */
export function editPlaceholderXml(xml: string, value: string, field?: string): string {
  const tx = /<p:txBody>[\s\S]*?<\/p:txBody>/.exec(xml)
  if (!tx) {
    const run = field
      ? `<a:fld id="{${randomUUID()}}" type="${field}"><a:rPr/><a:t>${escapeXmlText(value)}</a:t></a:fld>`
      : `<a:r><a:rPr/><a:t>${escapeXmlText(value)}</a:t></a:r>`
    return xml.replace(
      '</p:sp>',
      () => `<p:txBody><a:bodyPr/><a:lstStyle/><a:p>${run}</a:p></p:txBody></p:sp>`,
    )
  }
  const body = tx[0]
  const run = /<a:(r|fld)\b[^>]*>[\s\S]*?<\/a:\1>/.exec(body)
  const escaped = escapeXmlText(value)
  if (!run) {
    const fresh = field
      ? `<a:fld id="{${randomUUID()}}" type="${field}"><a:rPr/><a:t>${escaped}</a:t></a:fld>`
      : `<a:r><a:rPr/><a:t>${escaped}</a:t></a:r>`
    const nextBody = /<a:p\b[^>]*>/.test(body)
      ? body.replace(/(<a:p\b[^>]*>)/, (_m, open: string) => `${open}${fresh}`)
      : body.replace('</p:txBody>', () => `<a:p>${fresh}</a:p></p:txBody>`)
    return xml.replace(body, () => nextBody)
  }
  let first = run[0]
  const inner = first.replace(/^<a:(?:r|fld)\b[^>]*>/, '').replace(/<\/a:(?:r|fld)>$/, '')
  const textInner = /<a:t\b[^>]*>[\s\S]*?<\/a:t>/.test(inner)
    ? inner.replace(
        /(<a:t\b[^>]*>)[\s\S]*?(<\/a:t>)/,
        (_m, open: string, close: string) => `${open}${escaped}${close}`,
      )
    : `${inner}<a:t>${escaped}</a:t>`
  const originalFieldOpen = /^<a:fld\b[^>]*>/.exec(first)?.[0]
  const fieldOpen = originalFieldOpen
    ? /\btype=["'][^"']*["']/.test(originalFieldOpen)
      ? originalFieldOpen.replace(/\btype=["'][^"']*["']/, `type="${field}"`)
      : originalFieldOpen.replace(/>$/, ` type="${field}">`)
    : `<a:fld id="{${randomUUID()}}" type="${field}">`
  first = field ? `${fieldOpen}${textInner}</a:fld>` : `<a:r>${textInner}</a:r>`
  let seen = false
  const nextBody = body.replace(/<a:(?:r|fld)\b[^>]*>[\s\S]*?<\/a:(?:r|fld)>/g, (slice) => {
    if (!seen) {
      seen = true
      return first
    }
    // Keep later formatting and unsupported run properties, but clear stale text.
    const withoutText = slice.replace(
      /(<a:t\b[^>]*>)[\s\S]*?(<\/a:t>)/g,
      (_m, open: string, close: string) => `${open}${close}`,
    )
    return withoutText.replace(/^<a:fld\b[^>]*>([\s\S]*)<\/a:fld>$/, '<a:r>$1</a:r>')
  })
  return xml.replace(body, () => nextBody)
}

function titleLayout(opened: OpenedPptx, slide: Slide): boolean {
  if (!slide.layoutPath) return false
  const layout = opened.archive.readText(slide.layoutPath)
  return !!layout && /<p:sldLayout\b[^>]*\btype=["']title["']/.test(layout)
}

function inheritedTemplate(
  opened: OpenedPptx,
  slide: Slide,
  type: 'dt' | 'ftr' | 'sldNum',
): { xml: string; path: string } | null {
  for (const path of [slide.layoutPath, slide.masterPath]) {
    if (!path) continue
    const xml = opened.archive.readText(path)
    if (!xml) continue
    for (const match of xml.matchAll(/<p:sp\b[^>]*>[\s\S]*?<\/p:sp>/g)) {
      if (new RegExp(`<p:ph\\b[^>]*\\btype=["']${type}["']`).test(match[0]))
        return { xml: match[0], path }
    }
  }
  return null
}

/** A copied layout/master shape must use relationships owned by the destination slide. */
function cloneInheritedTemplate(
  opened: OpenedPptx,
  slide: Slide,
  type: 'dt' | 'ftr' | 'sldNum',
): string | null {
  const source = inheritedTemplate(opened, slide, type)
  if (!source) return null
  const refs = [...source.xml.matchAll(/\br:(?:id|embed|link)=["']([^"']+)["']/g)]
  const sourceRels = opened.archive.readRels(source.path)
  if (refs.some((ref) => !sourceRels.has(ref[1]!))) {
    throw new Error('Inherited header/footer placeholder has an unresolved relationship')
  }
  const replacements = new Map<string, string>()
  for (const ref of refs) {
    const oldId = ref[1]!
    if (replacements.has(oldId)) continue
    const relation = sourceRels.get(oldId)
    if (!relation) return null
    const target =
      relation.targetMode === 'External'
        ? relation.target
        : posix.relative(posix.dirname(slide.path), resolveTarget(source.path, relation.target))
    if (
      relation.targetMode !== 'External' &&
      !opened.archive.has(resolveTarget(source.path, relation.target))
    ) {
      throw new Error('Inherited header/footer placeholder references a missing package part')
    }
    const existing = [...opened.archive.readRels(slide.path).values()].find(
      (candidate) =>
        candidate.type === relation.type &&
        candidate.target === target &&
        candidate.targetMode === relation.targetMode,
    )
    if (existing) {
      replacements.set(oldId, existing.id)
      continue
    }
    const relsPath = relsPathFor(slide.path)
    const rels =
      opened.archive.readText(relsPath) ??
      '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>'
    const ids = [...rels.matchAll(/\bId=["']rId(\d+)["']/g)].map((match) => Number(match[1]))
    const newId = `rId${Math.max(0, ...ids) + 1}`
    const entry = `<Relationship Id="${newId}" Type="${escapeXmlAttr(relation.type)}" Target="${escapeXmlAttr(target)}"${relation.targetMode ? ` TargetMode="${escapeXmlAttr(relation.targetMode)}"` : ''}/>`
    opened.archive.entries.set(
      relsPath,
      Buffer.from(rels.replace('</Relationships>', () => `${entry}</Relationships>`)),
    )
    replacements.set(oldId, newId)
  }
  return source.xml
    .replace(
      /\br:(id|embed|link)=("|')([^"']+)\2/g,
      (match, kind: string, quote: string, oldId: string) =>
        replacements.has(oldId) ? `r:${kind}=${quote}${replacements.get(oldId)}${quote}` : match,
    )
    .replace(
      /(<p:cNvPr\b[^>]*\bid=)["']\d+["']/,
      (_m, prefix: string) => `${prefix}"${nextCNvPrId(slide)}"`,
    )
}

/** A title layout's p:hf flag also suppresses inherited master footer shapes. */
function setTitleLayoutExclusion(opened: OpenedPptx, excluded: boolean): Set<string> {
  const changed = new Set<string>()
  const layouts = new Set(
    opened.deck.slides.map((slide) => slide.layoutPath).filter((path): path is string => !!path),
  )
  for (const path of layouts) {
    const xml = opened.archive.readText(path)
    if (!xml || !/<p:sldLayout\b[^>]*\btype=["']title["']/.test(xml)) continue
    const existing = /<p:hf\b[^>]*\/>/.exec(xml)?.[0]
    let next = xml
    if (existing) {
      let attrs = existing.replace(/^<p:hf\b|\/>$/g, '')
      attrs = attrs.replace(/\s+(?:dt|ftr|sldNum)=["'][^"']*["']/g, '')
      if (excluded) attrs += ' dt="0" ftr="0" sldNum="0"'
      next = xml.replace(existing, () => (attrs.trim() ? `<p:hf${attrs}/>` : ''))
    } else if (excluded) {
      next = xml.replace(
        /<p:extLst\b|<\/p:sldLayout>/,
        (marker) => `<p:hf dt="0" ftr="0" sldNum="0"/>${marker}`,
      )
    }
    if (next !== xml) {
      opened.archive.entries.set(path, Buffer.from(next, 'utf8'))
      changed.add(path)
    }
  }
  return changed
}

/** Apply to selected slides or all slides. Returns true when any slide changes. */
export function applyHeaderFooter(
  opened: OpenedPptx,
  opts: HeaderFooterOptions,
  slideIndexes?: number[],
): boolean {
  const { deck } = opened
  const targets = slideIndexes ?? deck.slides.map((_, i) => i)
  const { cx, cy } = deck.size
  const barH = Math.round(cy * 0.045)
  const barY = Math.round(cy * 0.93)
  const style = { fontSize: 12, color: '898989' }
  const changedLayouts = slideIndexes
    ? new Set<string>()
    : setTitleLayoutExclusion(opened, !!opts.dontShowOnTitleSlide)
  const numberChanged =
    opts.firstSlideNum != null ? setFirstSlideNumber(opened.archive, opts.firstSlideNum) : false
  let changed = changedLayouts.size > 0 || numberChanged

  for (const i of new Set(targets)) {
    let slide = deck.slides[i]
    if (!slide) continue
    if (
      slide.structureDirty ||
      slide.elements.some(
        (el) =>
          el.dirty ||
          el.dirtyTransform ||
          el.dirtyFill ||
          el.dirtyStroke ||
          el.dirtySrcRect ||
          el.dirtyPPr,
      )
    ) {
      slide = materializeSlide(opened, i) ?? slide
    }
    const excluded = !!opts.dontShowOnTitleSlide && titleLayout(opened, slide)
    const values: Record<'dt' | 'ftr' | 'sldNum', { value: string; field?: string } | null> = {
      dt:
        !excluded && opts.date
          ? { value: opts.date, field: opts.dateAuto ? fieldType(opts.dateFormat) : undefined }
          : null,
      ftr: !excluded && opts.footer ? { value: opts.footer } : null,
      sldNum:
        !excluded && opts.slideNum
          ? { value: String(getFirstSlideNumber(opened.archive) + i), field: 'slidenum' }
          : null,
    }
    let dirty = false
    const removedIds = new Set<number>()
    for (const type of ['dt', 'ftr', 'sldNum'] as const) {
      const existing = slide.elements.filter((el) => el.placeholder === type)
      const wanted = values[type]
      if (!wanted) {
        if (existing.length) {
          // A local empty placeholder masks a layout/master placeholder on this slide.
          const first = existing[0]!
          const blank = editPlaceholderXml(first.anchor.originalXml, '')
          if (blank !== first.anchor.originalXml) {
            first.anchor.originalXml = blank
            dirty = true
          }
          for (const extra of existing.slice(1)) {
            for (const id of cNvPrIdsInXml(extra.anchor.originalXml)) removedIds.add(id)
            slide.elements.splice(slide.elements.indexOf(extra), 1)
            dirty = true
          }
        } else if (slide.decorations?.some((el) => el.placeholder === type)) {
          const source = cloneInheritedTemplate(opened, slide, type)
          if (source) {
            const xml = editPlaceholderXml(source, '')
            slide.elements.push({
              id: `hfmask_${slide.elements.length}_${Date.now().toString(36)}`,
              type: 'text',
              anchor: { spIndex: slide.elements.length, originalXml: xml, range: [0, 0] },
              transform: {
                offset: { x: 0, y: 0, cx: 0, cy: 0 },
                rot: 0,
                flipH: false,
                flipV: false,
              },
            } satisfies TextElement)
            dirty = true
          }
        }
        continue
      }
      if (existing.length) {
        const el = existing[0]!
        const nextXml = editPlaceholderXml(el.anchor.originalXml, wanted.value, wanted.field)
        if (nextXml !== el.anchor.originalXml) {
          el.anchor.originalXml = nextXml
          dirty = true
        }
        for (const extra of existing.slice(1)) {
          for (const id of cNvPrIdsInXml(extra.anchor.originalXml)) removedIds.add(id)
          slide.elements.splice(slide.elements.indexOf(extra), 1)
          dirty = true
        }
        continue
      }
      const box =
        type === 'dt'
          ? { x: Math.round(cx * 0.05), y: barY, cx: Math.round(cx * 0.22), cy: barH }
          : type === 'ftr'
            ? { x: Math.round(cx * 0.3), y: barY, cx: Math.round(cx * 0.4), cy: barH }
            : { x: Math.round(cx * 0.86), y: barY, cx: Math.round(cx * 0.09), cy: barH }
      const para: Paragraph = {
        align: type === 'dt' ? 'left' : type === 'ftr' ? 'center' : 'right',
        runs: [{ text: wanted.value, ...(wanted.field ? { field: wanted.field } : {}), ...style }],
      }
      const source = slide.decorations?.some((el) => el.placeholder === type)
        ? cloneInheritedTemplate(opened, slide, type)
        : null
      const xml = source
        ? editPlaceholderXml(source, wanted.value, wanted.field)
        : hfSpXml(slide, type, type === 'dt' ? 10 : type === 'ftr' ? 11 : 12, box, para)
      const el: TextElement = {
        id: `hfnew_${slide.elements.length}_${Date.now().toString(36)}`,
        type: 'text',
        anchor: { spIndex: slide.elements.length, originalXml: xml, range: [0, 0] },
        transform: { offset: { x: 0, y: 0, cx: 0, cy: 0 }, rot: 0, flipH: false, flipV: false },
      }
      slide.elements.push(el)
      dirty = true
    }
    if (!dirty && (!slide.layoutPath || !changedLayouts.has(slide.layoutPath))) continue
    if (removedIds.size) pruneTimingForSpids(slide, removedIds)
    slide.structureDirty = true
    materializeSlide(opened, i)
    changed = true
  }
  return changed
}

/** Read a slide's current footer state for dialog echo. */
export function readHeaderFooter(slide: Slide): HeaderFooterState {
  let footer: string | null = null
  let date: string | null = null
  let dateAuto = false
  let dateFormat = 'datetime1'
  let slideNum = false
  for (const type of ['ftr', 'dt', 'sldNum'] as const) {
    const el =
      slide.elements.find((candidate) => candidate.placeholder === type) ??
      slide.decorations?.find((candidate) => candidate.placeholder === type)
    if (!el || (el.type !== 'text' && el.type !== 'shape')) continue
    const paras = (el as TextElement).text?.paragraphs ?? []
    const text = paras
      .map((p) => p.runs.map((r) => r.text).join(''))
      .join('\n')
      .trim()
    if (type === 'ftr') footer = text || null
    else if (type === 'dt') {
      date = text || null
      const field = paras.flatMap((p) => p.runs).find((r) => r.field?.startsWith('datetime'))?.field
      dateAuto = !!field
      dateFormat = fieldType(field)
    } else if (type === 'sldNum') {
      slideNum = !!text || paras.some((p) => p.runs.some((r) => r.field === 'slidenum'))
    }
  }
  return {
    footer,
    slideNum,
    date,
    dateAuto,
    dateFormat,
    dontShowOnTitleSlide: false,
    firstSlideNum: 1,
  }
}

/** Infer the title exclusion from stored slide placeholders. When the selected
 * title slide is empty, use the first non-title slide's settings for the dialog. */
export function readDeckHeaderFooter(
  opened: OpenedPptx,
  slideIndex: number,
): HeaderFooterState & { isTitleSlide: boolean; slideNumber: number } {
  const slides = opened.deck.slides
  const titleSlides = slides.filter((slide) => titleLayout(opened, slide))
  const contentSlides = slides.filter((slide) => !titleLayout(opened, slide))
  const active = (state: HeaderFooterState) => !!(state.footer || state.date || state.slideNum)
  const flagOff = (slide: Slide) => {
    const xml = slide.layoutPath ? opened.archive.readText(slide.layoutPath) : null
    const hf = xml && /<p:hf\b[^>]*\/>/.exec(xml)?.[0]
    return (
      !!hf &&
      ['dt', 'ftr', 'sldNum'].every((name) => new RegExp(`\\b${name}=["'](?:0|false)["']`).test(hf))
    )
  }
  const excluded =
    titleSlides.length > 0 &&
    titleSlides.every(flagOff) &&
    titleSlides.every((slide) => !active(readHeaderFooter(slide)))
  const selected = slides[slideIndex]
  const fallback =
    excluded && selected && titleLayout(opened, selected)
      ? contentSlides.find((slide) => active(readHeaderFooter(slide)))
      : undefined
  return {
    ...readHeaderFooter(fallback ?? selected ?? slides[0]!),
    dontShowOnTitleSlide: excluded,
    firstSlideNum: getFirstSlideNumber(opened.archive),
    isTitleSlide: !!selected && titleLayout(opened, selected),
    slideNumber: getFirstSlideNumber(opened.archive) + slideIndex,
  }
}
