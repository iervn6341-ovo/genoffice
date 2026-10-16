import type { Editor } from '@tiptap/core'
import type { Mark as PmMark, Node as PmNode } from '@tiptap/pm/model'
import { patchRunPropertiesXml, type TextEffect } from '@genoffice/docx-engine'
import { charScaleEm, wordKerns } from '../line-metrics'
import { charScaleXAttr } from './text-effects'

export type FontChanges = Partial<{
  fontAscii: string
  font: string
  sizeHalfPoints: number
  color: string
  bold: boolean
  italic: boolean
  underlineStyle: string
  underlineColor: string
  strike: boolean
  dstrike: boolean
  caps: 'all' | 'small' | 'none'
  vanish: boolean
  vertAlign: 'superscript' | 'subscript' | ''
  textEffect: TextEffect | 'none'
  em: 'dot' | 'comma' | 'circle' | 'underDot' | 'none'
  charScalePct: number
  charSpacingTwips: number
  positionHalfPoints: number
  kernHalfPoints: number
  snapToGrid: boolean
  ligatures: 'none' | 'standard' | 'standardContextual' | 'historicalDiscretionary' | 'all'
  numberSpacing: 'default' | 'proportional' | 'tabular'
  numberForm: 'default' | 'lining' | 'oldStyle'
  contextualAlternates: boolean
  stylisticSet: number
}>

export interface FontSample {
  text: string
  attrs: Record<string, unknown>
  marks: ReadonlySet<string>
}

const protectedMarks = new Set(['instrField', 'refField', 'hfField', 'ctrlCheckbox'])

function editableText(node: PmNode): boolean {
  return node.isText && !node.marks.some((mark) => protectedMarks.has(mark.type.name))
}

export function fontSelectionSamples(editor: Editor, from: number, to: number): FontSample[] {
  const samples: FontSample[] = []
  if (from === to) {
    const marks = editor.state.storedMarks ?? editor.state.doc.resolve(from).marks()
    samples.push({
      text: '',
      attrs: marks.find((mark) => mark.type.name === 'docTextStyle')?.attrs ?? {},
      marks: new Set(marks.map((mark) => mark.type.name)),
    })
    return samples
  }
  editor.state.doc.nodesBetween(from, to, (node) => {
    if (!editableText(node)) return
    samples.push({
      text: node.text ?? '',
      attrs: node.marks.find((mark) => mark.type.name === 'docTextStyle')?.attrs ?? {},
      marks: new Set(node.marks.map((mark) => mark.type.name)),
    })
  })
  return samples
}

/** Undefined means no direct value; mixed is separate so an untouched dialog never normalizes it. */
export function commonFontValue<T>(
  samples: FontSample[],
  valueOf: (sample: FontSample) => T,
): { value: T | undefined; mixed: boolean } {
  if (!samples.length) return { value: undefined, mixed: false }
  const value = valueOf(samples[0])
  return { value, mixed: samples.slice(1).some((sample) => valueOf(sample) !== value) }
}

const xmlBool = (name: string, on: boolean) => `<${name} w:val="${on ? '1' : '0'}"/>`
const xmlAttr = (value: string) => value.replace(/&/g, '&amp;').replace(/"/g, '&quot;')

function explicitRFontsXml(raw: string, changes: FontChanges): string | undefined {
  if (changes.fontAscii === undefined && changes.font === undefined) return undefined
  const old = /<w:rFonts\b[^>]*\/?>(?:\s*<\/w:rFonts>)?/.exec(raw)?.[0] ?? ''
  const attrs = new Map<string, string>()
  for (const attr of old.matchAll(/\s+([\w:]+)=(['"])(.*?)\2/g)) attrs.set(attr[1], attr[3])
  if (changes.fontAscii !== undefined) {
    attrs.set('w:ascii', xmlAttr(changes.fontAscii))
    attrs.set('w:hAnsi', xmlAttr(changes.fontAscii))
    attrs.delete('w:asciiTheme')
    attrs.delete('w:hAnsiTheme')
  }
  if (changes.font !== undefined) {
    attrs.set('w:eastAsia', xmlAttr(changes.font))
    attrs.delete('w:eastAsiaTheme')
  }
  return `<w:rFonts${[...attrs].map(([key, value]) => ` ${key}="${value}"`).join('')}/>`
}

const W14_NAMESPACE = 'http://schemas.microsoft.com/office/word/2010/wordml'
const MC_NAMESPACE = 'http://schemas.openxmlformats.org/markup-compatibility/2006'

function aliasScopedTag(tag: string, prefix: string): string {
  if (prefix === 'w14') return tag
  let scoped = tag
  if (!new RegExp(`\\bxmlns:${prefix}=`).test(scoped))
    scoped = scoped.replace(/\s*\/?>(?=$)/, (end) => ` xmlns:${prefix}="${W14_NAMESPACE}"${end}`)
  if (!/\bxmlns:mc=/.test(scoped))
    scoped = scoped.replace(/\s*\/?>(?=$)/, (end) => ` xmlns:mc="${MC_NAMESPACE}"${end}`)
  const ignorable = /\bmc:Ignorable=(['"])(.*?)\1/.exec(scoped)
  if (ignorable) {
    if (!ignorable[2].split(/\s+/).includes(prefix))
      scoped = scoped.replace(
        ignorable[0],
        `mc:Ignorable=${ignorable[1]}${ignorable[2]} ${prefix}${ignorable[1]}`,
      )
  } else scoped = scoped.replace(/\s*\/?>(?=$)/, (end) => ` mc:Ignorable="${prefix}"${end}`)
  return scoped
}

function word2010Prefixes(raw: string, inherited?: readonly string[]): Set<string> {
  const prefixes = new Set(inherited ?? [])
  for (const declaration of raw.matchAll(/\bxmlns:([\w.-]+)\s*=\s*(["'])(.*?)\2/g)) {
    if (declaration[3] === W14_NAMESPACE) prefixes.add(declaration[1])
    else prefixes.delete(declaration[1])
  }
  if (!prefixes.size) prefixes.add('w14')
  return prefixes
}

function word2010PropertyPrefix(
  raw: string,
  localName: string,
  inherited?: readonly string[],
): string {
  const known = word2010Prefixes(raw, inherited)
  for (const tag of raw.matchAll(/<([A-Za-z_][\w.-]*):([A-Za-z_][\w.-]*)\b/g))
    if (tag[2] === localName && known.has(tag[1])) return tag[1]
  return known.values().next().value ?? 'w14'
}

export function word2010PropertyValue(
  raw: string,
  localName: string,
  inherited?: readonly string[],
): string | undefined {
  const prefix = word2010PropertyPrefix(raw, localName, inherited)
  const tag = new RegExp(`<${prefix}:${localName}\\b[^>]*>`).exec(raw)?.[0]
  return tag ? new RegExp(`\\b${prefix}:val="([^"]*)"`).exec(tag)?.[1] : undefined
}

export function word2010PropertyPresent(
  raw: string,
  localName: string,
  inherited?: readonly string[],
): boolean {
  const prefix = word2010PropertyPrefix(raw, localName, inherited)
  return new RegExp(`<${prefix}:${localName}\\b`).test(raw)
}

function stylisticSetWrapper(
  raw: string,
  inherited?: readonly string[],
): { prefix: string; open: string; children: string; selfClosing: boolean } | null {
  const prefixes = word2010Prefixes(raw, inherited)
  const openings = [...raw.matchAll(/<([A-Za-z_][\w.-]*):stylisticSets\b[^>]*>/g)]
  if (openings.length !== 1 || !prefixes.has(openings[0][1])) return null
  const prefix = openings[0][1]
  if (openings[0][0].endsWith('/>'))
    return { prefix, open: openings[0][0], children: '', selfClosing: true }
  const closing = `</${prefix}:stylisticSets>`
  const end = raw.indexOf(closing, openings[0].index! + openings[0][0].length)
  if (end < 0) return null
  return {
    prefix,
    open: openings[0][0],
    children: raw.slice(openings[0].index! + openings[0][0].length, end),
    selfClosing: false,
  }
}

/** Null means an imported set list contains structure that this selector cannot safely model. */
export function enabledStylisticSets(
  raw: string | null | undefined,
  inherited?: readonly string[],
): number[] | null {
  const source = raw ?? ''
  const hasWrapper = /<[A-Za-z_][\w.-]*:stylisticSets\b/.test(source)
  if (!hasWrapper) return []
  const wrapper = stylisticSetWrapper(source, inherited)
  if (!wrapper) return null
  if (wrapper.selfClosing) return []
  const prefixes = word2010Prefixes(source, inherited)
  const child = /<([A-Za-z_][\w.-]*):styleSet\b[^>]*(?:\/>|>\s*<\/\1:styleSet>)/g
  const matches = [...wrapper.children.matchAll(child)]
  if (wrapper.children.replace(child, '').trim()) return null
  const ids: number[] = []
  for (const match of matches) {
    if (!prefixes.has(match[1])) return null
    const id = Number(new RegExp(`\\b${match[1]}:id="(\\d+)"`).exec(match[0])?.[1])
    if (!Number.isInteger(id) || id < 1 || id > 20) return null
    if (new RegExp(`\\b${match[1]}:val="(?:0|false|off)"`).test(match[0])) continue
    ids.push(id)
  }
  return [...new Set(ids)].sort((a, b) => a - b)
}

/** Advanced controls edit their own raw rPr children; all unrelated OOXML stays intact. */
export function fontRawRPrPatch(
  raw: string | null | undefined,
  changes: FontChanges,
  w14Prefixes?: readonly string[],
): string {
  const replacements: Record<string, string | null> = {}
  const rFonts = explicitRFontsXml(raw ?? '', changes)
  if (rFonts) replacements['w:rFonts'] = rFonts
  if (changes.underlineStyle !== undefined || changes.underlineColor !== undefined) {
    const oldStyle = /<w:u\b[^>]*\bw:val="([^"]*)"/.exec(raw ?? '')?.[1]
    const oldColor = /<w:u\b[^>]*\bw:color="([^"]*)"/.exec(raw ?? '')?.[1]
    const styleCandidate = changes.underlineStyle ?? oldStyle ?? 'none'
    const colorCandidate = changes.underlineColor ?? oldColor ?? ''
    const style = /^[A-Za-z0-9]+$/.test(styleCandidate) ? styleCandidate : 'none'
    const color = /^(?:[0-9A-Fa-f]{6}|auto)$/.test(colorCandidate) ? colorCandidate : ''
    replacements['w:u'] = `<w:u w:val="${style}"${color ? ` w:color="${color}"` : ''}/>`
  }
  if (changes.color !== undefined && /^(?:[0-9A-Fa-f]{6}|auto)$/.test(changes.color))
    replacements['w:color'] = `<w:color w:val="${changes.color}"/>`
  if (changes.dstrike !== undefined)
    replacements['w:dstrike'] = xmlBool('w:dstrike', changes.dstrike)
  if (changes.caps !== undefined) {
    replacements['w:caps'] = xmlBool('w:caps', changes.caps === 'all')
    replacements['w:smallCaps'] = xmlBool('w:smallCaps', changes.caps === 'small')
  }
  if (changes.vanish !== undefined) replacements['w:vanish'] = xmlBool('w:vanish', changes.vanish)
  if (changes.textEffect !== undefined) {
    for (const effect of ['outline', 'shadow', 'emboss', 'imprint'] as const)
      replacements[`w:${effect}`] = changes.textEffect === effect ? `<w:${effect}/>` : null
  }
  if (changes.em !== undefined) replacements['w:em'] = `<w:em w:val="${changes.em}"/>`
  if (changes.charScalePct !== undefined)
    replacements['w:w'] = `<w:w w:val="${changes.charScalePct}"/>`
  if (changes.charSpacingTwips !== undefined)
    replacements['w:spacing'] = `<w:spacing w:val="${changes.charSpacingTwips}"/>`
  if (changes.positionHalfPoints !== undefined)
    replacements['w:position'] = `<w:position w:val="${changes.positionHalfPoints}"/>`
  if (changes.kernHalfPoints !== undefined)
    replacements['w:kern'] = `<w:kern w:val="${changes.kernHalfPoints}"/>`
  if (changes.snapToGrid !== undefined)
    replacements['w:snapToGrid'] = xmlBool('w:snapToGrid', changes.snapToGrid)
  if (changes.ligatures !== undefined) {
    const prefix = word2010PropertyPrefix(raw ?? '', 'ligatures', w14Prefixes)
    replacements[`${prefix}:ligatures`] = aliasScopedTag(
      `<${prefix}:ligatures ${prefix}:val="${changes.ligatures}"/>`,
      prefix,
    )
  }
  if (changes.numberSpacing !== undefined) {
    const prefix = word2010PropertyPrefix(raw ?? '', 'numSpacing', w14Prefixes)
    replacements[`${prefix}:numSpacing`] = aliasScopedTag(
      `<${prefix}:numSpacing ${prefix}:val="${changes.numberSpacing}"/>`,
      prefix,
    )
  }
  if (changes.numberForm !== undefined) {
    const prefix = word2010PropertyPrefix(raw ?? '', 'numForm', w14Prefixes)
    replacements[`${prefix}:numForm`] = aliasScopedTag(
      `<${prefix}:numForm ${prefix}:val="${changes.numberForm}"/>`,
      prefix,
    )
  }
  if (changes.contextualAlternates !== undefined) {
    const prefix = word2010PropertyPrefix(raw ?? '', 'cntxtAlts', w14Prefixes)
    replacements[`${prefix}:cntxtAlts`] = aliasScopedTag(
      changes.contextualAlternates
        ? `<${prefix}:cntxtAlts/>`
        : `<${prefix}:cntxtAlts ${prefix}:val="0"/>`,
      prefix,
    )
  }
  if (changes.stylisticSet !== undefined) {
    const id = changes.stylisticSet
    const source = raw ?? ''
    const wrapper = stylisticSetWrapper(source, w14Prefixes)
    const unknown = /<[A-Za-z_][\w.-]*:stylisticSets\b/.test(source) && !wrapper
    if (!unknown) {
      const prefix = wrapper?.prefix ?? w14Prefixes?.[0] ?? 'w14'
      if (id === 0) replacements[`${prefix}:stylisticSets`] = null
      else if (Number.isInteger(id) && id >= 1 && id <= 20) {
        const open = aliasScopedTag(
          wrapper?.open.replace(/\s*\/>$/, '>') ?? `<${prefix}:stylisticSets>`,
          prefix,
        )
        replacements[`${prefix}:stylisticSets`] =
          `${open}<${prefix}:styleSet ${prefix}:id="${id}"/></${prefix}:stylisticSets>`
      }
    }
  }
  return Object.keys(replacements).length
    ? patchRunPropertiesXml(raw, replacements, [...word2010Prefixes(raw ?? '', w14Prefixes)])
    : (raw ?? '')
}

function patchedStyleAttrs(
  attrs: Record<string, unknown>,
  text: string,
  changes: FontChanges,
  marks: readonly PmMark[],
): Record<string, unknown> {
  const next = { ...attrs }
  for (const key of [
    'fontAscii',
    'font',
    'sizeHalfPoints',
    'color',
    'underlineStyle',
    'underlineColor',
    'dstrike',
    'caps',
    'vanish',
    'textEffect',
    'em',
    'charSpacingTwips',
    'charScalePct',
    'positionHalfPoints',
    'kernHalfPoints',
  ] as const) {
    const value = changes[key]
    if (value !== undefined)
      next[key] = value === '' || (value === 'none' && key === 'textEffect') ? null : value
  }
  if (changes.bold !== undefined) next.boldOff = changes.bold ? null : true
  if (changes.color !== undefined) next.themeColor = null
  if (changes.fontAscii !== undefined || changes.font !== undefined) {
    const themed = attrs.themeRFonts
      ? (JSON.parse(String(attrs.themeRFonts)) as {
          font?: string
          fontAscii?: string
        })
      : {}
    if (changes.fontAscii !== undefined) {
      delete themed.fontAscii
      const rawFonts = /<w:rFonts\b[^>]*>/.exec(String(attrs.rawRPr ?? ''))?.[0] ?? ''
      if (!/\bw:eastAsia(?:Theme)?=/.test(rawFonts) && changes.font === undefined) {
        next.font = changes.fontAscii
        delete themed.font
      }
    }
    if (changes.font !== undefined) delete themed.font
    next.themeRFonts = Object.keys(themed).length ? JSON.stringify(themed) : null
  }
  if (changes.italic !== undefined) next.italicOff = changes.italic ? null : true
  if (changes.vertAlign !== undefined) next.vertAlign = changes.vertAlign || null
  if (changes.em === 'none') next.em = null
  if (changes.textEffect === 'none') next.textEffect = null
  if (
    changes.charScalePct !== undefined ||
    changes.underlineStyle !== undefined ||
    changes.strike !== undefined
  ) {
    const pct = Number(next.charScalePct ?? 100)
    const underline =
      changes.underlineStyle !== undefined
        ? changes.underlineStyle !== 'none'
        : marks.some((mark) => mark.type.name === 'underline')
    const strike = changes.strike ?? marks.some((mark) => mark.type.name === 'strike')
    const link = marks.some((mark) => mark.type.name === 'link')
    const scaleX = pct !== 100 && text !== '' && !/\s/.test(text) && !underline && !strike && !link
    next.charScaleX = scaleX ? charScaleXAttr(text, pct, charScaleEm(text, pct)) : null
    next.charScaleEm = pct !== 100 && !scaleX ? charScaleEm(text, pct) : null
  }
  if (changes.kernHalfPoints !== undefined || changes.sizeHalfPoints !== undefined) {
    const threshold = next.kernHalfPoints == null ? NaN : Number(next.kernHalfPoints)
    next.kern = Number.isFinite(threshold)
      ? (wordKerns(threshold, Number(next.sizeHalfPoints) || undefined) ?? null)
      : null
  }
  const raw = fontRawRPrPatch(
    attrs.rawRPr as string | null | undefined,
    changes,
    (attrs.rawRPrW14Prefixes as string[] | null | undefined) ?? undefined,
  )
  next.rawRPr = raw || null
  if (
    !next.rawRPrW14Prefixes &&
    (changes.ligatures !== undefined ||
      changes.numberSpacing !== undefined ||
      changes.numberForm !== undefined ||
      changes.contextualAlternates !== undefined ||
      changes.stylisticSet !== undefined)
  )
    next.rawRPrW14Prefixes = ['w14']
  return next
}

/** Patch each selected text run independently, leaving protected fields and atoms untouched. */
export function applyFontChanges(
  editor: Editor,
  from: number,
  to: number,
  changes: FontChanges,
): boolean {
  if (Object.keys(changes).length === 0 || !editor.isEditable) return false
  const { state, view } = editor
  const tr = state.tr
  const styleType = state.schema.marks.docTextStyle
  const flagTypes = ['bold', 'italic', 'strike'] as const
  const flagValues = {
    bold: changes.bold,
    italic: changes.italic,
    strike: changes.strike,
  }
  const patchUnderline = (start: number, end: number, marks: readonly PmMark[]) => {
    if (changes.underlineStyle === undefined && changes.underlineColor === undefined) return
    const type = state.schema.marks.underline
    const old = marks.find((mark) => mark.type === type)
    const style = changes.underlineStyle ?? String(old?.attrs.style ?? 'single')
    if (style === 'none') {
      tr.removeMark(start, end, type)
    } else if (old || changes.underlineStyle !== undefined) {
      tr.addMark(
        start,
        end,
        type.create({
          style,
          color: changes.underlineColor ?? old?.attrs.color ?? null,
        }),
      )
    }
  }
  const patchRange = (start: number, end: number, node: PmNode, pos: number) => {
    const old = node.marks.find((mark) => mark.type === styleType)
    const attrs = patchedStyleAttrs(
      old?.attrs ?? {},
      node.text?.slice(start - pos, end - pos) ?? '',
      changes,
      node.marks,
    )
    tr.addMark(start, end, styleType.create(attrs))
    for (const flag of flagTypes) {
      const value = flagValues[flag]
      if (value === undefined) continue
      const type = state.schema.marks[flag]
      if (value) tr.addMark(start, end, type.create())
      else tr.removeMark(start, end, type)
    }
    patchUnderline(start, end, node.marks)
  }
  if (from === to) {
    const marks = state.storedMarks ?? state.doc.resolve(from).marks()
    if (marks.some((mark) => protectedMarks.has(mark.type.name))) return false
    const old = marks.find((mark) => mark.type === styleType)
    const next = marks.filter((mark) => mark.type !== styleType) as PmMark[]
    next.push(styleType.create(patchedStyleAttrs(old?.attrs ?? {}, '', changes, marks)))
    for (const flag of flagTypes) {
      const value = flagValues[flag]
      if (value === undefined) continue
      const type = state.schema.marks[flag]
      const index = next.findIndex((mark) => mark.type === type)
      if (index >= 0) next.splice(index, 1)
      if (value) next.push(type.create())
    }
    if (changes.underlineStyle !== undefined || changes.underlineColor !== undefined) {
      const type = state.schema.marks.underline
      const index = next.findIndex((mark) => mark.type === type)
      const oldUnderline = index >= 0 ? next.splice(index, 1)[0] : undefined
      const style = changes.underlineStyle ?? String(oldUnderline?.attrs.style ?? 'single')
      if (style !== 'none' && (oldUnderline || changes.underlineStyle !== undefined))
        next.push(
          type.create({
            style,
            color: changes.underlineColor ?? oldUnderline?.attrs.color ?? null,
          }),
        )
    }
    tr.setStoredMarks(next)
  } else {
    state.doc.nodesBetween(from, to, (node, pos) => {
      if (!editableText(node)) return
      const start = Math.max(from, pos)
      const end = Math.min(to, pos + node.nodeSize)
      if (start < end) patchRange(start, end, node, pos)
    })
  }
  if (!tr.docChanged && !tr.storedMarksSet) return false
  view.dispatch(tr.scrollIntoView())
  view.focus()
  return true
}
