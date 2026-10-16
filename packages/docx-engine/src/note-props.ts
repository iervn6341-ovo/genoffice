import type { NoteProps } from './types'
import { escapeXmlAttr } from './xml-utils'

export type NotePropsPatch = {
  [K in keyof NoteProps]?: NoteProps[K] | null
}

type NoteKind = 'footnote' | 'endnote'
type Property = keyof NoteProps

const PROPERTY_TAGS: Record<Property, string> = {
  pos: 'pos',
  numFmt: 'numFmt',
  numStart: 'numStart',
  numRestart: 'numRestart',
}
const ORDER: Property[] = ['pos', 'numFmt', 'numStart', 'numRestart']
const SETTINGS_AFTER_NOTE_PROPS = [
  'compat',
  'docVars',
  'rsids',
  'mathPr',
  'attachedSchema',
  'themeFontLang',
  'clrSchemeMapping',
  'doNotIncludeSubdocsInStats',
  'doNotAutoCompressPictures',
  'forceUpgrade',
  'captions',
  'readModeInkLockDown',
  'smartTagType',
  'schemaLibrary',
  'shapeDefaults',
  'doNotEmbedSmartTags',
  'decimalSymbol',
  'listSeparator',
]

function checkedValue(kind: NoteKind, key: Property, value: string | number): string {
  if (key === 'pos') {
    const allowed = kind === 'footnote' ? ['pageBottom', 'beneathText'] : ['docEnd', 'sectEnd']
    if (!allowed.includes(String(value))) throw new Error(`Invalid ${kind} position`)
  } else if (key === 'numFmt') {
    if (!/^[A-Za-z][A-Za-z0-9]{0,49}$/.test(String(value)))
      throw new Error(`Invalid ${kind} number format`)
  } else if (key === 'numStart') {
    if (!Number.isInteger(value) || Number(value) < 1 || Number(value) > 32767)
      throw new Error(`Invalid ${kind} number start`)
  } else if (!['continuous', 'eachSect', 'eachPage'].includes(String(value))) {
    throw new Error(`Invalid ${kind} number restart`)
  }
  return String(value)
}

function patchInner(inner: string, kind: NoteKind, patch: NotePropsPatch): string {
  let next = inner
  for (const key of ORDER) {
    const value = patch[key]
    if (value === undefined) continue
    const tag = PROPERTY_TAGS[key]
    const re = new RegExp(`<w:${tag}\\b[^>]*\\/>`, 'g')
    const existing = re.exec(next)
    const replacement =
      value === null ? '' : `<w:${tag} w:val="${escapeXmlAttr(checkedValue(kind, key, value))}"/>`
    if (existing) {
      next = next.replace(re, (match, offset: number) =>
        offset === existing.index ? replacement : '',
      )
    } else if (replacement) {
      const later = ORDER.slice(ORDER.indexOf(key) + 1)
        .map((laterKey) => new RegExp(`<w:${PROPERTY_TAGS[laterKey]}\\b`).exec(next)?.index)
        .filter((index): index is number => index !== undefined)
      const separator = new RegExp(`<w:${kind}\\b`).exec(next)?.index
      if (separator !== undefined) later.push(separator)
      const at = later.length ? Math.min(...later) : next.length
      next = next.slice(0, at) + replacement + next.slice(at)
    }
  }
  return next
}

/** Sparse edit of a single note-property element; every untouched child keeps its bytes. */
export function patchNotePropsElement(
  xml: string,
  kind: NoteKind,
  patch: NotePropsPatch,
  create: boolean,
): string {
  if (Object.keys(patch).length === 0) return xml
  const tag = `w:${kind}Pr`
  const pair = new RegExp(`<${tag}\\b([^>]*)>([\\s\\S]*?)</${tag}>`)
  const matched = pair.exec(xml)
  if (matched) {
    const updated = `<${tag}${matched[1]}>${patchInner(matched[2], kind, patch)}</${tag}>`
    return xml.slice(0, matched.index) + updated + xml.slice(matched.index + matched[0].length)
  }
  const self = new RegExp(`<${tag}\\b([^>]*)\\/>`).exec(xml)
  if (self) {
    const updated = `<${tag}${self[1]}>${patchInner('', kind, patch)}</${tag}>`
    return xml.slice(0, self.index) + updated + xml.slice(self.index + self[0].length)
  }
  if (!create || Object.values(patch).every((value) => value === null)) return xml
  const inner = patchInner('', kind, patch)
  const element = `<${tag}>${inner}</${tag}>`
  const parent = xml.includes('<w:settings') ? 'settings' : 'sectPr'
  if (parent === 'settings') {
    const later =
      kind === 'footnote' ? ['endnotePr', ...SETTINGS_AFTER_NOTE_PROPS] : SETTINGS_AFTER_NOTE_PROPS
    const at = later
      .map((name) => new RegExp(`<w:${name}\\b`).exec(xml)?.index)
      .filter((index): index is number => index !== undefined)
      .reduce<number | undefined>(
        (min, index) => (min === undefined ? index : Math.min(min, index)),
        undefined,
      )
    if (at !== undefined) return xml.slice(0, at) + element + xml.slice(at)
    return xml.replace('</w:settings>', `${element}</w:settings>`)
  }
  const at = /<w:(?:footnotePr|endnotePr|type|pgSz|pgMar)\b/.exec(xml)?.index
  if (at !== undefined) return xml.slice(0, at) + element + xml.slice(at)
  return xml.replace('</w:sectPr>', `${element}</w:sectPr>`)
}

/** Whole-document edits also change existing local overrides so they cannot shadow settings.xml. */
export function patchExistingSectionNoteProps(
  blockXml: string,
  edits: Partial<Record<NoteKind, NotePropsPatch>>,
): string {
  const start = blockXml.indexOf('<w:sectPr')
  if (start < 0) return blockXml
  const close = blockXml.lastIndexOf('</w:sectPr>')
  if (close < start) return blockXml
  const change = blockXml.indexOf('<w:sectPrChange', start)
  const liveEnd = change >= 0 && change < close ? change : close
  let live = blockXml.slice(start, liveEnd)
  for (const kind of ['footnote', 'endnote'] as const) {
    const patch = edits[kind]
    if (patch) live = patchNotePropsElement(live, kind, patch, false)
  }
  return blockXml.slice(0, start) + live + blockXml.slice(liveEnd)
}
