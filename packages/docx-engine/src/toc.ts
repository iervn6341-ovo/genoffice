import { DEFAULT_TOC_OPTIONS, type TocOptions } from './generate'
import { decodeEntities } from './parse-xml-text'
import { escapeXmlText } from './xml-utils'

/** Read the instruction of a TOC field without changing any of its source XML. */
export function tocInstructionFromXml(xml: string): string | null {
  const parts: string[] = []
  let inField = false
  let depth = 0
  const tokens =
    /<w:fldChar\b[^>]*\bw:fldCharType="(begin|separate|end)"[^>]*\/?>(?:<\/w:fldChar>)?|<w:instrText\b[^>]*>([\s\S]*?)<\/w:instrText>/g
  let token: RegExpExecArray | null
  while ((token = tokens.exec(xml)) !== null) {
    if (token[1] === 'begin') {
      if (!inField) inField = true
      depth += 1
    } else if (token[1] === 'end') depth -= 1
    else if (token[1] === 'separate' && depth === 1) break
    else if (token[2] !== undefined && inField && depth === 1) parts.push(decodeEntities(token[2]))
  }
  const instruction = parts.join('')
  return /^\s*TOC(?:\s|$)/i.test(instruction) ? instruction : null
}

/** Only options we can represent are editable; unknown Word switches stay protected. */
export function parseTocInstruction(instruction: string): TocOptions | null {
  if (!/^\s*TOC(?:\s|$)/i.test(instruction)) return null
  const options: TocOptions = {
    ...DEFAULT_TOC_OPTIONS,
    levels: 9,
    useHyperlinks: false,
    useOutlineLevels: false,
    styleLevels: {},
  }
  let useHeadingStyles = false
  const tail = instruction.replace(/^\s*TOC\b/i, '')
  const switches = /\\([A-Za-z])(?:\s+(?:"([^"]*)"|([^\s\\]+)))?/g
  let previousEnd = 0
  let match: RegExpExecArray | null
  while ((match = switches.exec(tail)) !== null) {
    if (tail.slice(previousEnd, match.index).trim()) return null
    previousEnd = switches.lastIndex
    const key = match[1]!.toLowerCase()
    const value = match[2] ?? match[3]
    if (key === 'o') {
      const levels = /^1-([1-9])$/.exec(value ?? '')
      if (!levels) return null
      options.levels = Number(levels[1])
      useHeadingStyles = true
    } else if (key === 't') {
      if (!value) return null
      const parts = value.split(',').map((part) => part.trim())
      if (parts.length % 2 || parts.some((part) => !part)) return null
      for (let i = 0; i < parts.length; i += 2) {
        const level = Number(parts[i + 1])
        if (!Number.isInteger(level) || level < 1 || level > 9) return null
        options.styleLevels[parts[i]!] = level
      }
    } else if (key === 'f') {
      if (value) return null // a TC identifier needs its own UI and entry filtering
      options.useTcFields = true
    } else if (key === 'h') options.useHyperlinks = true
    else if (key === 'u') options.useOutlineLevels = true
    else if (key === 'n') {
      if (value) return null // an omitted-level range differs from the whole-table toggle
      options.showPageNumbers = false
    } else if (key === 'z') {
      // This only hides page numbers in Word's Web layout.
    } else return null
  }
  if (tail.slice(previousEnd).trim()) return null
  options.useHeadingStyles =
    useHeadingStyles ||
    (!options.useTcFields &&
      !options.useOutlineLevels &&
      Object.keys(options.styleLevels).length === 0)
  return options
}

/** Backfill settings stored in the TOC result paragraphs rather than the field instruction. */
export function tocOptionsFromField(
  instruction: string,
  firstParagraphXml: string,
): TocOptions | null {
  const options = parseTocInstruction(instruction)
  if (!options) return null
  const pPr = /<w:pPr>[\s\S]*?<\/w:pPr>/.exec(firstParagraphXml)?.[0] ?? ''
  if (options.showPageNumbers) {
    const right = /<w:tab\b(?=[^>]*\bw:val="right")[^>]*\/>/.exec(pPr)?.[0]
    options.rightAlignPageNumbers = !!right
    if (right) {
      const leader = /\bw:leader="(none|dot|hyphen|underscore|heavy|middleDot)"/.exec(right)?.[1]
      options.leader = (leader as TocOptions['leader']) ?? 'none'
    }
  }
  const result = firstParagraphXml.slice(firstParagraphXml.indexOf('w:fldCharType="separate"'))
  const rPr = /<w:rPr>[\s\S]*?<\/w:rPr>/.exec(result)?.[0] ?? ''
  if (/w:color w:val="4472C4"/.test(rPr)) options.format = 'modern'
  else if (/<w:smallCaps\/>/.test(rPr)) options.format = 'formal'
  else if (/<w:i\/>/.test(rPr)) options.format = 'fancy'
  else if (/<w:b\/>/.test(rPr)) options.format = 'distinctive'
  else if (/w:sz w:val="20"/.test(rPr)) options.format = 'simple'
  else if (/w:sz w:val="22"/.test(rPr)) options.format = 'classic'
  return options
}

/** Change only a cached TOC page number, retaining its field code, links, and run formatting. */
export function patchTocCachedPageNumber(
  paragraphXml: string,
  previous: string,
  next: string,
): string | null {
  if (previous === next) return paragraphXml
  const tabAt = Math.max(
    paragraphXml.lastIndexOf('<w:tab/>'),
    paragraphXml.lastIndexOf('<w:tab />'),
  )
  if (tabAt < 0) return null
  const suffix = paragraphXml.slice(tabAt)
  const candidates = [...suffix.matchAll(/<w:t\b([^>]*)>([\s\S]*?)<\/w:t>/g)]
  const last = candidates[candidates.length - 1]
  if (!last || decodeEntities(last[2] ?? '').trim() !== previous.trim()) return null
  const start = tabAt + last.index! + last[0].indexOf('>') + 1
  const end = start + (last[2]?.length ?? 0)
  return paragraphXml.slice(0, start) + escapeXmlText(next) + paragraphXml.slice(end)
}
