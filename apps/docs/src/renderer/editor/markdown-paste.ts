/**
 * Plain-text Markdown paste. Clipboards that carry only text/plain (fenced
 * code blocks, terminals, .md files in a text editor, LLM output) lose all
 * structure when pasted literally: "## Heading" keeps its hashes, "**bold**"
 * its asterisks. When such text looks like Markdown, it is converted to HTML
 * and inserted through the same pipeline as an HTML paste. Clipboards that
 * already carry text/html never reach this path.
 *
 * Detection errs toward literal paste: prose that merely mentions a "#" or a
 * lone "1." must not be reformatted, so conversion requires an unambiguous
 * Markdown construct (heading, fence, link, emphasis, table) or a repeated
 * one (several list/quote lines).
 */
import { marked } from 'marked'

/// Constructs that essentially never appear in prose accidentally.
const STRONG_SIGNALS: readonly RegExp[] = [
  /^#{1,6}\s+\S/m, // ATX heading
  /^\s{0,3}```/m, // fenced code block
  /\[[^\]\n]+\]\([^\s)]+\)/, // [text](url) link
  // **bold** only — the __bold__ form is indistinguishable from Python
  // dunder identifiers (__init__, __name__) and would convert pasted code.
  /(?:^|\W)\*\*[^*\n]+\*\*(?:\W|$)/,
  /^\s{0,3}\|.+\|\s*$\n^\s{0,3}\|[\s:|-]+\|\s*$/m, // table header + separator row
]

/// Constructs that appear in prose too; require several matching lines.
const LIST_LINE = /^\s{0,3}(?:[-*+]|\d{1,3}[.)])\s+\S/
const QUOTE_LINE = /^\s{0,3}>\s?\S/

export function looksLikeMarkdown(text: string): boolean {
  if (STRONG_SIGNALS.some((signal) => signal.test(text))) return true
  const lines = text.split('\n')
  const listLines = lines.filter((line) => LIST_LINE.test(line)).length
  const quoteLines = lines.filter((line) => QUOTE_LINE.test(line)).length
  return listLines >= 2 || quoteLines >= 2
}

/**
 * HTML for a plain-text paste that looks like Markdown, or null to leave the
 * default literal paste in place. The output feeds the same DOM → ProseMirror
 * parse as a native HTML paste, so unknown elements degrade the same way.
 */
export function markdownPasteHtml(text: string): string | null {
  // a multi-megabyte paste is data, not prose: never scan it on the paste path
  if (text.length > MAX_MARKDOWN_PASTE) return null
  const normalized = healPipeTables(text.replace(/\r\n?/g, '\n'))
  if (!looksLikeMarkdown(normalized)) return null
  try {
    return marked.parse(normalized, { gfm: true, breaks: false, async: false })
  } catch {
    return null
  }
}

const MAX_MARKDOWN_PASTE = 400_000

const PIPE_ROW = /^\s{0,3}\|.*\|\s*$/
const PIPE_SEPARATOR = /^\s{0,3}\|[\s:|-]*-[\s:|-]*\|\s*$/

/**
 * A block of pipe rows without the `| --- |` line under the header is not a GFM table, so it
 * would stay literal text. Notes and chat output often drop that line: give such a block (two
 * or more consecutive pipe rows, outside code fences) the separator its first row implies, so
 * it reads as a table with a header row.
 */
/**
 * A table row as typed with a CJK input method: full-width bars (U+FF5C) stand for "|" and a
 * separator line may use full-width / em dashes. Returns the row in plain Markdown, or the text
 * unchanged when it is not a pipe row at all.
 */
export function normalizeTableRow(text: string): string {
  if (!/[|｜]/.test(text)) return text
  const bars = text.replace(/｜/g, '|')
  if (!PIPE_ROW.test(bars)) return text
  return /^[\s|:\-–—－]+$/.test(bars) ? bars.replace(/[–—－]/g, '-') : bars
}

export const isPipeRow = (line: string): boolean => PIPE_ROW.test(line)
export const isPipeSeparator = (line: string): boolean => PIPE_SEPARATOR.test(line)

const countCells = (line: string): number =>
  line
    .trim()
    .replace(/^\||\|$/g, '')
    .split('|').length

/** A delimiter row GFM accepts: as many cells as the header, each `---` with optional colons. */
function validSeparator(line: string, headerCells: number): boolean {
  const cells = line
    .trim()
    .replace(/^\||\|$/g, '')
    .split('|')
  return cells.length === headerCells && cells.every((cell) => /^\s*:?-+:?\s*$/.test(cell))
}

export function healPipeTables(text: string): string {
  if (!/[|｜]/.test(text)) return text
  const lines = text.split('\n').map(normalizeTableRow)
  const out: string[] = []
  let inFence = false
  for (let i = 0; i < lines.length; i += 1) {
    const line = lines[i]!
    if (/^\s{0,3}(```|~~~)/.test(line)) inFence = !inFence
    out.push(line)
    if (inFence || !PIPE_ROW.test(line)) continue
    const previous = lines[i - 1]
    const next = lines[i + 1]
    const startsBlock = previous === undefined || !PIPE_ROW.test(previous)
    if (startsBlock && next !== undefined && PIPE_ROW.test(next)) {
      const cells = Math.max(1, countCells(line))
      const clean = `|${' --- |'.repeat(cells)}`
      if (!PIPE_SEPARATOR.test(next)) out.push(clean)
      else if (!validSeparator(next, cells)) lines[i + 1] = clean
    }
  }
  return out.join('\n')
}

/**
 * HTML clipboards that are only source markup: editors and terminals copy a markdown file as a
 * <div>/<span>/<pre> soup with colors but no document structure. A page with real headings,
 * lists, tables, links or emphasis is rich content and keeps its HTML.
 */
export function isSourceMarkupHtml(html: string): boolean {
  if (/<(?:h[1-6]|ul|ol|li|table|blockquote|a|img|strong|b|em|i|hr)[\s>/]/i.test(html)) return false
  return true
}
