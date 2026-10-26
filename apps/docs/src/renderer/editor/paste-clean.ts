/**
 * Clean-up of pasted HTML before ProseMirror parses it (Word / web fragments).
 * Extracted from App.tsx so the paste pipeline is testable.
 */
import { CLIP_SLICE_ATTR } from './clipboard-transfer'
import { convertMsoLists, inlineStyleSheets, preserveMsoProps } from './foreign-style'
import { isForeignPasteHtml, pushDownWebInlineStyles } from './paste-web-html'

/**
 * Word parity: pasting a copy of a SINGLE spreadsheet cell inserts its
 * content as text, not a 1x1 table (a lone cell copied from Sheets/Excel
 * pasted into Docs as a table and flipped the ribbon into table tools). Only a payload whose entire content is one table with one
 * cell unwraps; anything else — multi-cell tables, tables mixed with prose —
 * stays intact.
 */
export function unwrapSingleCellTable(html: string): string {
  if (!/<table/i.test(html)) return html
  // our own copies rebuild from their exact payload (a one-cell table stays a table)
  if (html.includes(CLIP_SLICE_ATTR)) return html
  try {
    const doc = new window.DOMParser().parseFromString(html, 'text/html')
    const body = doc.body
    const tables = body.querySelectorAll('table')
    if (tables.length !== 1) return html
    const table = tables[0]!
    // The table must be the only real content of the paste. Text equality
    // covers prose ANYWHERE outside the table — including inside wrappers
    // that also contain it — and Element.textContent excludes HTML comments,
    // so Excel's StartFragment markers don't block the unwrap.
    if ((body.textContent ?? '').trim() !== (table.textContent ?? '').trim()) return html
    // text-less content outside the table (images, separators) also keeps it
    if (
      [...body.querySelectorAll('img,svg,video,hr')].some((element) => !table.contains(element))
    ) {
      return html
    }
    const cells = table.querySelectorAll('td,th')
    if (cells.length !== 1) return html
    if (cells[0]!.querySelector('table')) return html
    return cells[0]!.innerHTML
  } catch {
    return html
  }
}

/** Clean pasted Word/web HTML: mso conditional comments, <o:p>, and unwrapping <li><p>x</p></li>;
 *  foreign fragments also get block-level font/size/color pushed down onto
 *  spans so whole-paragraph web copies keep their formatting (r181) */
export function cleanPastedHtml(html: string): string {
  const cleaned = unwrapSingleCellTable(
    html
      .replace(/<!--\[if[\s\S]*?<!\[endif\]-->/g, '')
      .replace(/<o:p>[\s\S]*?<\/o:p>/g, '')
      .replace(/<li([^>]*)>\s*<p[^>]*>([\s\S]*?)<\/p>\s*<\/li>/g, '<li$1>$2</li>'),
  )
  if (!isForeignPasteHtml(html)) return cleaned
  return pushDownWebInlineStyles(foldForeignStyles(cleaned))
}

/** `<style>` rules into inline styles, Word's mso-list paragraphs into list items */
export function foldForeignStyles(html: string): string {
  if (!/<style[\s>]|mso-list:|mso-highlight/i.test(html)) return html
  try {
    const doc = new window.DOMParser().parseFromString(html, 'text/html')
    inlineStyleSheets(doc)
    convertMsoLists(doc)
    preserveMsoProps(doc)
    return doc.body.innerHTML
  } catch {
    return html
  }
}
