/**
 * External flavor of the clipboard HTML. The editor's own renderHTML leans on
 * the document stylesheet (cell borders, fills and widths are CSS variables
 * and class rules, list markers are ::before counters), which a pasting
 * program does not have. The fragment is rendered once inside the live editor
 * context and re-emitted with every resolved style inlined, the way a browser
 * copies a selection (html-export.ts does the same for whole documents).
 */
import { serializeChildrenInline } from '../html-export'
import { CLIP_SLICE_ATTR } from './clipboard-transfer'

/** fragments past this many elements are not inlined (the copy must stay instant) */
export const INLINE_ELEMENT_LIMIT = 6000

/** what the live editor draws on a list item (marker text, picture bullet, marker custom properties) */
export interface ListItemDecor {
  attrs: Record<string, string>
  /** `--li-*` custom properties of the item's inline style */
  vars: string[]
}

/** put the editor's attributes on `host` so the document stylesheet applies inside it */
export type HostMounter = (host: HTMLElement) => () => void

/** the first element of an HTML string gets `attrs` (the slice markers ride on it) */
export function withRootAttrs(html: string, attrs: Record<string, string | null>): string {
  const m = /^(\s*(?:<!--[\s\S]*?-->\s*)*<[a-zA-Z][^\s>/]*)/.exec(html)
  if (!m) return html
  const extra = Object.entries(attrs)
    .filter(([, v]) => v !== null)
    .map(([k, v]) => ` ${k}="${String(v).replace(/&/g, '&amp;').replace(/"/g, '&quot;')}"`)
    .join('')
  return m[1] + extra + html.slice(m[1]!.length)
}

/**
 * Clipboard HTML for `dom` (the container ProseMirror serialized the slice
 * into): style-inlined when it can be, ProseMirror's own markup otherwise.
 * The first element's `data-pm-slice` (open depths / context) and the payload
 * attribute are carried over verbatim.
 */
export function enrichClipboardDom(
  dom: HTMLElement,
  mount?: HostMounter,
  listDecor: readonly ListItemDecor[] = [],
): string {
  const plain = dom.innerHTML
  if (!mount) return plain
  const first = dom.firstElementChild
  const carried = {
    'data-pm-slice': first?.getAttribute('data-pm-slice') ?? null,
    [CLIP_SLICE_ATTR]: first?.getAttribute(CLIP_SLICE_ATTR) ?? null,
  }
  if (dom.querySelectorAll('*').length > INLINE_ELEMENT_LIMIT) return plain
  const host = dom.ownerDocument.createElement('div')
  const unmount = mount(host)
  try {
    for (const child of Array.from(dom.childNodes)) host.appendChild(child.cloneNode(true))
    // list markers are decorations of the live document, not part of the serialized slice
    host.querySelectorAll<HTMLElement>('.doc-li').forEach((item, i) => {
      const decor = listDecor[i]
      if (!decor) return
      for (const [k, v] of Object.entries(decor.attrs)) item.setAttribute(k, v)
      for (const decl of decor.vars) {
        const at = decl.indexOf(':')
        if (at > 0) item.style.setProperty(decl.slice(0, at).trim(), decl.slice(at + 1).trim())
      }
    })
    // the payload is the exact model; it must not be inlined through the element walk
    host.querySelector(`[${CLIP_SLICE_ATTR}]`)?.removeAttribute(CLIP_SLICE_ATTR)
    const inlined = serializeChildrenInline(host)
    if (!inlined.trim()) return plain
    return withRootAttrs(inlined, carried)
  } catch {
    return plain
  } finally {
    unmount()
  }
}
