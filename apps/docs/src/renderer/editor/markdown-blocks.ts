/**
 * Block-level helpers for Markdown content that enters the document from outside the clipboard's
 * own format: default table borders, keeping tables apart, trailing paragraph rules.
 */
import { Fragment, type Node as PmNode, type Schema } from '@tiptap/pm/model'

const DEFAULT_TABLE_LINE = { style: 'single', szEighths: 4, color: 'auto' }
const DEFAULT_TABLE_BORDERS = {
  top: DEFAULT_TABLE_LINE,
  bottom: DEFAULT_TABLE_LINE,
  left: DEFAULT_TABLE_LINE,
  right: DEFAULT_TABLE_LINE,
  insideH: DEFAULT_TABLE_LINE,
  insideV: DEFAULT_TABLE_LINE,
}

/** True when neither the table nor any of its cells declares a border. */
export function tableHasNoBorders(table: PmNode): boolean {
  if (table.attrs.borders) return false
  let declared = false
  table.descendants((node) => {
    if (declared) return false
    if (node.attrs.borders) declared = true
    return !declared
  })
  return !declared
}

/** Tables from other programs carry no border declaration (HTML table borders are CSS we do not
 *  read) and would render invisible. Give them Word's default single borders — what Insert Table
 *  does. Our own clipboard never goes through here, so deliberately borderless tables stay so. */
export function withDefaultTableBorders(content: Fragment): Fragment {
  const out: PmNode[] = []
  content.forEach((node) => {
    if (node.type.name === 'docTable') {
      out.push(
        tableHasNoBorders(node)
          ? node.type.create(
              { ...node.attrs, borders: DEFAULT_TABLE_BORDERS },
              node.content,
              node.marks,
            )
          : node,
      )
    } else if (node.childCount > 0 && !node.isTextblock) {
      out.push(node.copy(withDefaultTableBorders(node.content)))
    } else out.push(node)
  })
  return Fragment.fromArray(out)
}

/** Converted Markdown blocks: default table borders, an empty paragraph between consecutive tables
 *  (Word fuses adjacent tables on save) and optionally one closing a trailing table. */
export function prepareMarkdownBlocks(
  content: Fragment,
  schema: Schema,
  trailingParagraph: boolean,
): Fragment {
  const paragraph = schema.nodes.docParagraph
  const nodes: PmNode[] = []
  withDefaultTableBorders(content).forEach((node) => {
    if (
      node.type.name === 'docTable' &&
      nodes[nodes.length - 1]?.type.name === 'docTable' &&
      paragraph
    ) {
      nodes.push(paragraph.create())
    }
    nodes.push(node)
  })
  if (trailingParagraph && paragraph && nodes[nodes.length - 1]?.type.name === 'docTable') {
    nodes.push(paragraph.create())
  }
  return Fragment.fromArray(nodes)
}

/** A table must never end up directly against another one (Word fuses them on save): add an empty
 *  paragraph where `blocks` would touch the table `before` / `after` the insertion point. */
export function separateTables(
  blocks: Fragment,
  before: PmNode | null,
  after: PmNode | null,
  schema: Schema,
): Fragment {
  const paragraph = schema.nodes.docParagraph?.create()
  const nodes: PmNode[] = []
  blocks.forEach((node) => nodes.push(node))
  if (paragraph && before?.type.name === 'docTable' && nodes[0]?.type.name === 'docTable') {
    nodes.unshift(paragraph)
  }
  if (
    paragraph &&
    after?.type.name === 'docTable' &&
    nodes[nodes.length - 1]?.type.name === 'docTable'
  ) {
    nodes.push(paragraph)
  }
  return Fragment.fromArray(nodes)
}
