import type { Editor } from '@tiptap/core'
import type { Node as PmDocNode } from '@tiptap/pm/model'
import type { SectionInfo } from '@genoffice/docx-engine'
import { markDocSeen } from './tools'

/**
 * Section owning a top-level block. Blocks are mapped by the docxIndex of the
 * nearest original block at or before them; a section-break paragraph inserted
 * this session (`pendingBreak`, its sectPr in a generated paragraph's genXml)
 * ends its section, so blocks after it belong to the following one.
 */
export function sectionIndexAtBlock(
  pmDoc: PmDocNode,
  sections: SectionInfo[],
  topIndex: number,
): number {
  if (sections.length === 0) return 0
  let docxIndex: number | null = null
  for (let i = Math.min(topIndex, pmDoc.childCount - 1); i >= 0; i--) {
    const node = pmDoc.child(i)
    const di = node.attrs?.docxIndex as number | null | undefined
    if (di !== null && di !== undefined) {
      docxIndex = di
      break
    }
    const genXml = node.attrs?.genXml
    if (i < topIndex && typeof genXml === 'string' && genXml.includes('<w:sectPr')) {
      const pending = sectionIndexAtBlock(pmDoc, sections, i)
      if (sections[pending]?.pendingBreak && genXml.includes(sections[pending]!.sectPrXml)) {
        return Math.min(pending + 1, sections.length - 1)
      }
    }
  }
  const s = docxIndex === null ? 0 : sections.findIndex((sec) => docxIndex! <= sec.lastBlockIndex)
  return s >= 0 ? s : sections.length - 1
}

/**
 * Rewrite the sectPr of an unsaved section break in place (the generated
 * paragraph's genXml). Copies of one sectPr can sit in several break
 * paragraphs, so the paragraph is matched by the section it closes, not by
 * its XML alone. False when the paragraph is gone.
 */
export function patchPendingSectPr(
  editor: Editor,
  sections: SectionInfo[],
  index: number,
  nextXml: string,
): boolean {
  const sec = sections[index]
  if (!sec?.pendingBreak) return false
  const pmDoc = editor.state.doc
  let pos = 0
  for (let i = 0; i < pmDoc.childCount; i++) {
    const node = pmDoc.child(i)
    const genXml = node.attrs.genXml
    if (
      node.type.name === 'docProtected' &&
      typeof genXml === 'string' &&
      genXml.includes(sec.sectPrXml) &&
      sectionIndexAtBlock(pmDoc, sections, i) === index
    ) {
      editor.view.dispatch(
        editor.state.tr.setNodeMarkup(pos, undefined, {
          ...node.attrs,
          genXml: genXml.replace(sec.sectPrXml, nextXml),
        }),
      )
      markDocSeen(editor)
      return true
    }
    pos += node.nodeSize
  }
  return false
}

type StartType = SectionInfo['startType']

export interface InsertedBreakTypes {
  /** start type for the document's final section (its body-level sectPr) */
  readonly trailing: StartType | null
  /** start type for a non-final original section, by section index */
  readonly bySection: ReadonlyMap<number, StartType>
  /** rewritten genXml for an inserted break paragraph (top-level index) that
   *  ends a section which began at an earlier inserted break */
  readonly genXml: ReadonlyMap<number, string>
}

/** a section-break paragraph inserted this session: a generated protected block carrying a sectPr */
function isInsertedBreak(node: PmDocNode): boolean {
  return (
    node.type.name === 'docProtected' &&
    (node.attrs.docxIndex === null || node.attrs.docxIndex === undefined) &&
    typeof node.attrs.genXml === 'string' &&
    node.attrs.genXml.includes('<w:sectPr')
  )
}

/**
 * The start types chosen for section breaks inserted this session, derived
 * from the break paragraphs still in the document at save time. Word stores
 * a section's start type in the sectPr that ENDS it, so the type chosen for a
 * break belongs to the next sectPr: the next inserted break in the same
 * section, or the original section's own sectPr. Deriving it here (instead of
 * holding it in app state at insert time) keeps undo/redo of the insertion
 * consistent with what the save writes. Breaks inside a pending (AI-modelled)
 * section are skipped — that path patches its sectPr in place.
 */
export function insertedBreakTypes(
  pmDoc: PmDocNode,
  sections: SectionInfo[],
  applyStartType: (sectPrXml: string, type: StartType) => string,
): InsertedBreakTypes {
  let trailing: StartType | null = null
  const bySection = new Map<number, StartType>()
  const genXml = new Map<number, string>()
  let carried: { owner: number; type: StartType } | null = null
  const flush = () => {
    if (!carried) return
    if (sections.length === 0 || carried.owner >= sections.length - 1) trailing = carried.type
    else bySection.set(carried.owner, carried.type)
    carried = null
  }
  for (let i = 0; i < pmDoc.childCount; i++) {
    const node = pmDoc.child(i)
    const owner = sectionIndexAtBlock(pmDoc, sections, i)
    if (carried && carried.owner !== owner) flush()
    if (!isInsertedBreak(node) || sections[owner]?.pendingBreak) continue
    if (carried) {
      genXml.set(i, applyStartType(String(node.attrs.genXml), carried.type))
      carried = null
    }
    const type = node.attrs.breakStartType as StartType | null | undefined
    if (type) carried = { owner, type }
  }
  flush()
  return { trailing, bySection, genXml }
}
