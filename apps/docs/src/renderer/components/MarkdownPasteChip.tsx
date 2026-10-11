/**
 * "Markdown detected" chip: when pasted text looks like Markdown (headings, tables, lists,
 * fences… — from a .md file, an editor, a chat), it is pasted as typed and this small chip appears
 * at the end of it offering to convert it to formatted content. Convert undoes the literal paste
 * and inserts the converted content in its place; the chip goes away on any other edit, Esc, its
 * close button or after a few seconds, so ignoring it costs nothing.
 *
 * Mounted at the app root with fixed positioning (the PasteOptionsChip pattern).
 */
import React, { useEffect, useRef, useState } from 'react'
import type { Editor } from '@tiptap/core'
import type { EditorState, Transaction } from '@tiptap/pm/state'
import type { EditorView } from '@tiptap/pm/view'
import { useI18n } from '../i18n/locale'
import { takeMarkdownPaste } from '../editor/markdown-paste-hint'
import { markdownPasteHtml } from '../editor/markdown-paste'
import { capturePasteRestore, revertPaste, type PasteRestore } from '../editor/paste-options'
import { changedRangeOf } from './PasteOptionsChip'

type ChipState = { text: string; to: number; restore: PasteRestore }

const GAP = 6
const EDGE = 8
const SHOW_MS = 12_000
const REVERT_META = 'markdownPasteRevert'

export function MarkdownPasteChip({
  editor,
  apply,
}: {
  editor: Editor
  /** insert converted Markdown HTML at the selection; false = nothing was inserted */
  apply: (view: EditorView, html: string) => boolean
}) {
  const { t } = useI18n()
  const [chip, setChip] = useState<ChipState | null>(null)
  const [pos, setPos] = useState<{ left: number; top: number } | null>(null)
  const beforeState = useRef<EditorState | null>(null)

  useEffect(() => {
    const onBefore = ({ transaction }: { transaction: Transaction }) => {
      if (transaction.docChanged) beforeState.current = editor.state
    }
    const onTransaction = ({
      transaction: tr,
      appendedTransactions,
    }: {
      transaction: Transaction
      appendedTransactions: Transaction[]
    }) => {
      if (!tr.docChanged) return
      const text = tr.getMeta(REVERT_META) ? null : takeMarkdownPaste()
      if (text === null) {
        setChip(null)
        return
      }
      const range = changedRangeOf(tr)
      const restore =
        range && beforeState.current
          ? capturePasteRestore(beforeState.current, [tr, ...appendedTransactions])
          : null
      setChip(range && restore ? { text, to: range.to, restore } : null)
    }
    editor.on('beforeTransaction', onBefore)
    editor.on('transaction', onTransaction)
    return () => {
      editor.off('beforeTransaction', onBefore)
      editor.off('transaction', onTransaction)
    }
  }, [editor])

  // anchor at the end of the pasted text, follow scrolling, vanish after a while
  useEffect(() => {
    if (!chip) {
      setPos(null)
      return
    }
    const place = () => {
      try {
        const coords = editor.view.coordsAtPos(Math.min(chip.to, editor.state.doc.content.size))
        setPos({
          left: Math.max(EDGE, Math.min(coords.left, window.innerWidth - 300)),
          top: Math.min(coords.bottom + GAP, window.innerHeight - EDGE - 36),
        })
      } catch {
        setChip(null)
      }
    }
    place()
    const timer = window.setTimeout(() => setChip(null), SHOW_MS)
    const onKey = (event: KeyboardEvent) => {
      if (event.key === 'Escape') setChip(null)
    }
    window.addEventListener('scroll', place, true)
    window.addEventListener('resize', place)
    window.addEventListener('keydown', onKey)
    return () => {
      window.clearTimeout(timer)
      window.removeEventListener('scroll', place, true)
      window.removeEventListener('resize', place)
      window.removeEventListener('keydown', onKey)
    }
  }, [editor, chip])

  if (!chip || !pos) return null

  const convert = () => {
    const html = markdownPasteHtml(chip.text)
    const view = editor.view
    const revert = html ? revertPaste(view.state, chip.restore) : null
    setChip(null)
    if (!html || !revert) return
    view.dispatch(revert.setMeta(REVERT_META, true))
    if (!apply(view, html)) {
      // nothing convertible after all: put the literal text back
      editor.commands.undo()
    }
    editor.commands.focus()
  }

  return (
    <div
      className="md-paste-chip"
      style={{ left: pos.left, top: pos.top }}
      role="status"
      data-testid="markdown-paste-chip"
      onMouseDown={(event) => event.preventDefault()}
    >
      <span className="md-paste-chip-label">{t('appMdPasteDetected')}</span>
      <button type="button" className="md-paste-chip-convert" onClick={convert}>
        {t('appMdPasteConvert')}
      </button>
      <button
        type="button"
        className="md-paste-chip-close"
        aria-label={t('appMdPasteDismiss')}
        onClick={() => setChip(null)}
      >
        ×
      </button>
    </div>
  )
}
