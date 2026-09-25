/**
 * Notes pane editor (PowerPoint Normal view): the notes body as formatted, directly editable
 * text. Each paragraph is a block carrying its source index and each run a span carrying its
 * own, so extractParagraphs returns traced EditParagraphs and the setNotes op keeps every
 * untouched run's bytes. While the caret is here the Home ribbon's Font / alignment commands
 * act on this selection through the same helpers slide text editing uses.
 */
import React, { useEffect, useLayoutEffect, useRef } from 'react'
import type { EditParagraph, NotesParagraphView } from '../../shared/ipc'
import { displayFontFamily } from '../konva-adapter'
import { extractParagraphs, firstFontFamily } from '../TextEditOverlay'

/** The notes pane shows notes at their real point size (100% zoom): 1pt = 96/72 px */
const PX_PER_PT = 96 / 72

function populate(root: HTMLElement, paras: NotesParagraphView[]): void {
  root.replaceChildren()
  paras.forEach((p, pi) => {
    const div = document.createElement('div')
    div.dataset.srcPara = String(pi)
    div.dataset.level = String(p.level)
    if (p.level) div.style.setProperty('--notes-level', String(p.level))
    if (p.align) div.style.textAlign = p.align
    if (p.bullet) div.dataset.glyph = p.bullet
    p.runs.forEach((r, ri) => {
      if (!r.text) return
      const span = document.createElement('span')
      span.dataset.srcRun = String(ri)
      span.textContent = r.text
      const s = span.style
      s.fontSize = `${r.fontSizePt * PX_PER_PT}px`
      if (r.fontFamily) {
        s.fontFamily = displayFontFamily(r.fontFamily)
        // unchanged display font → extractParagraphs commits the model name (or none if inherited)
        span.dataset.displayFont = firstFontFamily(s.fontFamily)
        if (r.fontExplicit) span.dataset.font = r.fontFamily
      }
      if (r.bold) s.fontWeight = 'bold'
      if (r.italic) s.fontStyle = 'italic'
      const deco = [r.underline && 'underline', r.strike && 'line-through'].filter(Boolean)
      if (deco.length) s.textDecoration = deco.join(' ')
      if (r.baseline > 0) s.verticalAlign = 'super'
      else if (r.baseline < 0) s.verticalAlign = 'sub'
      if (r.color) s.color = r.color
      if (r.highlight) s.backgroundColor = r.highlight
      div.appendChild(span)
    })
    if (!div.textContent) div.appendChild(document.createElement('br'))
    root.appendChild(div)
  })
}

/** The pane's current notes: formatted paragraphs plus their plain text */
export function readNotesEditor(root: HTMLElement): { paragraphs: EditParagraph[]; text: string } {
  const paragraphs = extractParagraphs(root, 1)
  const text = paragraphs.map((p) => p.runs.map((r) => r.text).join('')).join('\n')
  return { paragraphs: text ? paragraphs : [], text }
}

export function NotesEditor({
  paragraphs,
  placeholder,
  editorRef,
  defaultFont,
  onEdit,
  onFocus,
  onBlur,
}: {
  /** null while the slide's notes are loading */
  paragraphs: NotesParagraphView[] | null
  placeholder: string
  editorRef: React.RefObject<HTMLDivElement | null>
  /** Theme body font: new text typed into empty notes shows in it (runs name their own) */
  defaultFont: string | null
  /** Any DOM change by typing or a ribbon command */
  onEdit: () => void
  onFocus: () => void
  onBlur: () => void
}): React.JSX.Element {
  const onEditRef = useRef(onEdit)
  onEditRef.current = onEdit
  const observerRef = useRef<MutationObserver | null>(null)

  useLayoutEffect(() => {
    const root = editorRef.current
    if (!root || !paragraphs) return
    populate(root, paragraphs)
    root.classList.toggle('notes-empty', !root.textContent)
    // the rebuild is not an edit
    observerRef.current?.takeRecords()
  }, [paragraphs, editorRef])

  useEffect(() => {
    const root = editorRef.current
    if (!root) return
    // ribbon helpers restyle spans without an input event: watch the DOM itself
    const mo = new MutationObserver(() => {
      root.classList.toggle('notes-empty', !root.textContent)
      onEditRef.current()
    })
    mo.observe(root, {
      subtree: true,
      childList: true,
      characterData: true,
      // formatting lives in inline styles; the root's own placeholder class is not an edit
      attributeFilter: ['style'],
    })
    observerRef.current = mo
    return () => mo.disconnect()
  }, [editorRef])

  return (
    <div
      ref={editorRef}
      className="notes-editor notes-empty"
      contentEditable
      suppressContentEditableWarning
      role="textbox"
      aria-multiline="true"
      aria-label={placeholder}
      data-placeholder={placeholder}
      data-norm="1"
      style={defaultFont ? { fontFamily: displayFontFamily(defaultFont) } : undefined}
      spellCheck
      onFocus={onFocus}
      onBlur={onBlur}
      onPaste={(e) => {
        // plain text, like the notes textarea before: web markup (images, tables, page fonts)
        // has no place in speaker notes; the caret's formatting carries on
        const text = e.clipboardData.getData('text/plain')
        e.preventDefault()
        if (text) document.execCommand('insertText', false, text.replace(/\r\n?/g, '\n'))
      }}
      onKeyDown={(e) => {
        // Esc leaves the notes (PowerPoint returns focus to the slide)
        if (e.key === 'Escape') (e.currentTarget as HTMLElement).blur()
      }}
    />
  )
}

/**
 * Read-only formatted notes (presenter view): sizes are relative to the 12pt notes default, so
 * the presenter's A⁺ / A⁻ scale every run together, as in PowerPoint's presenter view.
 */
export function NotesView({ paragraphs }: { paragraphs: NotesParagraphView[] }): React.JSX.Element {
  return (
    <>
      {paragraphs.map((p, pi) => (
        <div
          key={pi}
          className="notes-view-para"
          data-glyph={p.bullet}
          style={{
            ...(p.align ? { textAlign: p.align } : {}),
            ...(p.level ? { paddingLeft: `${p.level * 1.5}em` } : {}),
          }}
        >
          {p.runs.map((r, ri) =>
            r.text ? (
              <span
                key={ri}
                style={{
                  fontSize: `${r.fontSizePt / 12}em`,
                  ...(r.fontFamily ? { fontFamily: displayFontFamily(r.fontFamily) } : {}),
                  ...(r.bold ? { fontWeight: 'bold' } : {}),
                  ...(r.italic ? { fontStyle: 'italic' } : {}),
                  ...(r.underline || r.strike
                    ? {
                        textDecoration: [r.underline && 'underline', r.strike && 'line-through']
                          .filter(Boolean)
                          .join(' '),
                      }
                    : {}),
                  ...(r.baseline ? { verticalAlign: r.baseline > 0 ? 'super' : 'sub' } : {}),
                  ...(r.color ? { color: r.color } : {}),
                  ...(r.highlight ? { backgroundColor: r.highlight } : {}),
                }}
              >
                {r.text}
              </span>
            ) : null,
          )}
          {!p.runs.some((r) => r.text) && <br />}
        </div>
      ))}
    </>
  )
}
