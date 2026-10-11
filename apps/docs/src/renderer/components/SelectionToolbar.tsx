/**
 * Floating format toolbar over a text selection (Notion's selection menu): block type,
 * text / background color, bold, italic, underline, strikethrough, clear formatting and link.
 *
 * Mounted at the app root with fixed positioning (coordsAtPos speaks viewport coordinates, the
 * PasteOptionsChip pattern). It shows once the selection is settled — not while the mouse is
 * still dragging — and hides on a collapsed selection, a table-cell selection, IME composition
 * or when the document loses focus to something other than the toolbar itself. Every button
 * keeps the editor selection (mousedown is not allowed to move focus).
 */
import React, { useCallback, useEffect, useRef, useState } from 'react'
import type { Editor } from '@tiptap/core'
import { TextSelection } from '@tiptap/pm/state'
import { useI18n, type StringKey } from '../i18n/locale'
import { FORMAT_MARKS } from '../editor/caret-marks'
import type { ListAutoFormatStorage } from '../editor/list-autoformat'
import { IconClearFormat, IconLink } from './icons'

const GAP = 8
const EDGE = 8
const HEIGHT = 36

/** Notion's palette: text colors and their soft background twins (document data, hex without '#') */
const COLORS: Array<{ name: string; text: string | null; fill: string | null }> = [
  { name: 'default', text: null, fill: null },
  { name: 'gray', text: '9B9A97', fill: 'EBECED' },
  { name: 'brown', text: '64473A', fill: 'E9E5E3' },
  { name: 'orange', text: 'D9730D', fill: 'FAEBDD' },
  { name: 'yellow', text: 'DFAB01', fill: 'FBF3DB' },
  { name: 'green', text: '0F7B6C', fill: 'DDEDEA' },
  { name: 'blue', text: '0B6E99', fill: 'DDEBF1' },
  { name: 'purple', text: '6940A5', fill: 'EAE4F2' },
  { name: 'pink', text: 'AD1A72', fill: 'F4DFEB' },
  { name: 'red', text: 'E03E3E', fill: 'FBE4E4' },
]

type BlockKind = 'text' | 'h1' | 'h2' | 'h3' | 'bullet' | 'ordered'

const BLOCKS: Array<[BlockKind, StringKey]> = [
  ['text', 'ribbonStyleNormal'],
  ['h1', 'ribbonStyleHeading1'],
  ['h2', 'ribbonStyleHeading2'],
  ['h3', 'ribbonStyleHeading3'],
  ['bullet', 'ribbonBullets'],
  ['ordered', 'ribbonNumbering'],
]

/** Whether a selection is worth a toolbar: a non-empty run of text inside the document body */
export function wantsToolbar(editor: Editor): boolean {
  const { selection } = editor.state
  if (!editor.isEditable || !(selection instanceof TextSelection) || selection.empty) return false
  // never read a huge selection (select-all in a long document) just to learn it holds text
  const probeEnd = Math.min(selection.to, selection.from + 400)
  if (editor.state.doc.textBetween(selection.from, probeEnd, ' ').trim().length > 0) return true
  return selection.to > probeEnd
}

function blockKindOf(editor: Editor): { kind: BlockKind; level: number } {
  if (editor.isActive('docHeading')) {
    const level = Number(editor.getAttributes('docHeading').level) || 1
    return { kind: level <= 3 ? (`h${level}` as BlockKind) : 'h3', level }
  }
  if (editor.isActive('docListItem')) {
    return {
      kind: editor.getAttributes('docListItem').kind === 'ordered' ? 'ordered' : 'bullet',
      level: 0,
    }
  }
  return { kind: 'text', level: 0 }
}

export function SelectionToolbar({ editor, onLink }: { editor: Editor; onLink: () => void }) {
  const { t } = useI18n()
  const [pos, setPos] = useState<{ left: number; top: number } | null>(null)
  const [menu, setMenu] = useState<'block' | 'color' | null>(null)
  const [, setTick] = useState(0)
  const rootRef = useRef<HTMLDivElement | null>(null)
  const dragging = useRef(false)

  const refresh = useCallback(() => {
    if (dragging.current || !wantsToolbar(editor) || editor.view.composing) {
      setPos(null)
      setMenu(null)
      return
    }
    const { from, to } = editor.state.selection
    try {
      const a = editor.view.coordsAtPos(from)
      const b = editor.view.coordsAtPos(to)
      const sameLine = Math.abs(a.top - b.top) < 4
      const center = sameLine ? (a.left + b.right) / 2 : a.left
      const width = rootRef.current?.offsetWidth ?? 320
      const left = Math.min(Math.max(EDGE, center - width / 2), window.innerWidth - width - EDGE)
      const above = Math.min(a.top, b.top) - HEIGHT - GAP
      const top = above >= EDGE ? above : Math.max(a.bottom, b.bottom) + GAP
      setPos({ left, top })
      setTick((n) => n + 1)
    } catch {
      setPos(null)
    }
  }, [editor])

  useEffect(() => {
    const onDown = (event: MouseEvent) => {
      if (rootRef.current?.contains(event.target as Node)) return
      if (editor.view.dom.contains(event.target as Node)) {
        dragging.current = true
        setPos(null)
      }
      setMenu(null)
    }
    const onUp = () => {
      if (!dragging.current) return
      dragging.current = false
      // let the browser settle the final selection before reading it
      window.setTimeout(refresh, 0)
    }
    const onKeyUp = (event: KeyboardEvent) => {
      if (event.key === 'Escape') {
        setMenu(null)
        return
      }
      if (event.shiftKey || event.key.startsWith('Arrow')) refresh()
    }
    editor.on('selectionUpdate', refresh)
    editor.on('transaction', refresh)
    const onBlur = ({ event }: { event: FocusEvent }) => {
      const next = event?.relatedTarget as Node | null
      if (next && rootRef.current?.contains(next)) return
      setPos(null)
      setMenu(null)
    }
    editor.on('blur', onBlur)
    window.addEventListener('mousedown', onDown, true)
    window.addEventListener('mouseup', onUp, true)
    window.addEventListener('keyup', onKeyUp)
    // scroll/resize fire in bursts: reposition at most once per frame
    let frame = 0
    const reposition = () => {
      if (frame) return
      frame = window.requestAnimationFrame(() => {
        frame = 0
        refresh()
      })
    }
    window.addEventListener('scroll', reposition, true)
    window.addEventListener('resize', reposition)
    return () => {
      editor.off('selectionUpdate', refresh)
      editor.off('transaction', refresh)
      editor.off('blur', onBlur)
      window.removeEventListener('mousedown', onDown, true)
      window.removeEventListener('mouseup', onUp, true)
      window.removeEventListener('keyup', onKeyUp)
      window.removeEventListener('scroll', reposition, true)
      window.removeEventListener('resize', reposition)
      if (frame) window.cancelAnimationFrame(frame)
    }
  }, [editor, refresh])

  if (!pos) return null

  const chain = () => editor.chain().focus()
  const current = blockKindOf(editor)
  const currentLabel =
    current.kind === 'text'
      ? t('ribbonStyleNormal')
      : current.kind === 'bullet'
        ? t('ribbonBullets')
        : current.kind === 'ordered'
          ? t('ribbonNumbering')
          : t(`ribbonStyleHeading${Math.min(current.level, 3)}` as StringKey).replace(
              /\d+$/,
              String(current.level),
            )

  const turnInto = (kind: BlockKind) => {
    setMenu(null)
    if (kind === 'text') {
      chain().setNode('docParagraph', { styleId: null }).run()
    } else if (kind === 'bullet' || kind === 'ordered') {
      const store = editor.storage.listAutoFormat as ListAutoFormatStorage | undefined
      const numId = store?.resolve?.(
        kind === 'bullet'
          ? { kind: 'bullet', glyph: '•' }
          : { kind: 'ordered', numFmt: 'decimal', pattern: '%1.' },
      )
      if (numId) chain().setNode('docListItem', { kind, numId, ilvl: 0 }).run()
    } else {
      chain()
        .setNode('docHeading', { level: Number(kind.slice(1)), styleId: null, outlineOnly: null })
        .run()
    }
  }

  const setStyle = (patch: Record<string, unknown>) => {
    chain().setMark('docTextStyle', patch).run()
    setMenu(null)
  }

  const markBtn = (
    name: 'bold' | 'italic' | 'underline' | 'strike',
    tip: StringKey,
    label: React.ReactNode,
  ) => (
    <button
      type="button"
      className={`sel-tb-btn${editor.isActive(name) ? ' active' : ''}`}
      data-tip={t(tip)}
      aria-label={t(tip)}
      aria-pressed={editor.isActive(name)}
      onClick={() => {
        chain().toggleMark(name).run()
      }}
    >
      {label}
    </button>
  )

  return (
    <div
      ref={rootRef}
      className="sel-tb"
      style={{ left: pos.left, top: pos.top }}
      role="toolbar"
      onMouseDown={(event) => {
        // keep the document selection: only the menu inputs (none) may take focus
        event.preventDefault()
      }}
    >
      <div className="sel-tb-anchor">
        <button
          type="button"
          className="sel-tb-btn sel-tb-block"
          aria-haspopup="menu"
          aria-expanded={menu === 'block'}
          onClick={() => setMenu(menu === 'block' ? null : 'block')}
        >
          {currentLabel}
          <span className="sel-tb-caret">▾</span>
        </button>
        {menu === 'block' && (
          <div className="sel-tb-menu" role="menu">
            {BLOCKS.map(([kind, key]) => (
              <button
                key={kind}
                type="button"
                role="menuitem"
                className={`sel-tb-item${current.kind === kind ? ' active' : ''}`}
                onClick={() => turnInto(kind)}
              >
                <span className="sel-tb-check">{current.kind === kind ? '✓' : ''}</span>
                {t(key)}
              </button>
            ))}
          </div>
        )}
      </div>
      <span className="sel-tb-sep" />
      {markBtn('bold', 'ribbonBoldTip', <b>B</b>)}
      {markBtn('italic', 'ribbonItalicTip', <i>I</i>)}
      {markBtn('underline', 'ribbonUnderlineTip', <u>U</u>)}
      {markBtn('strike', 'ribbonStrikethrough', <s>S</s>)}
      <span className="sel-tb-sep" />
      <div className="sel-tb-anchor">
        <button
          type="button"
          className="sel-tb-btn"
          data-tip={t('ribbonFontColor')}
          aria-label={t('ribbonFontColor')}
          aria-haspopup="menu"
          aria-expanded={menu === 'color'}
          onClick={() => setMenu(menu === 'color' ? null : 'color')}
        >
          <span className="sel-tb-a">A</span>
          <span className="sel-tb-caret">▾</span>
        </button>
        {menu === 'color' && (
          <div className="sel-tb-menu sel-tb-colors" role="menu">
            <div className="sel-tb-title">{t('ribbonFontColor')}</div>
            <div className="sel-tb-swatches">
              {COLORS.map((c) => (
                <button
                  key={`t-${c.name}`}
                  type="button"
                  className="sel-tb-swatch"
                  aria-label={c.name}
                  onClick={() => setStyle({ color: c.text })}
                >
                  <span style={c.text ? { color: `#${c.text}` } : undefined}>A</span>
                </button>
              ))}
            </div>
            <div className="sel-tb-title">{t('ribbonHighlightColors')}</div>
            <div className="sel-tb-swatches">
              {COLORS.map((c) => (
                <button
                  key={`f-${c.name}`}
                  type="button"
                  className="sel-tb-swatch"
                  aria-label={c.name}
                  onClick={() => setStyle({ shading: c.fill, highlight: null })}
                >
                  <span
                    className="sel-tb-fill"
                    style={c.fill ? { background: `#${c.fill}` } : undefined}
                  />
                </button>
              ))}
            </div>
          </div>
        )}
      </div>
      <button
        type="button"
        className="sel-tb-btn"
        data-tip={t('ribbonClearFormatting')}
        aria-label={t('ribbonClearFormatting')}
        onClick={() => {
          let c = chain()
          for (const name of FORMAT_MARKS) c = c.unsetMark(name)
          c.run()
        }}
      >
        <IconClearFormat size={16} />
      </button>
      <button
        type="button"
        className={`sel-tb-btn${editor.isActive('link') ? ' active' : ''}`}
        data-tip={t('ribbonLink')}
        aria-label={t('ribbonLink')}
        onClick={onLink}
      >
        <IconLink size={16} />
      </button>
    </div>
  )
}
