import { useEffect, useRef, useState } from 'react'
import type { Editor } from '@tiptap/core'
import type { Command } from '@tiptap/pm/state'
import { NodeSelection, TextSelection } from '@tiptap/pm/state'
import {
  CellSelection,
  addColumnAfter,
  addColumnBefore,
  addRowAfter,
  addRowBefore,
  deleteColumn,
  deleteRow,
  deleteTable,
  isInTable,
  mergeCells,
  selectedRect,
  splitCell,
} from '@tiptap/pm/tables'
import { platformShortcuts } from '@genoffice/i18n'
import { Dropdown } from '@genoffice/ui'
import type { TabStop } from '@genoffice/docx-engine'
import { getLang, useI18n, type StringKey } from '../i18n/locale'
import { activeParaAttrs } from './ribbon-tabs'
import { wordRangeAtCaret } from '../editor/comments'
import {
  applyParagraphChanges,
  commonParagraphValue,
  selectedParagraphs,
} from '../editor/paragraph-advanced'
import { IconSparkle } from './icons'
import { useModalKeys } from './modal-keys'

/**
 * Editor context menu (right click in the document body):
 * Cut/Copy/Paste · Font… · Paragraph… · Synonyms · Translate · Hyperlink… · New Comment.
 */

export interface ContextMenuState {
  x: number
  y: number
  /** src of the picture under the pointer, when the click landed on one */
  imageSrc?: string | null
  noteRef?: { kind: 'footnote' | 'endnote'; id: string }
}

interface EditorContextMenuProps {
  editor: Editor
  menu: ContextMenuState
  onClose: () => void
  onFontDialog: () => void
  onParagraphDialog: () => void
  onLink: () => void
  onNewComment: () => void
  onViewImage: (src: string) => void
  onSaveImageAs: (src: string) => void
  onAiPreset: (instruction: string) => void
  /** List items: restart numbering / continue numbering (shown when the cursor is on a docListItem) */
  onRestartNumbering?: () => void
  onContinueNumbering?: () => void
  /** F9 update fields (shown when the cursor is on an inline field) */
  onUpdateFields?: () => void
  onConvertNote?: (kind: 'footnote' | 'endnote', id: string) => void
}

/** target languages mirrored from the Review → Translate dropdown; the localized label also goes into the LLM prompt */
const TRANSLATE_TARGETS: Array<{ labelKey: StringKey }> = [
  { labelKey: 'appLangEnglish' },
  { labelKey: 'appLangSimplifiedChinese' },
  { labelKey: 'appLangJapanese' },
  { labelKey: 'appLangKorean' },
  { labelKey: 'appLangFrench' },
  { labelKey: 'appLangGerman' },
  { labelKey: 'appLangSpanish' },
]

const MENU_WIDTH = 240

export function EditorContextMenu({
  editor,
  menu,
  onClose,
  onFontDialog,
  onParagraphDialog,
  onLink,
  onNewComment,
  onViewImage,
  onSaveImageAs,
  onAiPreset,
  onRestartNumbering,
  onContinueNumbering,
  onUpdateFields,
  onConvertNote,
}: EditorContextMenuProps) {
  const { t } = useI18n()
  const ref = useRef<HTMLDivElement>(null)
  const [submenu, setSubmenu] = useState<string | null>(null)

  const { from, to } = editor.state.selection
  const hasSelection = from !== to
  const canComment = hasSelection || wordRangeAtCaret(editor) !== null
  const canEdit = editor.isEditable
  const selectedText = hasSelection ? editor.state.doc.textBetween(from, to, ' ').trim() : ''
  // Synonyms targets a word / short phrase, not long selections
  const synonymText = selectedText.length > 0 && selectedText.length <= 20 ? selectedText : ''

  // keep the menu inside the viewport (flip up / clamp left near the edges)
  const [pos, setPos] = useState({ left: menu.x, top: menu.y })
  useEffect(() => {
    const el = ref.current
    if (!el) return
    const rect = el.getBoundingClientRect()
    let left = menu.x
    let top = menu.y
    if (left + rect.width > window.innerWidth - 8) left = window.innerWidth - rect.width - 8
    if (top + rect.height > window.innerHeight - 8) top = Math.max(8, menu.y - rect.height)
    setPos({ left, top })
  }, [menu])

  useEffect(() => {
    const close = (e: MouseEvent) => {
      if (!ref.current?.contains(e.target as Node)) onClose()
    }
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') onClose()
    }
    window.addEventListener('mousedown', close)
    window.addEventListener('keydown', onKey)
    window.addEventListener('blur', onClose)
    // shell tab-strip presses never reach this document; the preload relays
    // them (app:chrome-pressed) so the menu still dismisses
    const offChrome = window.desktop?.onChromePressed?.(onClose)
    return () => {
      window.removeEventListener('mousedown', close)
      window.removeEventListener('keydown', onKey)
      window.removeEventListener('blur', onClose)
      offChrome?.()
    }
  }, [onClose])

  const run = (action: () => void) => () => {
    onClose()
    action()
  }

  // ---- table section (shown when the cursor is inside a table, Word parity) ----
  // the move handle selects the whole table as a NodeSelection, which isInTable
  // does not treat as "inside" — the menu must still offer the table commands then
  const tableSelected =
    editor.state.selection instanceof NodeSelection &&
    editor.state.selection.node.type.name === 'docTable'
  const inTable = isInTable(editor.state) || tableSelected
  /** cell commands can't run on a whole-table NodeSelection: drop the caret into the first cell */
  const enterFirstCell = () => {
    const sel = editor.state.selection
    if (sel instanceof NodeSelection && sel.node.type.name === 'docTable') {
      editor.view.dispatch(
        editor.state.tr.setSelection(TextSelection.near(editor.state.doc.resolve(sel.from + 1))),
      )
    }
  }
  const runTable = (command: Command) => {
    editor.view.focus()
    enterFirstCell()
    command(editor.state, editor.view.dispatch)
  }
  /** anchor cell of the current selection (top-left of a multi-cell selection) */
  const anchorCellPos = (): number | null => {
    if (!inTable) return null
    enterFirstCell()
    const rect = selectedRect(editor.state)
    return rect.tableStart + rect.map.map[rect.top * rect.map.width + rect.left]
  }
  const selectRowOrColumn = (kind: 'row' | 'column') => {
    const pos = anchorCellPos()
    if (pos === null) return
    const $cell = editor.state.doc.resolve(pos)
    const selection =
      kind === 'row' ? CellSelection.rowSelection($cell) : CellSelection.colSelection($cell)
    editor.view.focus()
    editor.view.dispatch(editor.state.tr.setSelection(selection))
  }
  const selectWholeTable = () => {
    const { $from } = editor.state.selection
    for (let depth = $from.depth; depth >= 1; depth--) {
      if ($from.node(depth).type.name === 'docTable') {
        editor.view.focus()
        editor.view.dispatch(
          editor.state.tr.setSelection(NodeSelection.create(editor.state.doc, $from.before(depth))),
        )
        return
      }
    }
  }

  const protAttrs = editor.getAttributes('docProtected')
  const isImage = protAttrs?.blockType === 'image'
  const isFloating =
    isImage ||
    (Array.isArray(protAttrs?.textboxes) && (protAttrs.textboxes as unknown[]).length > 0)
  const currentWrap = (protAttrs?.imageWrap as string | null) ?? null
  const setWrap = (wrap: string | null) => {
    const clearedPosition =
      wrap === null
        ? { imagePosH: null, imagePosV: null, imageOffsetXEmu: null, imageOffsetYEmu: null }
        : {}
    editor
      .chain()
      .focus()
      .updateAttributes('docProtected', { imageWrap: wrap, ...clearedPosition })
      .run()
  }
  // Stacking order among overlapping floating pictures. z-order only has a
  // visible effect on floating (front/behind) images, so the menu enables it
  // there; a bring-forward on an inline image also floats it (Word parity).
  const currentZOrder = Number((protAttrs?.imageZOrder as number | null) ?? 0)
  const isFloatingWrap = currentWrap === 'front' || currentWrap === 'behind'
  const setZOrder = (z: number) => {
    const attrs: Record<string, unknown> = { imageZOrder: z }
    // an inline image has no paint order; floating it (in front) makes the
    // reorder meaningful, matching Word's "Bring to Front" on an inline picture
    if (!isFloatingWrap) attrs.imageWrap = 'front'
    editor.chain().focus().updateAttributes('docProtected', attrs).run()
  }
  /** z-order of every floating anchor in the document (Word's to-front/to-back are document-global) */
  const floatingZOrders = (): number[] => {
    const zs: number[] = [currentZOrder]
    editor.state.doc.descendants((n) => {
      if (
        n.type.name === 'docProtected' &&
        (n.attrs.imageWrap === 'front' || n.attrs.imageWrap === 'behind')
      )
        zs.push(Number(n.attrs.imageZOrder ?? 0))
    })
    return zs
  }
  const bringToFront = () => setZOrder(Math.max(...floatingZOrders()) + 1)
  const sendToBack = () => setZOrder(Math.min(...floatingZOrders()) - 1)
  const bringForward = () => setZOrder(currentZOrder + 1)
  const sendBackward = () => setZOrder(currentZOrder - 1)

  /** Plain-text insertion: no HTML parsing (insertContent(string) would treat < > as tags) */
  const insertPlainText = (text: string) => {
    const lines = text.replace(/\r/g, '').split('\n')
    if (lines.length === 1) {
      editor.chain().focus().insertContent({ type: 'text', text: lines[0] }).run()
      return
    }
    editor
      .chain()
      .focus()
      .insertContent(
        lines.map((line) => ({
          type: 'docParagraph',
          ...(line ? { content: [{ type: 'text', text: line }] } : {}),
        })),
      )
      .run()
  }

  const clipboard = async (action: 'cut' | 'copy' | 'paste' | 'pastePlain') => {
    if (action === 'paste') {
      // rich paste: prefer HTML (parsed by editor paste rules, keeps formatting), otherwise plain text
      try {
        const items = await navigator.clipboard.read()
        for (const item of items) {
          if (item.types.includes('text/html')) {
            const html = await (await item.getType('text/html')).text()
            editor
              .chain()
              .focus()
              .insertContent(html, { parseOptions: { preserveWhitespace: true } })
              .run()
            return
          }
        }
      } catch {
        /* clipboard.read unavailable (permission/format): fall back to plain text */
      }
      const text = await navigator.clipboard.readText()
      if (text) insertPlainText(text)
    } else if (action === 'pastePlain') {
      const text = await navigator.clipboard.readText()
      if (text) insertPlainText(text)
    } else {
      editor.commands.focus()
      document.execCommand(action)
    }
  }

  const item = (
    label: string,
    opts: {
      key?: string
      disabled?: boolean
      onClick?: () => void
      submenuKey?: string
      ai?: boolean
    },
  ) => (
    <button
      className="ctx-item"
      disabled={opts.disabled}
      data-tip={opts.ai ? t('appAiBadgeTip') : undefined}
      onMouseEnter={() => setSubmenu(opts.submenuKey ?? null)}
      onClick={opts.submenuKey ? undefined : opts.onClick}
    >
      <span className="ctx-label">{label}</span>
      {opts.ai && (
        <span className="copilot-badge copilot-badge-menu">
          <IconSparkle size={10} />
        </span>
      )}
      {opts.key && <span className="ctx-key">{platformShortcuts(opts.key)}</span>}
      {opts.submenuKey && <span className="ctx-arrow">›</span>}
    </button>
  )

  return (
    <div
      ref={ref}
      className="ctx-menu"
      style={{ left: pos.left, top: pos.top, minWidth: MENU_WIDTH }}
      onContextMenu={(e) => e.preventDefault()}
    >
      {menu.imageSrc && (
        <>
          {item(t('appViewImage'), { onClick: run(() => onViewImage(menu.imageSrc!)) })}
          {item(t('appSaveImageAs'), { onClick: run(() => onSaveImageAs(menu.imageSrc!)) })}
          <div className="ctx-sep" />
        </>
      )}
      {item(t('appCut'), {
        key: '⌘X',
        disabled: !hasSelection || !canEdit,
        onClick: run(() => void clipboard('cut')),
      })}
      {item(t('appCopy'), {
        key: '⌘C',
        disabled: !hasSelection,
        onClick: run(() => void clipboard('copy')),
      })}
      {item(t('appPaste'), {
        key: '⌘V',
        disabled: !canEdit,
        onClick: run(() => void clipboard('paste')),
      })}
      {item(t('appPastePlain'), {
        disabled: !canEdit,
        onClick: run(() => void clipboard('pastePlain')),
      })}
      <div className="ctx-sep" />
      {item(t('appFontMenu'), { key: '⌘D', onClick: run(onFontDialog) })}
      {item(t('appParagraphMenu'), { key: '⌥⌘M', onClick: run(onParagraphDialog) })}
      {inTable && (
        <>
          <div className="ctx-sep" />
          <div className="ctx-item-wrap" onMouseLeave={() => setSubmenu(null)}>
            {item(t('ribbonInsert'), { submenuKey: 'tableInsert', disabled: !canEdit })}
            {submenu === 'tableInsert' && canEdit && (
              <div className="ctx-submenu">
                {(
                  [
                    ['ribbonInsertAbove', addRowBefore],
                    ['ribbonInsertBelow', addRowAfter],
                    ['ribbonInsertLeft', addColumnBefore],
                    ['ribbonInsertRight', addColumnAfter],
                  ] as Array<[StringKey, Command]>
                ).map(([labelKey, command]) => (
                  <button
                    key={labelKey}
                    className="ctx-item"
                    onClick={run(() => runTable(command))}
                  >
                    <span className="ctx-label">{t(labelKey)}</span>
                  </button>
                ))}
              </div>
            )}
          </div>
          <div className="ctx-item-wrap" onMouseLeave={() => setSubmenu(null)}>
            {item(t('ribbonTableDeleteMenu'), { submenuKey: 'tableDelete', disabled: !canEdit })}
            {submenu === 'tableDelete' && canEdit && (
              <div className="ctx-submenu">
                {(
                  [
                    ['ribbonDeleteRow', deleteRow],
                    ['ribbonDeleteColumn', deleteColumn],
                    ['ribbonDeleteTable', deleteTable],
                  ] as Array<[StringKey, Command]>
                ).map(([labelKey, command]) => (
                  <button
                    key={labelKey}
                    className="ctx-item"
                    onClick={run(() => runTable(command))}
                  >
                    <span className="ctx-label">{t(labelKey)}</span>
                  </button>
                ))}
              </div>
            )}
          </div>
          <div className="ctx-item-wrap" onMouseLeave={() => setSubmenu(null)}>
            {item(t('ribbonSelect'), { submenuKey: 'tableSelect' })}
            {submenu === 'tableSelect' && (
              <div className="ctx-submenu">
                <button className="ctx-item" onClick={run(() => selectRowOrColumn('row'))}>
                  <span className="ctx-label">{t('ribbonSelectRow')}</span>
                </button>
                <button className="ctx-item" onClick={run(() => selectRowOrColumn('column'))}>
                  <span className="ctx-label">{t('ribbonSelectColumn')}</span>
                </button>
                <button className="ctx-item" onClick={run(selectWholeTable)}>
                  <span className="ctx-label">{t('ribbonSelectTable')}</span>
                </button>
              </div>
            )}
          </div>
          {item(t('ribbonMergeCells'), {
            disabled: !canEdit || !mergeCells(editor.state),
            onClick: run(() => runTable(mergeCells)),
          })}
          {item(t('ribbonSplitCells'), {
            disabled: !canEdit || !splitCell(editor.state),
            onClick: run(() => runTable(splitCell)),
          })}
        </>
      )}
      {editor.isActive('instrField') && onUpdateFields && (
        <>
          <div className="ctx-sep" />
          {item(t('appUpdateField'), { key: 'F9', onClick: run(() => onUpdateFields()) })}
        </>
      )}
      {menu.noteRef && onConvertNote && (
        <>
          <div className="ctx-sep" />
          {item(
            t(menu.noteRef.kind === 'footnote' ? 'appConvertToEndnote' : 'appConvertToFootnote'),
            {
              disabled: !canEdit,
              onClick: run(() => onConvertNote(menu.noteRef!.kind, menu.noteRef!.id)),
            },
          )}
        </>
      )}
      {editor.isActive('docListItem') && !!editor.getAttributes('docListItem').numId && (
        <>
          <div className="ctx-sep" />
          {item(t('appRestartNumbering'), {
            disabled: !onRestartNumbering,
            onClick: run(() => onRestartNumbering?.()),
          })}
          {item(t('appContinueNumbering'), {
            disabled: !onContinueNumbering,
            onClick: run(() => onContinueNumbering?.()),
          })}
        </>
      )}
      <div className="ctx-sep" />
      {item(t('appSynonyms'), {
        disabled: !synonymText,
        ai: true,
        onClick: run(() => onAiPreset(t('appSynonymsPrompt', { text: synonymText }))),
      })}
      <div className="ctx-item-wrap" onMouseLeave={() => setSubmenu(null)}>
        {item(t('appTranslate'), { disabled: !hasSelection, submenuKey: 'translate', ai: true })}
        {submenu === 'translate' && hasSelection && (
          <div className="ctx-submenu">
            {TRANSLATE_TARGETS.map((target) => (
              <button
                key={target.labelKey}
                className="ctx-item"
                onClick={run(() =>
                  onAiPreset(
                    t('appTranslateSelectionPrompt', {
                      lang: t(target.labelKey),
                      text: selectedText,
                    }),
                  ),
                )}
              >
                <span className="ctx-label">
                  {t('appTranslateTo', { lang: t(target.labelKey) })}
                </span>
              </button>
            ))}
          </div>
        )}
      </div>
      {isFloating && (
        <>
          <div className="ctx-sep" />
          <div className="ctx-item-wrap" onMouseLeave={() => setSubmenu(null)}>
            {item(t('appWrapTextMenu'), { submenuKey: 'wrap' })}
            {submenu === 'wrap' && (
              <div className="ctx-submenu">
                {WRAP_OPTIONS.map((opt) => (
                  <button
                    key={String(opt.value)}
                    className="ctx-item"
                    onClick={run(() => setWrap(opt.value))}
                  >
                    <span className="ctx-label">
                      {currentWrap === opt.value ? '✓ ' : ''}
                      {t(opt.labelKey)}
                    </span>
                  </button>
                ))}
              </div>
            )}
          </div>
          <div className="ctx-item-wrap" onMouseLeave={() => setSubmenu(null)}>
            {item(t('appArrangeMenu'), { submenuKey: 'arrange' })}
            {submenu === 'arrange' && (
              <div className="ctx-submenu">
                <button className="ctx-item" onClick={run(bringToFront)}>
                  <span className="ctx-label">{t('appBringToFront')}</span>
                </button>
                <button className="ctx-item" onClick={run(bringForward)}>
                  <span className="ctx-label">{t('appBringForward')}</span>
                </button>
                <button className="ctx-item" onClick={run(sendBackward)}>
                  <span className="ctx-label">{t('appSendBackward')}</span>
                </button>
                <button className="ctx-item" onClick={run(sendToBack)}>
                  <span className="ctx-label">{t('appSendToBack')}</span>
                </button>
              </div>
            )}
          </div>
        </>
      )}
      <div className="ctx-sep" />
      {item(t('appHyperlinkMenu'), { key: '⌘K', onClick: run(onLink) })}
      {item(t('appNewComment'), { disabled: !canComment, onClick: run(onNewComment) })}
    </div>
  )
}

/** Wrap options from Word's image context menu (inline/square/top-bottom/behind text/in front of text); shared with the Picture Format tab */
export const WRAP_OPTIONS: Array<{ labelKey: StringKey; value: string | null }> = [
  { labelKey: 'appWrapInline', value: null },
  { labelKey: 'appWrapSquareLeft', value: 'square-left' },
  { labelKey: 'appWrapSquareRight', value: 'square-right' },
  { labelKey: 'appWrapTopBottom', value: 'topBottom' },
  { labelKey: 'appWrapBehind', value: 'behind' },
  { labelKey: 'appWrapFront', value: 'front' },
]

/* ================= Font dialog ================= */

export { FontDialog } from './FontAdvancedDialog'

/* ================= Paragraph dialog ================= */

const PT_PER_TWIP = 1 / 20

type SpecialIndent = 'none' | 'first' | 'hanging'

const round2 = (v: number) => Math.round(v * 100) / 100

/** UI languages whose Word measures paragraph indents in characters (字元) */
const isCjkLang = (lang: string) =>
  lang === 'zh' || lang === 'zh-TW' || lang === 'ja' || lang === 'ko'

/** rendered font size (pt) at the caret — the width of one indent character */
function paraFontPt(editor: Editor): number {
  try {
    let node: Node | null = editor.view.domAtPos(editor.state.selection.from).node
    if (node && node.nodeType !== Node.ELEMENT_NODE) node = node.parentElement
    const px = node ? parseFloat(getComputedStyle(node as Element).fontSize) : NaN
    if (px > 0) return round2(px * 0.75)
  } catch {
    // detached view: fall through to the Word body default
  }
  return 12
}

type AlignValue = 'left' | 'center' | 'right' | 'justify'

/** visual alignment values; setSelectionAlign resolves them per paragraph direction */
const ALIGN_OPTIONS: Array<{ key: AlignValue; nameKey: StringKey }> = [
  { key: 'left', nameKey: 'appAlignLeft' },
  { key: 'center', nameKey: 'appAlignCenter' },
  { key: 'right', nameKey: 'appAlignRight' },
  { key: 'justify', nameKey: 'appAlignJustify' },
]

const LINE_SPACINGS: Array<{ value: number; nameKey: StringKey }> = [
  { value: 1, nameKey: 'appLineSingle' },
  { value: 1.15, nameKey: 'appLine115' },
  { value: 1.5, nameKey: 'appLine15' },
  { value: 2, nameKey: 'appLineDouble' },
  { value: 2.5, nameKey: 'appLine25' },
  { value: 3, nameKey: 'appLineTriple' },
]

type Compression = 'doNotCompress' | 'compressPunctuation' | 'compressPunctuationAndJapaneseKana'

export interface ParagraphDialogChange {
  from: number
  to: number
  attrs: Record<string, unknown>
  visualAlign?: AlignValue
  compression?: Compression
  defaultTabStopTwips?: number
}

export function ParagraphDialog({
  editor,
  onClose,
  compression,
  onCompressionChange,
  defaultTabStopTwips,
  onDefaultTabStopChange,
  onApplyAdvanced,
}: {
  editor: Editor
  onClose: () => void
  compression?: Compression | null
  onCompressionChange?: (value: Compression) => void
  defaultTabStopTwips?: number | null
  onDefaultTabStopChange?: (twips: number) => void
  onApplyAdvanced?: (change: ParagraphDialogChange) => void
}) {
  const { t } = useI18n()
  const modalKeys = useModalKeys(onClose)
  const [selection] = useState(() => ({
    from: editor.state.selection.from,
    to: editor.state.selection.to,
  }))
  const [samples] = useState(() => selectedParagraphs(editor, selection.from, selection.to))
  const attrs = samples[0]?.node.attrs ?? activeParaAttrs(editor)
  const [pane, setPane] = useState<'spacing' | 'breaks' | 'asian' | 'tabs'>('spacing')
  const [flags, setFlags] = useState<Record<string, boolean>>({})
  const [alignTouched, setAlignTouched] = useState(false)
  const flagValue = (key: string, fallback: boolean) => {
    const sample = commonParagraphValue(samples, (s) => s.node.attrs[key] as boolean | null)
    return {
      value: flags[key] ?? sample.value ?? fallback,
      mixed: key in flags ? false : sample.mixed,
    }
  }
  const setFlag = (key: string, value: boolean) => setFlags((old) => ({ ...old, [key]: value }))
  const check = (key: string, label: string, fallback = false) => {
    const state = flagValue(key, fallback)
    return (
      <label className="font-check" key={key}>
        <input
          type="checkbox"
          checked={state.value}
          ref={(el) => {
            if (el) el.indeterminate = state.mixed
          }}
          onChange={(event) => setFlag(key, event.target.checked)}
        />
        {label}
      </label>
    )
  }
  const [tabStops, setTabStops] = useState<TabStop[]>(() => {
    try {
      const parsed = JSON.parse(String(attrs.tabStops ?? '[]')) as TabStop[]
      return Array.isArray(parsed) ? parsed.filter((stop) => !stop.inherited && !stop.rel) : []
    } catch {
      return []
    }
  })
  const [tabsTouched, setTabsTouched] = useState(false)
  const [tabPosPt, setTabPosPt] = useState(36)
  const [tabAlign, setTabAlign] = useState<TabStop['val']>('left')
  const [tabLeader, setTabLeader] = useState<NonNullable<TabStop['leader']>>('none')
  const [textAlignment, setTextAlignment] = useState(String(attrs.textAlignment ?? 'auto'))
  const [compressionValue, setCompressionValue] = useState<Compression>(
    compression ?? 'doNotCompress',
  )
  const [compressionTouched, setCompressionTouched] = useState(false)
  const [defaultTabPt, setDefaultTabPt] = useState((defaultTabStopTwips ?? 720) / 20)
  const [defaultTabTouched, setDefaultTabTouched] = useState(false)
  const addTabStop = () => {
    const pos = Math.round(tabPosPt * 20)
    if (!Number.isFinite(pos) || pos < 0 || pos > 31680) return
    setTabStops((old) =>
      [
        ...old.filter((stop) => stop.pos !== pos),
        {
          pos,
          val: tabAlign,
          ...(tabLeader === 'none' ? {} : { leader: tabLeader }),
        },
      ].sort((a, b) => a.pos - b.pos),
    )
    setTabsTouched(true)
  }
  // unset align means "start": visually left in LTR, right in RTL (same as the ribbon)
  const [align, setAlign] = useState<AlignValue>(
    (attrs.align as AlignValue | null) ?? (attrs.bidi === true ? 'right' : 'left'),
  )
  // Line spacing rule: multiples are the preset select values; 'atLeast'/'exact'
  // take a pt value (w:spacing w:lineRule + w:line, already modeled by parse/save)
  const rawTwips = Number(attrs.lineRawTwips) || 0
  const initRule =
    attrs.lineRule === 'exact' || attrs.lineRule === 'atLeast' ? (attrs.lineRule as string) : ''
  const initMultiple =
    Number(attrs.lineSpacing) || (attrs.lineRule === 'auto' && rawTwips ? rawTwips / 240 : 1)
  const [lineRule, setLineRule] = useState<string>(initRule)
  const [lineSpacing, setLineSpacing] = useState<number>(Math.round(initMultiple * 100) / 100)
  const [linePt, setLinePt] = useState<number>(rawTwips ? Math.round(rawTwips / 2) / 10 : 12)
  const twipsToPt = (v: unknown) => round2((Number(v) || 0) * PT_PER_TWIP)
  // Word's indent block: Left / Right plus Special (none / first line /
  // hanging) By. East Asian Word measures indents in characters (1 ch = the
  // paragraph's font size, "2 字元" is the customary first-line indent);
  // elsewhere the default By is 0.5" (36pt). Left shows where the first
  // line starts: a hanging paragraph's w:left minus the hang (probed: Word
  // after ⌘T, w:left 480 w:hanging 480, shows Left 0, Hanging 24pt).
  const [unitTwips] = useState(() => (isCjkLang(getLang()) ? paraFontPt(editor) * 20 : 20))
  const charUnits = unitTwips !== 20
  const toUnits = (twips: number) => round2(twips / unitTwips)
  const initFirst = Number(attrs.indentFirstLine) || 0
  const initSpecial: SpecialIndent = initFirst > 0 ? 'first' : initFirst < 0 ? 'hanging' : 'none'
  const initLeft = (Number(attrs.indentLeft) || 0) - (initFirst < 0 ? -initFirst : 0)
  const initial = {
    left: toUnits(initLeft),
    right: toUnits(Number(attrs.indentRight) || 0),
    special: initSpecial,
    by: toUnits(Math.abs(initFirst)),
    spaceBefore: twipsToPt(attrs.spaceBefore),
    spaceAfter: twipsToPt(attrs.spaceAfter),
  }
  const [indentLeft, setIndentLeft] = useState(initial.left)
  const [indentRight, setIndentRight] = useState(initial.right)
  const [special, setSpecial] = useState<SpecialIndent>(initial.special)
  const [specialBy, setSpecialBy] = useState(initial.by)
  const [spaceBefore, setSpaceBefore] = useState(initial.spaceBefore)
  const [spaceAfter, setSpaceAfter] = useState(initial.spaceAfter)
  const pickSpecial = (v: SpecialIndent) => {
    setSpecial(v)
    if (v !== 'none' && specialBy <= 0) setSpecialBy(charUnits ? 2 : 36)
  }

  const apply = () => {
    if (!editor.isEditable) {
      onClose()
      return
    }
    const ptToTwips = (pt: number) => (pt > 0 ? Math.round(pt / PT_PER_TWIP) : null)
    const unitsToTwips = (u: number) => Math.round(u * unitTwips)
    // untouched fields keep their exact twips (no unit round-trip drift, and
    // an OK on an imported hanging paragraph never flattens it)
    const indentChanged =
      indentLeft !== initial.left ||
      special !== initial.special ||
      (special !== 'none' && specialBy !== initial.by)
    const by = special === 'none' ? 0 : unitsToTwips(specialBy)
    const indent: Record<string, number | null> = {}
    if (indentChanged) {
      const left = unitsToTwips(indentLeft) + (special === 'hanging' ? by : 0)
      indent.indentLeft = left > 0 ? left : null
      indent.indentFirstLine = by > 0 ? (special === 'hanging' ? -by : by) : null
    }
    if (indentRight !== initial.right) {
      indent.indentRight = indentRight > 0 ? unitsToTwips(indentRight) : null
    }
    const spacingAttrs: Record<string, number | null> = {}
    if (spaceBefore !== initial.spaceBefore) spacingAttrs.spaceBefore = ptToTwips(spaceBefore)
    if (spaceAfter !== initial.spaceAfter) spacingAttrs.spaceAfter = ptToTwips(spaceAfter)
    const spacing =
      lineRule === 'exact' || lineRule === 'atLeast'
        ? { lineSpacing: null, lineRule, lineRawTwips: Math.max(20, Math.round(linePt * 20)) }
        : {
            lineSpacing: lineSpacing === 1 ? null : lineSpacing,
            lineRule: null,
            lineRawTwips: null,
          }
    const changes: Record<string, unknown> = {
      ...(lineRule !== initRule ||
      lineSpacing !== Math.round(initMultiple * 100) / 100 ||
      linePt !== (rawTwips ? Math.round(rawTwips / 2) / 10 : 12)
        ? spacing
        : {}),
      ...indent,
      ...spacingAttrs,
      ...flags,
      ...('pageBreakBefore' in flags ? { pageBreakBeforeDirect: true } : {}),
      ...(textAlignment !== String(attrs.textAlignment ?? 'auto') ? { textAlignment } : {}),
      ...(tabsTouched ? { tabStops: JSON.stringify(tabStops) } : {}),
    }
    const compressionEdit =
      compressionTouched && compressionValue !== compression ? compressionValue : undefined
    const defaultTabEdit =
      defaultTabTouched &&
      defaultTabPt > 0 &&
      defaultTabPt <= 1584 &&
      Math.round(defaultTabPt * 20) !== defaultTabStopTwips
        ? Math.round(defaultTabPt * 20)
        : undefined
    if (onApplyAdvanced) {
      onApplyAdvanced({
        from: selection.from,
        to: selection.to,
        attrs: changes,
        ...(alignTouched ? { visualAlign: align } : {}),
        ...(compressionEdit ? { compression: compressionEdit } : {}),
        ...(defaultTabEdit ? { defaultTabStopTwips: defaultTabEdit } : {}),
      })
    } else {
      applyParagraphChanges(
        editor,
        selection.from,
        selection.to,
        changes,
        alignTouched ? align : undefined,
      )
      if (compressionEdit) onCompressionChange?.(compressionEdit)
      if (defaultTabEdit !== undefined) onDefaultTabStopChange?.(defaultTabEdit)
    }
    onClose()
  }

  const numInput = (
    label: string,
    value: number,
    set: (v: number) => void,
    unit: 'pt' | 'indent' = 'pt',
  ) => (
    <label>
      {label}
      <span className="para-num">
        <input
          type="number"
          min={0}
          max={unit === 'indent' && charUnits ? 100 : 400}
          step={unit === 'indent' && charUnits ? 0.5 : 1}
          value={value}
          onChange={(e) => set(Math.max(0, Number(e.target.value) || 0))}
        />
        <span className="para-unit">
          {unit === 'indent' && charUnits ? t('appUnitChars') : 'pt'}
        </span>
      </span>
    </label>
  )

  return (
    <div
      className="modal-backdrop"
      ref={modalKeys.ref}
      onKeyDown={modalKeys.onKeyDown}
      onMouseDown={(e) => e.target === e.currentTarget && onClose()}
    >
      <div className="modal para-advanced-dialog" data-testid="paragraph-dialog">
        <h2>{t('appParagraph')}</h2>
        <div className="font-dialog-tabs" role="tablist">
          {(
            [
              ['spacing', 'appParaTabSpacing'],
              ['breaks', 'appParaTabBreaks'],
              ['asian', 'appParaTabAsian'],
              ['tabs', 'appParaTabs'],
            ] as const
          ).map(([id, label]) => (
            <button
              key={id}
              type="button"
              role="tab"
              data-testid={`paragraph-pane-${id}`}
              aria-selected={pane === id}
              onClick={() => setPane(id)}
            >
              {t(label)}
            </button>
          ))}
        </div>
        {pane === 'spacing' && (
          <>
            <div className="font-dialog-row">
              <label>
                {t('appAlignment')}
                <Dropdown
                  value={align}
                  ariaLabel={t('appAlignment')}
                  options={ALIGN_OPTIONS.map((o) => ({ value: o.key, label: t(o.nameKey) }))}
                  onPick={(value) => {
                    setAlign(value)
                    setAlignTouched(true)
                  }}
                />
              </label>
              <label>
                {t('appLineSpacingLabel')}
                <Dropdown
                  value={lineRule || String(lineSpacing)}
                  ariaLabel={t('appLineSpacingLabel')}
                  options={[
                    ...(!lineRule && !LINE_SPACINGS.some((s) => s.value === lineSpacing)
                      ? [
                          {
                            value: String(lineSpacing),
                            label: t('appLineMultiple', { n: lineSpacing }),
                          },
                        ]
                      : []),
                    ...LINE_SPACINGS.map((s) => ({ value: String(s.value), label: t(s.nameKey) })),
                    { value: 'atLeast', label: t('appLineAtLeast') },
                    { value: 'exact', label: t('appLineExactly') },
                  ]}
                  onPick={(v) => {
                    if (v === 'exact' || v === 'atLeast') {
                      setLineRule(v)
                    } else {
                      setLineRule('')
                      setLineSpacing(Number(v))
                    }
                  }}
                />
              </label>
              {lineRule === 'exact' || lineRule === 'atLeast' ? (
                <label>
                  {t('appLineValue')}
                  <span className="para-num">
                    <input
                      type="number"
                      min={1}
                      max={1584}
                      step={0.5}
                      value={linePt}
                      onChange={(e) => setLinePt(Math.max(0, Number(e.target.value) || 0))}
                    />
                    <span className="para-unit">pt</span>
                  </span>
                </label>
              ) : (
                <label>
                  {t('appLineValue')}
                  <span className="para-num">
                    <input
                      type="number"
                      min={0.06}
                      max={132}
                      step={0.05}
                      value={lineSpacing}
                      onChange={(e) => setLineSpacing(Math.max(0.06, Number(e.target.value) || 1))}
                    />
                    <span className="para-unit">×</span>
                  </span>
                </label>
              )}
            </div>
            <div className="font-dialog-row">
              {numInput(t('appIndentLeft'), indentLeft, setIndentLeft, 'indent')}
              {numInput(t('appIndentRight'), indentRight, setIndentRight, 'indent')}
            </div>
            <div className="font-dialog-row">
              <label>
                {t('appIndentSpecial')}
                <Dropdown
                  value={special}
                  ariaLabel={t('appIndentSpecial')}
                  options={[
                    { value: 'none', label: t('appIndentSpecialNone') },
                    { value: 'first', label: t('appIndentSpecialFirst') },
                    { value: 'hanging', label: t('appIndentSpecialHanging') },
                  ]}
                  onPick={(v) => pickSpecial(v as SpecialIndent)}
                />
              </label>
              {special !== 'none' && numInput(t('appIndentBy'), specialBy, setSpecialBy, 'indent')}
            </div>
            <div className="font-dialog-row">
              {numInput(t('appSpaceBefore'), spaceBefore, setSpaceBefore)}
              {numInput(t('appSpaceAfter'), spaceAfter, setSpaceAfter)}
            </div>
          </>
        )}
        {pane === 'breaks' && (
          <div className="font-dialog-checks para-options">
            {check('widowControl', t('appParaWidow'), true)}
            {check('keepNext', t('appParaKeepNext'))}
            {check('keepLines', t('appParaKeepLines'))}
            {check('pageBreakBefore', t('appParaPageBreakBefore'))}
            {check('suppressLineNumbers', t('appParaSuppressLineNumbers'))}
            {check('suppressAutoHyphens', t('appParaDoNotHyphenate'))}
          </div>
        )}
        {pane === 'asian' && (
          <>
            <div className="font-dialog-checks para-options">
              {check('kinsoku', t('appParaKinsoku'), true)}
              {check('wordWrap', t('appParaLatinBreak'))}
              {check('overflowPunct', t('appParaHangPunct'), true)}
              {check('topLinePunct', t('appParaCompressLineStart'))}
              {check('autoSpaceDE', t('appParaAsianLatinSpace'), true)}
              {check('autoSpaceDN', t('appParaAsianDigitSpace'), true)}
            </div>
            <div className="font-dialog-row">
              <label>
                {t('appParaTextAlign')}
                <Dropdown
                  value={textAlignment}
                  ariaLabel={t('appParaTextAlign')}
                  options={(
                    [
                      ['auto', 'appParaTextAuto'],
                      ['baseline', 'appParaTextBaseline'],
                      ['center', 'appAlignCenter'],
                      ['top', 'appParaTextTop'],
                      ['bottom', 'appParaTextBottom'],
                    ] as const
                  ).map(([value, label]) => ({ value, label: t(label) }))}
                  onPick={(value) => setTextAlignment(value)}
                />
              </label>
              {(onCompressionChange || onApplyAdvanced) && (
                <label>
                  {t('appParaCompression')}
                  <Dropdown
                    value={compressionValue}
                    ariaLabel={t('appParaCompression')}
                    options={(
                      [
                        ['doNotCompress', 'appParaCompressNone'],
                        ['compressPunctuation', 'appParaCompressPunctuation'],
                        ['compressPunctuationAndJapaneseKana', 'appParaCompressKana'],
                      ] as const
                    ).map(([value, label]) => ({ value, label: t(label) }))}
                    onPick={(value) => {
                      setCompressionValue(value)
                      setCompressionTouched(true)
                    }}
                  />
                </label>
              )}
            </div>
          </>
        )}
        {pane === 'tabs' && (
          <>
            <div className="font-dialog-row">
              {(onDefaultTabStopChange || onApplyAdvanced) && (
                <label>
                  {t('appParaDefaultTab')}
                  <span className="para-num">
                    <input
                      type="number"
                      min="0.5"
                      max="1584"
                      step="0.5"
                      value={defaultTabPt}
                      onChange={(event) => {
                        setDefaultTabPt(Number(event.target.value))
                        setDefaultTabTouched(true)
                      }}
                    />
                    <span className="para-unit">pt</span>
                  </span>
                </label>
              )}
              <label>
                {t('appParaTabPosition')}
                <span className="para-num">
                  <input
                    type="number"
                    min="0"
                    max="1584"
                    step="0.5"
                    value={tabPosPt}
                    onChange={(event) => setTabPosPt(Number(event.target.value))}
                  />
                  <span className="para-unit">pt</span>
                </span>
              </label>
              <label>
                {t('appParaTabAlignment')}
                <Dropdown
                  value={tabAlign}
                  ariaLabel={t('appParaTabAlignment')}
                  options={(
                    [
                      ['left', 'appTabLeft'],
                      ['center', 'appTabCenter'],
                      ['right', 'appTabRight'],
                      ['decimal', 'appTabDecimal'],
                      ['bar', 'appTabBar'],
                    ] as const
                  ).map(([value, label]) => ({ value, label: t(label) }))}
                  onPick={(value) => setTabAlign(value as TabStop['val'])}
                />
              </label>
              <label>
                {t('appParaTabLeader')}
                <Dropdown
                  value={tabLeader}
                  ariaLabel={t('appParaTabLeader')}
                  options={(
                    [
                      ['none', 'appParaLeaderNone'],
                      ['dot', 'appParaLeaderDots'],
                      ['hyphen', 'appParaLeaderDashes'],
                      ['underscore', 'appParaLeaderUnderline'],
                      ['heavy', 'appParaLeaderHeavyDots'],
                    ] as const
                  ).map(([value, label]) => ({ value, label: t(label) }))}
                  onPick={(value) => setTabLeader(value as NonNullable<TabStop['leader']>)}
                />
              </label>
            </div>
            <div className="para-tab-actions">
              <button type="button" onClick={addTabStop}>
                {t('appParaTabAdd')}
              </button>
              <button
                type="button"
                onClick={() => {
                  setTabStops([])
                  setTabsTouched(true)
                }}
              >
                {t('appParaTabClearAll')}
              </button>
            </div>
            <div className="para-tab-list" role="list">
              {tabStops.map((stop) => (
                <div key={stop.pos} role="listitem">
                  <span>
                    {round2(stop.pos / 20)} pt ·{' '}
                    {t(
                      (
                        {
                          left: 'appTabLeft',
                          center: 'appTabCenter',
                          right: 'appTabRight',
                          decimal: 'appTabDecimal',
                          bar: 'appTabBar',
                          clear: 'appTabClear',
                        } as const
                      )[stop.val],
                    )}
                  </span>
                  <button
                    type="button"
                    onClick={() => {
                      setTabStops((old) => old.filter((item) => item.pos !== stop.pos))
                      setTabsTouched(true)
                    }}
                  >
                    {t('appParaTabRemove')}
                  </button>
                </div>
              ))}
            </div>
          </>
        )}
        <div className="modal-actions">
          <button className="btn-ghost" onClick={onClose}>
            {t('appCancel')}
          </button>
          <button className="btn-primary" onClick={apply}>
            {t('appOk')}
          </button>
        </div>
      </div>
    </div>
  )
}
