import { memo, useMemo, useState, type CSSProperties } from 'react'
import type { Editor } from '@tiptap/core'
import type { StyleInfo, StyleUpsert } from '@genoffice/docx-engine'
import {
  BUILTIN_PARA_STYLES,
  localizedStyleName,
  styleInfoOfUpsert,
} from '../editor/builtin-styles'
import { useI18n } from '../i18n/locale'
import { IconCaret, IconClose } from './icons'
import { applyParagraphStyle } from './ribbon-tabs'
import type { RibbonFormatState } from './ribbon-format-state'

export type StylesFilterMode = 'recommended' | 'inUse' | 'inDoc' | 'all'

export interface StylesPaneProps {
  editor: Editor
  styles: ReadonlyMap<string, StyleInfo>
  /** definitions waiting for the next save (shown even before they reach `styles`) */
  pending: Readonly<Record<string, StyleUpsert>>
  /** w:pStyle of the caret's paragraph, w:rStyle of the caret's run */
  paraStyleId: string | null
  charStyleId: string | null
  headingLevel: number | null
  formatState?: RibbonFormatState
  onEnsureStyle: (styleId: string) => void
  onUpdateStyleWithSelection?: (styleId: string) => void
  onUpsertStyle?: (style: StyleUpsert) => void
  onSelectAllWithStyle?: (styleId: string | null, headingLevel: number | null) => void
  onClose: () => void
}

/** the styles worth listing, in Word's order: Normal, headings, then the rest by name */
export function listStyles(
  styles: ReadonlyMap<string, StyleInfo>,
  pending: Readonly<Record<string, StyleUpsert>>,
  filter: StylesFilterMode = 'recommended',
  inUseStyleIds?: Set<string>,
): StyleInfo[] {
  const byId = new Map<string, StyleInfo>()
  for (const info of styles.values()) {
    if (info.type === 'table' || info.semiHidden || info.linkedCharShell) continue
    if (
      info.type === 'character' &&
      ['DefaultParagraphFont'].includes(info.styleId)
    )
      continue

    if (filter === 'recommended') {
      if (!info.qFormat && !info.headingLevel && !info.isDefault) continue
    } else if (filter === 'inUse') {
      if (!inUseStyleIds?.has(info.styleId) && !info.isDefault) continue
    }
    byId.set(info.styleId, info)
  }

  for (const up of Object.values(pending)) {
    if (!byId.has(up.styleId)) {
      if (filter === 'inUse' && !inUseStyleIds?.has(up.styleId)) continue
      byId.set(up.styleId, styleInfoOfUpsert(up))
    }
  }

  // built-ins the document does not define yet are still offered unless filtered out
  for (const b of BUILTIN_PARA_STYLES) {
    if (!byId.has(b.styleId)) {
      if (filter === 'inDoc') continue
      if (filter === 'inUse' && !inUseStyleIds?.has(b.styleId)) continue
      byId.set(b.styleId, styleInfoOfUpsert(b.upsert))
    }
  }

  const rank = (s: StyleInfo) => (s.isDefault ? 0 : s.headingLevel ? 1 : 2)
  return [...byId.values()].sort(
    (a, b) =>
      rank(a) - rank(b) ||
      (a.headingLevel ?? 0) - (b.headingLevel ?? 0) ||
      a.name.localeCompare(b.name),
  )
}

const previewStyle = (info: StyleInfo): CSSProperties => {
  const d = info.display
  const css: CSSProperties = {}
  if (d?.bold) css.fontWeight = 'bold'
  if (d?.italic) css.fontStyle = 'italic'
  if (d?.underline) css.textDecoration = 'underline'
  if (d?.color && d.color !== 'auto') css.color = `#${d.color}`
  // the pane keeps rows one line tall: the style's size only shows between 11 and 17 px
  if (d?.sizeHalfPoints) css.fontSize = Math.min(17, Math.max(11, d.sizeHalfPoints / 2))
  if (d?.fontAscii ?? d?.font) css.fontFamily = d?.fontAscii ?? d?.font ?? ''
  if (info.styleId === 'IntenseQuote') {
    css.borderTop = '1px solid #4472C4'
    css.borderBottom = '1px solid #4472C4'
    css.padding = '1px 0'
  }
  return css
}

interface StyleEditorModalProps {
  initial?: Partial<StyleUpsert>
  lang: string
  onSave: (upsert: StyleUpsert) => void
  onClose: () => void
}

function StyleEditorModal({ initial, lang, onSave, onClose }: StyleEditorModalProps) {
  const isZh = lang.startsWith('zh')
  const [name, setName] = useState(initial?.name ?? (isZh ? '新樣式' : 'New Style'))
  const [type, setType] = useState<'paragraph' | 'character'>(
    initial?.type === 'character' ? 'character' : 'paragraph',
  )
  const [font, setFont] = useState(initial?.rPr?.font ?? (isZh ? '微軟正黑體' : 'Calibri'))
  const [fontSizePt, setFontSizePt] = useState(
    initial?.rPr?.sizeHalfPoints ? initial.rPr.sizeHalfPoints / 2 : 12,
  )
  const [bold, setBold] = useState(initial?.rPr?.bold ?? false)
  const [italic, setItalic] = useState(initial?.rPr?.italic ?? false)
  const [underline, setUnderline] = useState(initial?.rPr?.underline ?? false)
  const [align, setAlign] = useState<'left' | 'center' | 'right' | 'justify'>(
    (initial?.pPr?.align as never) ?? (isZh ? 'justify' : 'left'),
  )
  const [lineSpacing, setLineSpacing] = useState(initial?.pPr?.lineSpacing ?? 1.5)
  const [firstLineCm, setFirstLineCm] = useState(
    initial?.pPr?.firstLineTwips ? (initial.pPr.firstLineTwips / 567).toFixed(2) : (isZh ? '0.85' : '0'),
  )

  const handleSave = () => {
    const styleId = initial?.styleId || `UserStyle_${Date.now()}`
    const firstLineTwips = Math.round((parseFloat(firstLineCm) || 0) * 567)
    const upsert: StyleUpsert = {
      styleId,
      name: name.trim() || styleId,
      type,
      basedOn: type === 'paragraph' ? 'Normal' : undefined,
      next: type === 'paragraph' ? 'Normal' : undefined,
      quickFormat: true,
      pPr:
        type === 'paragraph'
          ? {
              align,
              lineSpacing,
              firstLineTwips: firstLineTwips > 0 ? firstLineTwips : undefined,
            }
          : undefined,
      rPr: {
        font,
        sizeHalfPoints: Math.round(fontSizePt * 2),
        bold,
        italic,
        underline,
      },
    }
    onSave(upsert)
    onClose()
  }

  return (
    <div className="modal-backdrop" onMouseDown={(e) => e.target === e.currentTarget && onClose()}>
      <div
        className="modal"
        role="dialog"
        aria-modal="true"
        style={{ width: 380, padding: 20, display: 'flex', flexDirection: 'column', gap: 14 }}
      >
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
          <h2 style={{ margin: 0, fontSize: 16 }}>
            {initial?.styleId ? (isZh ? '修改樣式' : 'Modify Style') : (isZh ? '新增樣式' : 'New Style')}
          </h2>
          <button className="nav-pane-close" onClick={onClose}>
            <IconClose size={14} />
          </button>
        </div>

        <div style={{ display: 'flex', flexDirection: 'column', gap: 10, fontSize: 13 }}>
          <label style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
            <span>{isZh ? '名稱:' : 'Name:'}</span>
            <input
              type="text"
              value={name}
              onChange={(e) => setName(e.target.value)}
              style={{ padding: '4px 8px', borderRadius: 4, border: '1px solid var(--border)' }}
            />
          </label>

          <label style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
            <span>{isZh ? '樣式類型:' : 'Style type:'}</span>
            <select
              value={type}
              onChange={(e) => setType(e.target.value as 'paragraph' | 'character')}
              style={{ padding: '4px 8px', borderRadius: 4, border: '1px solid var(--border)' }}
            >
              <option value="paragraph">{isZh ? '段落 (¶)' : 'Paragraph (¶)'}</option>
              <option value="character">{isZh ? '字元 (a)' : 'Character (a)'}</option>
            </select>
          </label>

          <div style={{ display: 'flex', gap: 10 }}>
            <label style={{ flex: 1, display: 'flex', flexDirection: 'column', gap: 4 }}>
              <span>{isZh ? '字體:' : 'Font:'}</span>
              <input
                type="text"
                value={font}
                onChange={(e) => setFont(e.target.value)}
                style={{ padding: '4px 8px', borderRadius: 4, border: '1px solid var(--border)' }}
              />
            </label>
            <label style={{ width: 80, display: 'flex', flexDirection: 'column', gap: 4 }}>
              <span>{isZh ? '字級 (pt):' : 'Size (pt):'}</span>
              <input
                type="number"
                min="6"
                max="72"
                value={fontSizePt}
                onChange={(e) => setFontSizePt(Number(e.target.value) || 12)}
                style={{ padding: '4px 8px', borderRadius: 4, border: '1px solid var(--border)' }}
              />
            </label>
          </div>

          <div style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
            <label style={{ display: 'flex', alignItems: 'center', gap: 4, cursor: 'pointer' }}>
              <input type="checkbox" checked={bold} onChange={(e) => setBold(e.target.checked)} />
              <span>{isZh ? '粗體' : 'Bold'}</span>
            </label>
            <label style={{ display: 'flex', alignItems: 'center', gap: 4, cursor: 'pointer' }}>
              <input type="checkbox" checked={italic} onChange={(e) => setItalic(e.target.checked)} />
              <span>{isZh ? '斜體' : 'Italic'}</span>
            </label>
            <label style={{ display: 'flex', alignItems: 'center', gap: 4, cursor: 'pointer' }}>
              <input
                type="checkbox"
                checked={underline}
                onChange={(e) => setUnderline(e.target.checked)}
              />
              <span>{isZh ? '底線' : 'Underline'}</span>
            </label>
          </div>

          {type === 'paragraph' && (
            <>
              <div style={{ display: 'flex', gap: 10 }}>
                <label style={{ flex: 1, display: 'flex', flexDirection: 'column', gap: 4 }}>
                  <span>{isZh ? '對齊方式:' : 'Alignment:'}</span>
                  <select
                    value={align}
                    onChange={(e) => setAlign(e.target.value as never)}
                    style={{ padding: '4px 8px', borderRadius: 4, border: '1px solid var(--border)' }}
                  >
                    <option value="left">{isZh ? '靠左對齊' : 'Left'}</option>
                    <option value="center">{isZh ? '置中' : 'Center'}</option>
                    <option value="right">{isZh ? '靠右對齊' : 'Right'}</option>
                    <option value="justify">{isZh ? '左右對齊' : 'Justified'}</option>
                  </select>
                </label>
                <label style={{ flex: 1, display: 'flex', flexDirection: 'column', gap: 4 }}>
                  <span>{isZh ? '行距:' : 'Line spacing:'}</span>
                  <select
                    value={lineSpacing}
                    onChange={(e) => setLineSpacing(Number(e.target.value) || 1)}
                    style={{ padding: '4px 8px', borderRadius: 4, border: '1px solid var(--border)' }}
                  >
                    <option value={1}>1.0</option>
                    <option value={1.15}>1.15</option>
                    <option value={1.5}>1.5</option>
                    <option value={2}>2.0</option>
                  </select>
                </label>
              </div>

              <label style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
                <span>{isZh ? '第一行縮排 (公分):' : 'First line indent (cm):'}</span>
                <input
                  type="text"
                  value={firstLineCm}
                  onChange={(e) => setFirstLineCm(e.target.value)}
                  placeholder={isZh ? '例如: 0.85 (約 2 字元)' : 'e.g. 0.85'}
                  style={{ padding: '4px 8px', borderRadius: 4, border: '1px solid var(--border)' }}
                />
              </label>
            </>
          )}
        </div>

        <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 8, marginTop: 8 }}>
          <button className="btn-ghost" onClick={onClose}>
            {isZh ? '取消' : 'Cancel'}
          </button>
          <button className="btn-primary" onClick={handleSave}>
            {isZh ? '確定' : 'OK'}
          </button>
        </div>
      </div>
    </div>
  )
}

export const StylesPane = memo(function StylesPane({
  editor,
  styles,
  pending,
  paraStyleId,
  charStyleId,
  headingLevel,
  formatState,
  onEnsureStyle,
  onUpdateStyleWithSelection,
  onUpsertStyle,
  onSelectAllWithStyle,
  onClose,
}: StylesPaneProps) {
  const { t, lang } = useI18n()
  const isZh = lang.startsWith('zh')
  const [filter, setFilter] = useState<StylesFilterMode>('recommended')
  const [activeMenuStyleId, setActiveMenuStyleId] = useState<string | null>(null)
  const [editingStyle, setEditingStyle] = useState<Partial<StyleUpsert> | null>(null)

  // Collect styles actively used across the document
  const inUseStyleIds = useMemo(() => {
    const set = new Set<string>()
    editor.state.doc.descendants((node) => {
      if (node.isBlock) {
        if (node.attrs.styleId) set.add(String(node.attrs.styleId))
        if (node.type.name === 'docHeading' && node.attrs.level) {
          set.add(`Heading${node.attrs.level}`)
          set.add(`heading ${node.attrs.level}`)
        }
      }
      for (const mark of node.marks) {
        if (mark.type.name === 'docTextStyle' && mark.attrs.styleId) {
          set.add(String(mark.attrs.styleId))
        }
      }
    })
    return set
  }, [editor.state.doc])

  const list = useMemo(
    () => listStyles(styles, pending, filter, inUseStyleIds),
    [styles, pending, filter, inUseStyleIds],
  )

  const isActive = (s: StyleInfo): boolean =>
    s.type === 'character'
      ? charStyleId === s.styleId
      : s.headingLevel && !s.headingLevelInherited
        ? headingLevel === s.headingLevel
        : s.isDefault
          ? !paraStyleId && headingLevel === null
          : paraStyleId === s.styleId

  const apply = (s: StyleInfo) => {
    setActiveMenuStyleId(null)
    if (s.type === 'character') {
      if (charStyleId === s.styleId) editor.chain().focus().unsetMark('docTextStyle').run()
      else editor.chain().focus().setMark('docTextStyle', { styleId: s.styleId }).run()
      return
    }
    if (s.isDefault) {
      applyParagraphStyle(editor, 'p')
      return
    }
    if (s.headingLevel && !s.headingLevelInherited && s.headingLevel <= 3) {
      applyParagraphStyle(editor, `h${s.headingLevel}` as 'h1' | 'h2' | 'h3')
      return
    }
    onEnsureStyle(s.styleId)
    applyParagraphStyle(editor, { styleId: s.styleId })
  }

  // Format description of the active selection
  const currentFormatDesc = useMemo(() => {
    let baseName = ''
    if (headingLevel) {
      baseName = localizedStyleName(`Heading${headingLevel}`, `Heading ${headingLevel}`, lang)
    } else if (paraStyleId) {
      const s = styles.get(paraStyleId) ?? pending[paraStyleId]
      baseName = s ? localizedStyleName(s.styleId, s.name ?? s.styleId, lang) : paraStyleId
    } else {
      baseName = localizedStyleName('Normal', 'Normal', lang)
    }

    const overrides: string[] = []
    if (formatState) {
      if (formatState.align === 'justify') overrides.push(isZh ? '左右對齊' : 'Justified')
      else if (formatState.align === 'center') overrides.push(isZh ? '置中' : 'Centered')
      else if (formatState.align === 'right') overrides.push(isZh ? '靠右對齊' : 'Right-aligned')

      if (formatState.indentFirstLine) {
        const cm = (formatState.indentFirstLine / 567).toFixed(2)
        overrides.push(isZh ? `第一行: ${cm} 公分` : `First line: ${cm} cm`)
      }
      if (formatState.lineSpacing && formatState.lineSpacing !== 1) {
        overrides.push(isZh ? `行距: ${formatState.lineSpacing} 行` : `Line spacing: ${formatState.lineSpacing}`)
      }
      if (formatState.spaceBefore > 0) {
        const pt = Math.round(formatState.spaceBefore / 20)
        overrides.push(isZh ? `與前段距離: ${pt} 點` : `Space before: ${pt} pt`)
      }
      if (formatState.spaceAfter > 0) {
        const pt = Math.round(formatState.spaceAfter / 20)
        overrides.push(isZh ? `與後段距離: ${pt} 點` : `Space after: ${pt} pt`)
      }
    }

    if (overrides.length > 0) {
      return `${baseName} + ${overrides.join(', ')}`
    }
    return baseName
  }, [paraStyleId, headingLevel, formatState, styles, pending, lang, isZh])

  return (
    <aside className="styles-pane" aria-label={t('ribbonStylesPane')}>
      <div className="nav-pane-head">
        <span className="nav-pane-title">{isZh ? '樣式' : t('ribbonStylesPane')}</span>
        <button
          className="nav-pane-close"
          aria-label={t('appClose')}
          data-tip={t('appClose')}
          onClick={onClose}
        >
          <IconClose size={14} />
        </button>
      </div>

      {/* Current Style section matching Word Styles Pane */}
      <div className="styles-pane-section">
        <span className="styles-pane-label">{isZh ? '目前樣式:' : 'Current style:'}</span>
        <div className="styles-pane-current-box" title={currentFormatDesc}>
          {currentFormatDesc}
        </div>
        <div className="styles-pane-btn-row">
          <button
            className="styles-pane-btn"
            onClick={() => setEditingStyle({})}
          >
            {isZh ? '新增樣式...' : 'New Style...'}
          </button>
          <button
            className="styles-pane-btn"
            onClick={() => onSelectAllWithStyle?.(paraStyleId, headingLevel)}
          >
            {isZh ? '全選' : 'Select All'}
          </button>
        </div>
      </div>

      <div className="styles-pane-apply-header">
        <span className="styles-pane-label">{isZh ? '套用樣式:' : 'Pick formatting to apply:'}</span>
      </div>

      <div className="styles-pane-list">
        {list.map((s) => {
          const active = isActive(s)
          const menuOpen = activeMenuStyleId === s.styleId
          const displayName = localizedStyleName(s.styleId, s.name, lang)

          return (
            <div
              key={`${s.type}:${s.styleId}`}
              className={`styles-pane-item-wrap${active ? ' active' : ''}`}
            >
              <button
                className="styles-pane-item"
                aria-pressed={active}
                onMouseDown={(e) => e.preventDefault()}
                onClick={() => apply(s)}
              >
                <span className="styles-pane-name" style={previewStyle(s)}>
                  {displayName}
                </span>
                <span className="styles-pane-kind" aria-hidden>
                  {s.type === 'character' ? 'a' : '¶'}
                </span>
              </button>

              <button
                className="styles-pane-caret-btn"
                aria-label={isZh ? `操作 ${displayName}` : `Actions for ${displayName}`}
                onClick={(e) => {
                  e.stopPropagation()
                  setActiveMenuStyleId(menuOpen ? null : s.styleId)
                }}
              >
                <IconCaret />
              </button>

              {menuOpen && (
                <div
                  className="styles-pane-popover"
                  onMouseDown={(e) => e.stopPropagation()}
                >
                  <button onClick={() => apply(s)}>
                    {isZh ? '套用' : 'Apply'}
                  </button>
                  <button
                    onClick={() => {
                      setActiveMenuStyleId(null)
                      onUpdateStyleWithSelection?.(s.styleId)
                    }}
                  >
                    {isZh ? `以選取範圍更新 ${displayName}` : `Update ${displayName} to Match Selection`}
                  </button>
                  <button
                    onClick={() => {
                      setActiveMenuStyleId(null)
                      setEditingStyle({
                        styleId: s.styleId,
                        name: s.name,
                        type: s.type,
                        pPr: s.display
                          ? {
                              align: s.display.align as never,
                              lineSpacing: s.display.lineSpacing,
                              firstLineTwips: s.display.indentFirstLineTwips,
                            }
                          : undefined,
                        rPr: s.display
                          ? {
                              font: s.display.fontAscii ?? s.display.font,
                              sizeHalfPoints: s.display.sizeHalfPoints,
                              bold: s.display.bold,
                              italic: s.display.italic,
                              underline: s.display.underline,
                            }
                          : undefined,
                      })
                    }}
                  >
                    {isZh ? '修改樣式...' : 'Modify Style...'}
                  </button>
                  <button
                    onClick={() => {
                      setActiveMenuStyleId(null)
                      onSelectAllWithStyle?.(s.styleId, s.headingLevel ?? null)
                    }}
                  >
                    {isZh ? '選取所有相符項' : 'Select All Instances'}
                  </button>
                </div>
              )}
            </div>
          )
        })}
      </div>

      {/* Filter footer matching Word Styles Pane */}
      <div className="styles-pane-footer">
        <span>{isZh ? '清單:' : 'List:'}</span>
        <select
          value={filter}
          onChange={(e) => setFilter(e.target.value as StylesFilterMode)}
        >
          <option value="recommended">{isZh ? '建議使用' : 'Recommended'}</option>
          <option value="inUse">{isZh ? '使用中' : 'In use'}</option>
          <option value="inDoc">{isZh ? '在目前文件中' : 'In current document'}</option>
          <option value="all">{isZh ? '所有樣式' : 'All styles'}</option>
        </select>
      </div>

      {editingStyle && (
        <StyleEditorModal
          initial={editingStyle}
          lang={lang}
          onSave={(up) => onUpsertStyle?.(up)}
          onClose={() => setEditingStyle(null)}
        />
      )}
    </aside>
  )
})
