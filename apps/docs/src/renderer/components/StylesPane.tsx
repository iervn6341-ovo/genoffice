import { memo, type CSSProperties } from 'react'
import type { Editor } from '@tiptap/core'
import type { StyleInfo, StyleUpsert } from '@genoffice/docx-engine'
import { BUILTIN_PARA_STYLES, styleInfoOfUpsert } from '../editor/builtin-styles'
import { useI18n } from '../i18n/locale'
import { IconClose } from './icons'
import { applyParagraphStyle } from './ribbon-tabs'

/**
 * Word's Styles pane: every paragraph and character style the document offers (qFormat styles, the
 * headings, and the built-ins that are not defined yet), the one at the caret highlighted. A click
 * applies it like the gallery does; a built-in that the file lacks is defined first.
 */
export interface StylesPaneProps {
  editor: Editor
  styles: ReadonlyMap<string, StyleInfo>
  /** definitions waiting for the next save (shown even before they reach `styles`) */
  pending: Readonly<Record<string, StyleUpsert>>
  /** w:pStyle of the caret's paragraph, w:rStyle of the caret's run */
  paraStyleId: string | null
  charStyleId: string | null
  headingLevel: number | null
  onEnsureStyle: (styleId: string) => void
  onClose: () => void
}

/** the styles worth listing, in Word's order: Normal, headings, then the rest by name */
export function listStyles(
  styles: ReadonlyMap<string, StyleInfo>,
  pending: Readonly<Record<string, StyleUpsert>>,
): StyleInfo[] {
  const byId = new Map<string, StyleInfo>()
  for (const info of styles.values()) {
    if (info.type === 'table' || info.semiHidden || info.linkedCharShell) continue
    if (
      info.type === 'character' &&
      ['DefaultParagraphFont', 'Hyperlink', 'FollowedHyperlink'].includes(info.styleId)
    )
      continue
    // Word lists the quick-format styles, the headings and the one every paragraph falls back to
    if (!info.qFormat && !info.headingLevel && !info.isDefault) continue
    byId.set(info.styleId, info)
  }
  for (const up of Object.values(pending))
    if (!byId.has(up.styleId)) byId.set(up.styleId, styleInfoOfUpsert(up))
  // built-ins the document does not define yet are still offered
  for (const b of BUILTIN_PARA_STYLES)
    if (!byId.has(b.styleId)) byId.set(b.styleId, styleInfoOfUpsert(b.upsert))
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
  return css
}

export const StylesPane = memo(function StylesPane({
  editor,
  styles,
  pending,
  paraStyleId,
  charStyleId,
  headingLevel,
  onEnsureStyle,
  onClose,
}: StylesPaneProps) {
  const { t } = useI18n()
  const list = listStyles(styles, pending)

  const isActive = (s: StyleInfo): boolean =>
    s.type === 'character'
      ? charStyleId === s.styleId
      : s.headingLevel && !s.headingLevelInherited
        ? headingLevel === s.headingLevel
        : s.isDefault
          ? !paraStyleId && headingLevel === null
          : paraStyleId === s.styleId

  const apply = (s: StyleInfo) => {
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

  return (
    <aside className="styles-pane" aria-label={t('ribbonStylesPane')}>
      <div className="nav-pane-head">
        <span className="nav-pane-title">{t('ribbonStylesPane')}</span>
        <button
          className="nav-pane-close"
          aria-label={t('appClose')}
          data-tip={t('appClose')}
          onClick={onClose}
        >
          <IconClose size={14} />
        </button>
      </div>
      <div className="styles-pane-list">
        {list.map((s) => (
          <button
            key={`${s.type}:${s.styleId}`}
            className={`styles-pane-item${isActive(s) ? ' active' : ''}`}
            aria-pressed={isActive(s)}
            onMouseDown={(e) => e.preventDefault()}
            onClick={() => apply(s)}
          >
            <span className="styles-pane-name" style={previewStyle(s)}>
              {s.name.charAt(0).toUpperCase() + s.name.slice(1)}
            </span>
            <span className="styles-pane-kind" aria-hidden>
              {s.type === 'character' ? 'a' : '¶'}
            </span>
          </button>
        ))}
      </div>
    </aside>
  )
})
