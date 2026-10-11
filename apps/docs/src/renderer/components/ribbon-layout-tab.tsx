import { useRef, useState } from 'react'
import type { DocFlags } from '../editor/layout-history'
import { type SectionSettings } from '@genoffice/docx-engine'
import { WRAP_OPTIONS } from './ContextMenu'
import {
  ColumnsDialog,
  DEFAULT_LINE_NUMBERING,
  LineNumbersDialog,
  PaperSizeDialog,
  columnPresetWidths,
  withColumns,
} from './LayoutDialogs'
import { MarginDialog, cmFromTwips, marginsFitPage, type PageMargins } from './MarginDialog'
import { useI18n, type StringKey } from '../i18n/locale'
import {
  IconAlignObjects,
  IconBreaks,
  IconBringForward,
  IconCaret,
  IconColumns,
  IconHyphenation,
  IconLineNumbers,
  IconMargins,
  IconOrientation,
  IconPageSize,
  IconParaSpacing,
  IconPosition,
  IconSendBackward,
  IconWrapText,
} from './icons'
import { RibbonFoldGroup } from '@genoffice/ui'

/** icon size for the big icon-over-label ribbon buttons (slides ribbon parity) */
import {
  BIG,
  TabProps,
  activeParaAttrs,
  insertPageBreakAt,
  setParaAttrs,
  toggleDropdown,
} from './ribbon-tabs'

export const MARGIN_PRESETS: Array<{
  key: string
  nameKey: StringKey
  descKey: StringKey
  top: number
  right: number
  bottom: number
  left: number
}> = [
  {
    key: 'normal',
    nameKey: 'ribbonMarginNormal',
    descKey: 'ribbonMarginNormalDesc',
    top: 1440,
    right: 1440,
    bottom: 1440,
    left: 1440,
  },
  {
    key: 'narrow',
    nameKey: 'ribbonMarginNarrow',
    descKey: 'ribbonMarginNarrowDesc',
    top: 720,
    right: 720,
    bottom: 720,
    left: 720,
  },
  {
    key: 'moderate',
    nameKey: 'ribbonMarginModerate',
    descKey: 'ribbonMarginModerateDesc',
    top: 1440,
    right: 1080,
    bottom: 1440,
    left: 1080,
  },
  {
    key: 'wide',
    nameKey: 'ribbonMarginWide',
    descKey: 'ribbonMarginWideDesc',
    top: 1440,
    right: 2880,
    bottom: 1440,
    left: 2880,
  },
]

/** Word's Mirrored preset: left = inside (binding side), right = outside, the flag swaps them on even pages */
const MIRRORED_MARGINS: PageMargins = { top: 1440, right: 1440, bottom: 1440, left: 1800 }

const LAST_CUSTOM_MARGINS_KEY = 'aidocs.marginLastCustom'

function readLastCustomMargins(): PageMargins | null {
  try {
    const raw = localStorage.getItem(LAST_CUSTOM_MARGINS_KEY)
    if (!raw) return null
    const v = JSON.parse(raw) as PageMargins
    const sides = [v.top, v.right, v.bottom, v.left]
    return sides.every((n) => Number.isFinite(n) && n >= 0) ? v : null
  } catch {
    return null
  }
}

export const PAPER_SIZES = [
  { key: 'a4', name: 'A4', desc: '21 × 29.7 cm', w: 11906, h: 16838 },
  { key: 'letter', name: 'Letter', desc: '21.59 × 27.94 cm', w: 12240, h: 15840 },
  { key: 'legal', name: 'Legal', desc: '21.59 × 35.56 cm', w: 12240, h: 20160 },
  { key: 'b5', name: 'B5 (JIS)', desc: '18.2 × 25.7 cm', w: 10319, h: 14572 },
  { key: 'a3', name: 'A3', desc: '29.7 × 42 cm', w: 16838, h: 23811 },
  { key: 'a5', name: 'A5', desc: '14.8 × 21 cm', w: 8391, h: 11906 },
  { key: 'b4', name: 'B4 (JIS)', desc: '25.7 × 36.4 cm', w: 14570, h: 20636 },
  { key: 'executive', name: 'Executive', desc: '18.41 × 26.67 cm', w: 10440, h: 15120 },
  { key: 'tabloid', name: 'Tabloid', desc: '27.94 × 43.18 cm', w: 15840, h: 24480 },
]

interface LayoutTabProps extends TabProps {
  section: SectionSettings | null
  onSection: (next: SectionSettings, flags?: Partial<DocFlags>) => void
  /** Multi-section documents: index of the cursor's section (0-based); null for single-section */
  activeSection: number | null
  onInsertSectionBreak: (type: 'nextPage' | 'continuous' | 'evenPage' | 'oddPage') => void
  /** document-wide settings.xml flags (hyphenation, mirrored margins) */
  docFlags: DocFlags
  onDocFlags: (next: Partial<DocFlags>) => void
}

const PT_PER_TWIP = 1 / 20

export function LayoutTab({
  editor,
  hasDoc,
  dropdown,
  setDropdown,
  section,
  onSection,
  activeSection,
  onInsertSectionBreak,
  docFlags,
  onDocFlags,
}: LayoutTabProps) {
  const { t } = useI18n()
  const paraAttrs = activeParaAttrs(editor)
  const enabled = hasDoc && !!section
  const [marginDialog, setMarginDialog] = useState(false)
  const [columnsDialog, setColumnsDialog] = useState(false)
  const [paperDialog, setPaperDialog] = useState(false)
  const [lineNumDialog, setLineNumDialog] = useState(false)

  /**
   * Margins are totals on the binding side (a gutter is folded into the left margin, see
   * SectionSettings.gutter). Presets clear the gutter and the mirror flag like Word's do.
   */
  const applyMargins = (m: PageMargins, opts: { gutter?: number; mirror?: boolean } = {}) => {
    if (!section || !marginsFitPage(m, section.pageWidth, section.pageHeight)) return
    const gutter = opts.gutter ?? 0
    const { gutter: previousGutter, ...rest } = section
    const mirror = opts.mirror ?? false
    onSection(
      {
        ...rest,
        marginTop: m.top,
        marginRight: m.right,
        marginBottom: m.bottom,
        marginLeft: m.left,
        // a user-set value is an ordinary margin, not the file's header-proof fixed one
        marginTopFixed: undefined,
        marginBottomFixed: undefined,
        ...(gutter > 0 ? { gutter } : previousGutter !== undefined ? { gutter: 0 } : {}),
      },
      mirror !== docFlags.mirrorMargins ? { mirrorMargins: mirror } : undefined,
    )
  }

  const marginsActive = (m: PageMargins, mirror = false) =>
    !!section &&
    docFlags.mirrorMargins === mirror &&
    !section.gutter &&
    section.marginTop === m.top &&
    section.marginRight === m.right &&
    section.marginBottom === m.bottom &&
    section.marginLeft === m.left

  // Arrange group: enabled when a floating object (image/textbox) is selected; maps to Word's Position / Wrap Text
  const protAttrs = editor.getAttributes('docProtected')
  const isImage = protAttrs?.blockType === 'image' && !!protAttrs.imageDataUrl
  const isFloatingBox =
    Array.isArray(protAttrs?.textboxes) && (protAttrs.textboxes as unknown[]).length > 0
  const canWrap = hasDoc && (isImage || isFloatingBox)
  // position presets go through imagePatchOf (original-document images only); newly inserted objects work after saving
  const canPosition = hasDoc && isImage && protAttrs?.docxIndex != null
  const currentWrap = (protAttrs?.imageWrap as string | null) ?? null

  const applyWrap = (value: string | null) => {
    const cleared =
      value === null
        ? { imagePosH: null, imagePosV: null, imageOffsetXEmu: null, imageOffsetYEmu: null }
        : {}
    editor
      .chain()
      .focus()
      .updateAttributes('docProtected', { imageWrap: value, ...cleared })
      .run()
    setDropdown(() => null)
  }

  const applyPositionPreset = (h: 'left' | 'center' | 'right', v: 'top' | 'center' | 'bottom') => {
    editor
      .chain()
      .focus()
      .updateAttributes('docProtected', {
        imageWrap: h === 'right' ? 'square-right' : 'square-left',
        imagePosH: h,
        imagePosV: v,
        imageOffsetXEmu: null,
        imageOffsetYEmu: null,
      })
      .run()
    setDropdown(() => null)
  }

  const applyInlinePosition = () => applyWrap(null)

  /**
   * Arrange ▾ Align: page-relative placement of a floating picture on one axis, keeping the other
   * (a picture still in line with text becomes square-wrapped first, like Position does)
   */
  const alignObject = (axis: 'h' | 'v', value: 'left' | 'center' | 'right' | 'top' | 'bottom') => {
    const h = axis === 'h' ? value : ((protAttrs?.imagePosH as string | undefined) ?? 'left')
    const v = axis === 'v' ? value : ((protAttrs?.imagePosV as string | undefined) ?? 'top')
    applyPositionPreset(h as 'left' | 'center' | 'right', v as 'top' | 'center' | 'bottom')
  }

  const setOrientation = (orientation: 'portrait' | 'landscape') => {
    if (!section || section.orientation === orientation) return
    onSection({
      ...section,
      orientation,
      pageWidth: section.pageHeight,
      pageHeight: section.pageWidth,
    })
    setDropdown(() => null)
  }

  const setPaper = (w: number, h: number) => {
    if (!section) return
    const landscape = section.orientation === 'landscape'
    onSection({ ...section, pageWidth: landscape ? h : w, pageHeight: landscape ? w : h })
    setDropdown(() => null)
  }

  /** selection captured when a pt field takes focus (commit target on blur) */
  const ptInputTargetRef = useRef<{ from: number; to: number } | null>(null)
  /** min < 0 only for the indent fields: Word permits negative indents (text
   *  into the margin, r116/r131) while spacing stays non-negative. Commit on
   *  Enter/blur (not per keystroke): a controlled input resets the transient
   *  "-" while a negative number is being typed, eating the minus sign. */
  const ptInput = (attr: string, title: string, min = 0) => {
    const twips = Number(paraAttrs[attr]) || 0
    const shown = Math.round(twips * PT_PER_TWIP)
    // selection identity in the key: a commit that leaves the LIVE
    // selection's value unchanged must still remount the field, or it keeps
    // showing the number just applied to a different paragraph
    const selFrom = editor.state.selection.from
    return (
      <label className="layout-num" data-tip={title}>
        <span>{title}</span>
        <input
          type="number"
          min={min}
          max={400}
          step={1}
          disabled={!hasDoc}
          key={`${attr}:${shown}:${hasDoc}:${selFrom}`}
          defaultValue={shown}
          onKeyDown={(e) => {
            if (e.key === 'Enter') (e.target as HTMLInputElement).blur()
          }}
          // the paragraphs the entry is FOR: by blur, a click may already have
          // moved the live selection elsewhere
          onFocus={() => {
            const { from, to } = editor.state.selection
            ptInputTargetRef.current = { from, to }
          }}
          onBlur={(e) => {
            const target = ptInputTargetRef.current
            ptInputTargetRef.current = null
            const raw = e.target.value.trim()
            const pt = raw === '' ? 0 : Math.min(400, Math.max(min, Number(raw)))
            if (Number.isFinite(pt)) {
              // clamped input keeps the typed text (uncontrolled): sync it back
              if (String(pt) !== raw) e.target.value = String(pt)
              if (Math.round(pt / PT_PER_TWIP) !== twips) {
                setParaAttrs(
                  editor,
                  { [attr]: pt !== 0 ? Math.round(pt / PT_PER_TWIP) : null },
                  target ?? undefined,
                )
              }
            }
            // regardless of whether anything changed: Enter-blur (no
            // relatedTarget) hands focus back to the editor; a blur INTO
            // another control must not steal it back
            if (!e.relatedTarget) editor.commands.focus()
          }}
        />
        <span className="layout-unit">pt</span>
      </label>
    )
  }

  const pageSetupLabel =
    activeSection !== null
      ? t('ribbonGroupPageSetupSection', { n: activeSection + 1 })
      : t('ribbonGroupPageSetup')

  return (
    <>
      <RibbonFoldGroup
        label={pageSetupLabel}
        icon={<IconMargins size={BIG} />}
        caret={<IconCaret />}
        priority={3}
        caption
      >
        <div className="rb-split-wrap">
          <button
            className="rb-big"
            disabled={!enabled}
            data-tip={t('ribbonMargins')}
            onClick={() => toggleDropdown(setDropdown, 'margins')}
          >
            <span className="rb-big-icon">
              <IconMargins size={BIG} />
              <IconCaret />
            </span>
            <span>{t('ribbonMargins')}</span>
          </button>
          {dropdown === 'margins' && section && (
            <div data-rb-panel="" className="layout-menu">
              {(() => {
                const storedLastCustom = readLastCustomMargins()
                const lastCustom =
                  storedLastCustom &&
                  marginsFitPage(storedLastCustom, section.pageWidth, section.pageHeight)
                    ? storedLastCustom
                    : null
                return (
                  <>
                    {lastCustom && (
                      <button
                        className={marginsActive(lastCustom) ? 'active' : ''}
                        onClick={() => {
                          applyMargins(lastCustom)
                          setDropdown(() => null)
                        }}
                      >
                        <b>{t('ribbonMarginLastCustom')}</b>
                        <span>
                          {t('ribbonMarginTop')} {cmFromTwips(lastCustom.top)} ·{' '}
                          {t('ribbonMarginBottom')} {cmFromTwips(lastCustom.bottom)} ·{' '}
                          {t('ribbonMarginLeft')} {cmFromTwips(lastCustom.left)} ·{' '}
                          {t('ribbonMarginRight')} {cmFromTwips(lastCustom.right)} cm
                        </span>
                      </button>
                    )}
                    {MARGIN_PRESETS.map((m) => (
                      <button
                        key={m.key}
                        className={marginsActive(m) ? 'active' : ''}
                        disabled={!marginsFitPage(m, section.pageWidth, section.pageHeight)}
                        onClick={() => {
                          applyMargins(m)
                          setDropdown(() => null)
                        }}
                      >
                        <b>{t(m.nameKey)}</b>
                        <span>{t(m.descKey)}</span>
                      </button>
                    ))}
                    <button
                      className={marginsActive(MIRRORED_MARGINS, true) ? 'active' : ''}
                      disabled={
                        !marginsFitPage(MIRRORED_MARGINS, section.pageWidth, section.pageHeight)
                      }
                      onClick={() => {
                        applyMargins(MIRRORED_MARGINS, { mirror: true })
                        setDropdown(() => null)
                      }}
                    >
                      <b>{t('ribbonMarginMirrored')}</b>
                      <span>{t('ribbonMarginMirroredDesc')}</span>
                    </button>
                    <button
                      onClick={() => {
                        setDropdown(() => null)
                        setMarginDialog(true)
                      }}
                    >
                      <b>{t('ribbonMarginCustom')}</b>
                    </button>
                  </>
                )
              })()}
            </div>
          )}
        </div>
        <div className="rb-split-wrap">
          <button
            className="rb-big"
            disabled={!enabled}
            data-tip={t('ribbonOrientation')}
            onClick={() => toggleDropdown(setDropdown, 'orient')}
          >
            <span className="rb-big-icon">
              <IconOrientation size={BIG} />
              <IconCaret />
            </span>
            <span>{t('ribbonOrientation')}</span>
          </button>
          {dropdown === 'orient' && section && (
            <div data-rb-panel="" className="layout-menu">
              <button
                className={section.orientation === 'portrait' ? 'active' : ''}
                onClick={() => setOrientation('portrait')}
              >
                <b>{t('ribbonPortrait')}</b>
              </button>
              <button
                className={section.orientation === 'landscape' ? 'active' : ''}
                onClick={() => setOrientation('landscape')}
              >
                <b>{t('ribbonLandscape')}</b>
              </button>
            </div>
          )}
        </div>
        <div className="rb-split-wrap">
          <button
            className="rb-big"
            disabled={!enabled}
            data-tip={t('ribbonPaperSize')}
            onClick={() => toggleDropdown(setDropdown, 'paper')}
          >
            <span className="rb-big-icon">
              <IconPageSize size={BIG} />
              <IconCaret />
            </span>
            <span>{t('ribbonPaperSize')}</span>
          </button>
          {dropdown === 'paper' && section && (
            <div data-rb-panel="" className="layout-menu">
              {PAPER_SIZES.map((p) => {
                const portraitW = Math.min(section.pageWidth, section.pageHeight)
                return (
                  <button
                    key={p.key}
                    className={portraitW === p.w ? 'active' : ''}
                    onClick={() => setPaper(p.w, p.h)}
                  >
                    <b>{p.name}</b>
                    <span>{p.desc}</span>
                  </button>
                )
              })}
              <button
                onClick={() => {
                  setDropdown(() => null)
                  setPaperDialog(true)
                }}
              >
                <b>{t('ribbonMorePaperSizes')}</b>
              </button>
            </div>
          )}
        </div>
        <div className="rb-split-wrap">
          <button
            className={`rb-big ${section && section.columns > 1 ? 'active' : ''}`}
            disabled={!enabled}
            data-tip={t('ribbonColumns')}
            onClick={() => toggleDropdown(setDropdown, 'columns')}
          >
            <span className="rb-big-icon">
              <IconColumns size={BIG} />
              <IconCaret />
            </span>
            <span>{t('ribbonColumns')}</span>
          </button>
          {dropdown === 'columns' && section && (
            <div data-rb-panel="" className="layout-menu">
              {([1, 2, 3] as const).map((n) => (
                <button
                  key={n}
                  className={section.columns === n && !section.colWidths ? 'active' : ''}
                  onClick={() => {
                    onSection(withColumns(section, n))
                    setDropdown(() => null)
                  }}
                >
                  {n === 1
                    ? t('ribbonOneColumn')
                    : n === 2
                      ? t('ribbonTwoColumns')
                      : t('ribbonThreeColumns')}
                </button>
              ))}
              {(['left', 'right'] as const).map((kind) => {
                const w = section.colWidths
                const active =
                  section.columns === 2 &&
                  !!w &&
                  w.length === 2 &&
                  (kind === 'left' ? w[0] < w[1] : w[0] > w[1])
                return (
                  <button
                    key={kind}
                    className={active ? 'active' : ''}
                    onClick={() => {
                      const space = section.colSpace ?? 720
                      const textWidth = section.pageWidth - section.marginLeft - section.marginRight
                      onSection(
                        withColumns(section, 2, {
                          space,
                          widths: columnPresetWidths(kind, textWidth, space),
                        }),
                      )
                      setDropdown(() => null)
                    }}
                  >
                    {kind === 'left' ? t('ribbonColumnsLeft') : t('ribbonColumnsRight')}
                  </button>
                )
              })}
              <button
                onClick={() => {
                  setDropdown(() => null)
                  setColumnsDialog(true)
                }}
              >
                {t('ribbonMoreColumns')}
              </button>
            </div>
          )}
        </div>
        <div className="rb-split-wrap">
          <button
            className="rb-big"
            disabled={!enabled}
            data-tip={t('ribbonSectionBreakTip')}
            onClick={() => toggleDropdown(setDropdown, 'sectbreak')}
          >
            <span className="rb-big-icon">
              <IconBreaks size={BIG} />
              <IconCaret />
            </span>
            <span>{t('ribbonBreaks')}</span>
          </button>
          {dropdown === 'sectbreak' && (
            <div data-rb-panel="" className="layout-menu">
              <div className="layout-menu-head">{t('ribbonPageBreaksHeading')}</div>
              <button
                onClick={() => {
                  insertPageBreakAt(editor)
                  setDropdown(() => null)
                }}
              >
                <b>{t('ribbonBreakPage')}</b>
                <span>{t('ribbonBreakPageDesc')}</span>
              </button>
              <button
                onClick={() => {
                  editor
                    .chain()
                    .focus()
                    .insertContent({ type: 'hardBreak', attrs: { colBreak: true } })
                    .run()
                  setDropdown(() => null)
                }}
              >
                <b>{t('ribbonBreakColumn')}</b>
                <span>{t('ribbonBreakColumnDesc')}</span>
              </button>
              <button
                onClick={() => {
                  editor.chain().focus().insertContent({ type: 'hardBreak' }).run()
                  setDropdown(() => null)
                }}
              >
                <b>{t('ribbonBreakTextWrapping')}</b>
                <span>{t('ribbonBreakTextWrappingDesc')}</span>
              </button>
              <div className="layout-menu-head">{t('ribbonSectionBreak')}</div>
              {(
                [
                  ['nextPage', 'ribbonBreakNextPage', 'ribbonBreakNextPageDesc'],
                  ['continuous', 'ribbonBreakContinuous', 'ribbonBreakContinuousDesc'],
                  ['evenPage', 'ribbonBreakEvenPage', 'ribbonBreakEvenPageDesc'],
                  ['oddPage', 'ribbonBreakOddPage', 'ribbonBreakOddPageDesc'],
                ] as const
              ).map(([type, nameKey, descKey]) => (
                <button
                  key={type}
                  onClick={() => {
                    onInsertSectionBreak(type)
                    setDropdown(() => null)
                  }}
                >
                  <b>{t(nameKey)}</b>
                  <span>{t(descKey)}</span>
                </button>
              ))}
            </div>
          )}
        </div>
        <div className="rb-split-wrap">
          <button
            className={`rb-big ${section?.lineNumbers ? 'active' : ''}`}
            disabled={!enabled}
            data-tip={t('ribbonLineNumbersTip')}
            onClick={() => toggleDropdown(setDropdown, 'linenum')}
          >
            <span className="rb-big-icon">
              <IconLineNumbers size={BIG} />
              <IconCaret />
            </span>
            <span>{t('ribbonLineNumbers')}</span>
          </button>
          {dropdown === 'linenum' && section && (
            <div data-rb-panel="" className="layout-menu">
              {(
                [
                  ['none', 'ribbonLineNumNone'],
                  ['continuous', 'ribbonLineNumContinuous'],
                  ['newPage', 'ribbonLineNumRestartPage'],
                  ['newSection', 'ribbonLineNumRestartSection'],
                ] as const
              ).map(([mode, labelKey]) => {
                const ln = section.lineNumbers
                const active = mode === 'none' ? !ln : ln?.restart === mode
                return (
                  <button
                    key={mode}
                    className={active ? 'active' : ''}
                    onClick={() => {
                      onSection({
                        ...section,
                        lineNumbers:
                          mode === 'none'
                            ? null
                            : { ...(ln ?? DEFAULT_LINE_NUMBERING), restart: mode },
                      })
                      setDropdown(() => null)
                    }}
                  >
                    {t(labelKey)}
                  </button>
                )
              })}
              <button
                onClick={() => {
                  setDropdown(() => null)
                  setLineNumDialog(true)
                }}
              >
                {t('ribbonLineNumOptions')}
              </button>
            </div>
          )}
        </div>
        <div className="rb-split-wrap">
          <button
            className={`rb-big ${docFlags.autoHyphenation ? 'active' : ''}`}
            disabled={!hasDoc}
            data-tip={t('ribbonHyphenationTip')}
            onClick={() => toggleDropdown(setDropdown, 'hyphen')}
          >
            <span className="rb-big-icon">
              <IconHyphenation size={BIG} />
              <IconCaret />
            </span>
            <span>{t('ribbonHyphenation')}</span>
          </button>
          {dropdown === 'hyphen' && (
            <div data-rb-panel="" className="layout-menu">
              <button
                className={docFlags.autoHyphenation ? '' : 'active'}
                onClick={() => {
                  onDocFlags({ autoHyphenation: false })
                  setDropdown(() => null)
                }}
              >
                {t('ribbonHyphenNone')}
              </button>
              <button
                className={docFlags.autoHyphenation ? 'active' : ''}
                onClick={() => {
                  onDocFlags({ autoHyphenation: true })
                  setDropdown(() => null)
                }}
              >
                {t('ribbonHyphenAuto')}
              </button>
            </div>
          )}
        </div>
      </RibbonFoldGroup>

      <div className="ribbon-sep" />

      <RibbonFoldGroup
        label={t('ribbonGroupParagraphSelected')}
        icon={<IconParaSpacing size={BIG} />}
        caret={<IconCaret />}
        priority={2}
        itemsClassName="ribbon-group-items layout-para"
        caption
      >
        <div className="layout-col">
          {ptInput('indentLeft', t('ribbonIndentLeft'), -400)}
          {ptInput('indentRight', t('ribbonIndentRight'), -400)}
        </div>
        <div className="layout-col">
          {ptInput('spaceBefore', t('ribbonSpaceBefore'))}
          {ptInput('spaceAfter', t('ribbonSpaceAfter'))}
        </div>
      </RibbonFoldGroup>

      <div className="ribbon-sep" />

      <RibbonFoldGroup
        label={t('ribbonGroupArrange')}
        icon={<IconPosition size={BIG} />}
        caret={<IconCaret />}
        priority={1}
        caption
      >
        <div className="rb-split-wrap">
          <button
            className="rb-big"
            disabled={!canPosition}
            data-tip={t('ribbonPosition')}
            onClick={() => toggleDropdown(setDropdown, 'arrange-pos')}
          >
            <span className="rb-big-icon">
              <IconPosition size={BIG} />
              <IconCaret />
            </span>
            <span>{t('ribbonPosition')}</span>
          </button>
          {dropdown === 'arrange-pos' && (
            <div data-rb-panel="" className="layout-menu">
              <button className={currentWrap ? '' : 'active'} onClick={() => applyInlinePosition()}>
                <b>{t('appWrapInline')}</b>
              </button>
              <div className="pos-grid">
                {(['top', 'center', 'bottom'] as const).map((v) =>
                  (['left', 'center', 'right'] as const).map((h) => (
                    <button
                      key={`${v}-${h}`}
                      className={`pos-cell ph-${h} pv-${v}${
                        protAttrs?.imagePosH === h && protAttrs?.imagePosV === v ? ' active' : ''
                      }`}
                      data-tip={t('ribbonPosition')}
                      aria-label={t('ribbonPosition')}
                      onClick={() => applyPositionPreset(h, v)}
                    >
                      <span className="pos-dot" />
                    </button>
                  )),
                )}
              </div>
            </div>
          )}
        </div>
        <button
          className={`rb-big ${currentWrap === 'front' ? 'active' : ''}`}
          disabled={!canWrap}
          data-tip={t('appWrapFront')}
          onClick={() => applyWrap('front')}
        >
          <span className="rb-big-icon">
            <IconBringForward size={BIG} />
          </span>
          <span>{t('ribbonBringForward')}</span>
        </button>
        <button
          className={`rb-big ${currentWrap === 'behind' ? 'active' : ''}`}
          disabled={!canWrap}
          data-tip={t('appWrapBehind')}
          onClick={() => applyWrap('behind')}
        >
          <span className="rb-big-icon">
            <IconSendBackward size={BIG} />
          </span>
          <span>{t('ribbonSendBackward')}</span>
        </button>
        <div className="rb-split-wrap">
          <button
            className="rb-big"
            disabled={!canPosition}
            data-tip={t('ribbonObjAlign')}
            onClick={() => toggleDropdown(setDropdown, 'arrange-align')}
          >
            <span className="rb-big-icon">
              <IconAlignObjects size={BIG} />
              <IconCaret />
            </span>
            <span>{t('ribbonObjAlign')}</span>
          </button>
          {dropdown === 'arrange-align' && (
            <div data-rb-panel="" className="layout-menu">
              {(
                [
                  ['h', 'left', 'ribbonObjAlignLeft'],
                  ['h', 'center', 'ribbonObjAlignCenter'],
                  ['h', 'right', 'ribbonObjAlignRight'],
                  ['v', 'top', 'ribbonAlignTop'],
                  ['v', 'center', 'ribbonAlignMiddle'],
                  ['v', 'bottom', 'ribbonAlignBottom'],
                ] as const
              ).map(([axis, value, key]) => (
                <button key={`${axis}-${value}`} onClick={() => alignObject(axis, value as 'left')}>
                  {t(key)}
                </button>
              ))}
            </div>
          )}
        </div>
        <div className="rb-split-wrap">
          <button
            className="rb-big"
            disabled={!canWrap}
            data-tip={t('ribbonWrapText')}
            onClick={() => toggleDropdown(setDropdown, 'arrange-wrap')}
          >
            <span className="rb-big-icon">
              <IconWrapText size={BIG} />
              <IconCaret />
            </span>
            <span>{t('ribbonWrapText')}</span>
          </button>
          {dropdown === 'arrange-wrap' && (
            <div data-rb-panel="" className="layout-menu">
              {WRAP_OPTIONS.map((opt) => (
                <button
                  key={String(opt.value)}
                  className={currentWrap === opt.value ? 'active' : ''}
                  onClick={() => applyWrap(opt.value)}
                >
                  <b>{t(opt.labelKey)}</b>
                </button>
              ))}
            </div>
          )}
        </div>
      </RibbonFoldGroup>

      {marginDialog && section && (
        <MarginDialog
          margins={{
            top: section.marginTop,
            right: section.marginRight,
            bottom: section.marginBottom,
            left: section.marginLeft,
          }}
          gutter={section.gutterAtTop ? null : (section.gutter ?? 0)}
          mirror={docFlags.mirrorMargins}
          pageWidth={section.pageWidth}
          pageHeight={section.pageHeight}
          onApply={(m, extras) => {
            applyMargins(m, extras)
            if (extras.gutter === 0 && !extras.mirror)
              localStorage.setItem(LAST_CUSTOM_MARGINS_KEY, JSON.stringify(m))
          }}
          onClose={() => setMarginDialog(false)}
        />
      )}
      {columnsDialog && section && (
        <ColumnsDialog
          section={section}
          onApply={onSection}
          onClose={() => setColumnsDialog(false)}
        />
      )}
      {paperDialog && section && (
        <PaperSizeDialog
          section={section}
          onApply={onSection}
          onClose={() => setPaperDialog(false)}
        />
      )}
      {lineNumDialog && section && (
        <LineNumbersDialog
          current={section.lineNumbers}
          onApply={(lineNumbers) => onSection({ ...section, lineNumbers })}
          onClose={() => setLineNumDialog(false)}
        />
      )}
    </>
  )
}

/* ================= References ================= */
