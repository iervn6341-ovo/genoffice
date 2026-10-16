import { useState } from 'react'
import type { Editor } from '@tiptap/core'
import { closeHistory } from '@tiptap/pm/history'
import type { ZoteroCommand } from '../../shared/ipc'
import {
  bibliographyLine,
  citationText,
  generateCaptionXml,
  generateIndexFieldXml,
  generateTocFieldXml,
  DEFAULT_TOC_OPTIONS,
  tocOptionsFromField,
  tocInstruction,
  type Block,
  type SourceInfo,
  type StyleInfo,
  type TocEntry,
  type TocFormat,
  type TocLeader,
  type TocOptions,
} from '@genoffice/docx-engine'
import { Dropdown, RibbonFoldGroup } from '@genoffice/ui'
import { PromptModal } from './PromptModal'
import {
  anchorTocReplacement,
  collectTocSources,
  findTocField,
  preserveTocHistorySource,
  tocNodeXml,
  updateTocPageCaches,
} from '../editor/toc-actions'
import { t, useI18n, type StringKey } from '../i18n/locale'
import {
  IconBook,
  IconCaption,
  IconCaret,
  IconCitation,
  IconEndnote,
  IconFootnote,
  IconIndex,
  IconRefresh,
  IconToc,
} from './icons'

/** icon size for the big icon-over-label ribbon buttons (slides ribbon parity) */
import { BIG, TabProps, toggleDropdown } from './ribbon-tabs'

/** TOC sources + real page numbers in the same document order. */
function collectTocEntriesWithPages(
  editor: Editor,
  blocks: Block[],
  styles: ReadonlyMap<string, StyleInfo> | undefined,
  options: TocOptions,
  headingPages?: (positions?: number[]) => string[] | null,
): TocEntry[] {
  const entries = collectTocSources(editor, blocks, styles, options)
  const pages = headingPages?.(entries.map((entry) => entry.pos))
  if (pages && pages.length === entries.length) {
    return entries.map(({ level, text }, i) => ({ level, text, pageNo: pages[i] }))
  }
  return entries.map(({ level, text }) => ({ level, text }))
}

/** PM nodes for a freshly generated TOC field (one docProtected per line) */
function tocFieldNodes(
  entries: TocEntry[],
  options?: TocOptions,
  instruction?: string,
): Array<Record<string, unknown>> {
  return generateTocFieldXml(entries, options, instruction).map((xml, i) => ({
    type: 'docProtected',
    attrs: {
      docxIndex: null,
      blockType: 'passthrough',
      label: t('ribbonTocFieldLabel'),
      genXml: xml,
      fieldDisplay: {
        kind: 'tocLine',
        left: entries[i].text,
        right:
          options?.showPageNumbers === false
            ? ''
            : entries[i].pageNo !== undefined
              ? String(entries[i].pageNo)
              : '',
        level: entries[i].level,
        leader: options?.leader ?? 'dot',
      },
    },
  }))
}

const PAGE_BREAK_PARAGRAPH_XML = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'

const TOC_FORMATS: TocFormat[] = [
  'template',
  'classic',
  'distinctive',
  'fancy',
  'modern',
  'formal',
  'simple',
]
const TOC_LEADERS: TocLeader[] = ['none', 'dot', 'hyphen', 'underscore', 'heavy', 'middleDot']

function TocOptionsModal({
  initial,
  styles,
  onApply,
  onClose,
}: {
  initial: TocOptions
  styles?: ReadonlyMap<string, StyleInfo>
  onApply: (options: TocOptions) => void
  onClose: () => void
}) {
  const { t } = useI18n()
  const [options, setOptions] = useState<TocOptions>(() => ({
    ...initial,
    styleLevels: { ...initial.styleLevels },
  }))
  const [styleSearch, setStyleSearch] = useState('')
  const update = (patch: Partial<TocOptions>) =>
    setOptions((previous) => ({ ...previous, ...patch }))
  const styleChoices = [...(styles?.values() ?? [])]
    .filter((style) => style.type === 'paragraph' && !/^TOC[1-9]$/.test(style.styleId))
    .filter((style) => style.name.toLocaleLowerCase().includes(styleSearch.toLocaleLowerCase()))
    .sort((a, b) => a.name.localeCompare(b.name))

  return (
    <div
      className="modal-backdrop"
      onMouseDown={(event) => event.target === event.currentTarget && onClose()}
    >
      <div
        className="modal toc-options-modal"
        role="dialog"
        aria-modal="true"
        aria-label={t('ribbonTocCustom')}
      >
        <h2>{t('ribbonTocCustom')}</h2>
        <div className="toc-options-grid">
          <label>
            {t('ribbonTocFormat')}
            <select
              value={options.format}
              onChange={(event) => update({ format: event.target.value as TocFormat })}
            >
              {TOC_FORMATS.map((format) => (
                <option key={format} value={format}>
                  {t(`ribbonTocFormat_${format}` as StringKey)}
                </option>
              ))}
            </select>
          </label>
          <label>
            {t('ribbonTocLevels')}
            <input
              type="number"
              min="1"
              max="9"
              value={options.levels}
              onChange={(event) =>
                update({ levels: Math.min(9, Math.max(1, Number(event.target.value) || 1)) })
              }
            />
          </label>
          <label>
            {t('ribbonTocLeader')}
            <select
              value={options.leader}
              disabled={!options.showPageNumbers || !options.rightAlignPageNumbers}
              onChange={(event) => update({ leader: event.target.value as TocLeader })}
            >
              {TOC_LEADERS.map((leader) => (
                <option key={leader} value={leader}>
                  {t(`ribbonTocLeader_${leader}` as StringKey)}
                </option>
              ))}
            </select>
          </label>
        </div>
        <div className="toc-options-checks">
          {(
            [
              ['showPageNumbers', 'ribbonTocShowPages'],
              ['rightAlignPageNumbers', 'ribbonTocRightAlign'],
              ['useHyperlinks', 'ribbonTocHyperlinks'],
              ['useHeadingStyles', 'ribbonTocHeadingStyles'],
              ['useOutlineLevels', 'ribbonTocOutlineLevels'],
              ['useTcFields', 'ribbonTocTcFields'],
            ] as const
          ).map(([name, label]) => (
            <label key={name} className="toc-option-check">
              <input
                type="checkbox"
                checked={options[name]}
                disabled={name === 'rightAlignPageNumbers' && !options.showPageNumbers}
                onChange={(event) => update({ [name]: event.target.checked })}
              />
              <span>{t(label)}</span>
            </label>
          ))}
        </div>
        <fieldset className="toc-style-levels">
          <legend>{t('ribbonTocStyles')}</legend>
          <input
            value={styleSearch}
            onChange={(event) => setStyleSearch(event.target.value)}
            aria-label={t('ribbonTocSearchStyles')}
            placeholder={t('ribbonTocSearchStyles')}
          />
          <div className="toc-style-list">
            {styleChoices.map((style) => (
              <label key={style.styleId}>
                <span>{style.name}</span>
                <select
                  aria-label={`${style.name} ${t('ribbonTocLevel')}`}
                  disabled={/[",]/.test(style.name)}
                  title={/[",]/.test(style.name) ? t('ribbonTocUnsupported') : undefined}
                  value={options.styleLevels[style.name] ?? 0}
                  onChange={(event) => {
                    const styleLevels = { ...options.styleLevels }
                    const level = Number(event.target.value)
                    if (level === 0) delete styleLevels[style.name]
                    else styleLevels[style.name] = level
                    update({ styleLevels })
                  }}
                >
                  <option value="0">—</option>
                  {Array.from({ length: 9 }, (_, index) => (
                    <option key={index + 1} value={index + 1}>
                      {index + 1}
                    </option>
                  ))}
                </select>
              </label>
            ))}
          </div>
        </fieldset>
        <div className="modal-actions">
          <button className="btn-ghost" onClick={onClose}>
            {t('ribbonCancel')}
          </button>
          <button className="btn-primary" onClick={() => onApply(options)}>
            {t('ribbonInsert')}
          </button>
        </div>
      </div>
    </div>
  )
}

function TocUpdateModal({
  onApply,
  onClose,
}: {
  onApply: (mode: 'pages' | 'full') => void
  onClose: () => void
}) {
  const { t } = useI18n()
  const [mode, setMode] = useState<'pages' | 'full'>('pages')
  return (
    <div
      className="modal-backdrop"
      onMouseDown={(event) => event.target === event.currentTarget && onClose()}
    >
      <div
        className="modal toc-update-modal"
        role="dialog"
        aria-modal="true"
        aria-label={t('ribbonTocUpdate')}
      >
        <h2>{t('ribbonTocUpdate')}</h2>
        <label className="toc-option-check">
          <input
            type="radio"
            name="toc-update-mode"
            checked={mode === 'pages'}
            onChange={() => setMode('pages')}
          />
          <span>{t('ribbonTocUpdatePages')}</span>
        </label>
        <label className="toc-option-check">
          <input
            type="radio"
            name="toc-update-mode"
            checked={mode === 'full'}
            onChange={() => setMode('full')}
          />
          <span>{t('ribbonTocUpdateAll')}</span>
        </label>
        <div className="modal-actions">
          <button className="btn-ghost" onClick={onClose}>
            {t('ribbonCancel')}
          </button>
          <button className="btn-primary" onClick={() => onApply(mode)}>
            {t('ribbonTocUpdate')}
          </button>
        </div>
      </div>
    </div>
  )
}

/** OOXML of a top-level PM node: editor-generated fragment or original slice */
const xmlOfNode = tocNodeXml

/**
 * Find the TOC field region (field begin ... matching end, tracked by fldChar
 * depth across top-level blocks) and replace it with a regenerated dirty TOC
 * field built from the current headings.
 */
function updateTocField(
  editor: Editor,
  blocks: Block[],
  styles: ReadonlyMap<string, StyleInfo> | undefined,
  headingPages: ((positions?: number[]) => string[] | null) | undefined,
  mode: 'pages' | 'full',
): void {
  const region = findTocField(editor, blocks)
  if (!region) {
    window.alert(t('ribbonTocNotFound'))
    return
  }
  const options = tocOptionsFromField(region.instruction, region.nodes[0]?.xml ?? '')
  if (!options) {
    window.alert(t('ribbonTocUnsupported'))
    return
  }
  const entries = collectTocEntriesWithPages(editor, blocks, styles, options, headingPages)
  if (entries.length === 0) {
    window.alert(t('ribbonTocNoHeadings'))
    return
  }
  if (mode === 'pages') {
    const lines = region.nodes.filter(({ node }) => node.attrs.fieldDisplay?.kind === 'tocLine')
    const pages = entries.map((entry) => String(entry.pageNo ?? ''))
    if (
      lines.length !== pages.length ||
      pages.some((page) => !page) ||
      !updateTocPageCaches(editor, region, pages, blocks)
    ) {
      window.alert(t('ribbonTocPageUpdateUnavailable'))
    }
    return
  }
  const nodes = tocFieldNodes(entries, options, region.instruction)
  if (!anchorTocReplacement(nodes, region, blocks)) {
    window.alert(t('ribbonTocUnsupported'))
    return
  }
  if (region.keepPageBreak) {
    nodes.push({
      type: 'docProtected',
      attrs: {
        docxIndex: null,
        blockType: 'passthrough',
        label: t('ribbonPageBreak'),
        genXml: PAGE_BREAK_PARAGRAPH_XML,
        fieldDisplay: { kind: 'pageBreak' },
      },
    })
  }

  preserveTocHistorySource(editor, region, blocks)
  editor.view.dispatch(closeHistory(editor.state.tr))
  editor
    .chain()
    .focus()
    .deleteRange({ from: region.from, to: region.to })
    .insertContentAt(region.from, nodes as never)
    .run()
  editor.view.dispatch(closeHistory(editor.state.tr))
}

const CAPTION_LABEL_KEYS = [
  'ribbonCaptionFigure',
  'ribbonCaptionTable',
  'ribbonCaptionEquation',
] as const

/** Caption dialog: label + optional description, numbered by existing SEQ fields */
function CaptionModal({
  editor,
  blocks,
  onClose,
}: {
  editor: Editor
  blocks: Block[]
  onClose: () => void
}) {
  const { t } = useI18n()
  const CAPTION_LABELS = CAPTION_LABEL_KEYS.map((k) => t(k))
  const [label, setLabel] = useState<string>(() => t(CAPTION_LABEL_KEYS[0]))
  const [text, setText] = useState('')

  const nextNumber = (lbl: string): number => {
    let count = 0
    const re = new RegExp(`SEQ\\s+${lbl}[\\s\\\\]`)
    editor.state.doc.forEach((node) => {
      if (node.type.name !== 'docProtected') return
      if (re.test(xmlOfNode(node as never, blocks))) count += 1
    })
    return count + 1
  }

  const insert = () => {
    const number = nextNumber(label)
    const xml = generateCaptionXml(label, number, text.trim())
    const display = `${label} ${number}${text.trim() ? ` ${text.trim()}` : ''}`
    editor
      .chain()
      .focus()
      .insertContent({
        type: 'docProtected',
        attrs: {
          docxIndex: null,
          blockType: 'passthrough',
          label: t('ribbonCaption'),
          genXml: xml,
          fieldDisplay: { kind: 'text', left: display },
        },
      } as never)
      .run()
    onClose()
  }

  return (
    <div className="modal-backdrop" onMouseDown={(e) => e.target === e.currentTarget && onClose()}>
      <div className="modal">
        <h2>{t('ribbonCaptionInsertTitle')}</h2>
        <label>
          {t('ribbonCaptionLabel')}
          <Dropdown
            value={label}
            ariaLabel={t('ribbonCaptionLabel')}
            options={CAPTION_LABELS.map((l) => ({ value: l, label: l }))}
            onPick={setLabel}
          />
        </label>
        <label>
          {t('ribbonCaption')}
          <input
            value={text}
            onChange={(e) => setText(e.target.value)}
            placeholder={t('ribbonCaptionPh', { label, n: nextNumber(label) })}
            onKeyDown={(e) => e.key === 'Enter' && insert()}
          />
        </label>
        <div className="modal-actions">
          <button className="btn-ghost" onClick={onClose}>
            {t('ribbonCancel')}
          </button>
          <button className="btn-primary" onClick={insert}>
            {t('ribbonInsert')}
          </button>
        </div>
      </div>
    </div>
  )
}

const SOURCE_TYPES: Array<{ key: string; nameKey: StringKey }> = [
  { key: 'Book', nameKey: 'ribbonSourceBook' },
  { key: 'JournalArticle', nameKey: 'ribbonSourceJournal' },
  { key: 'InternetSite', nameKey: 'ribbonSourceWebsite' },
  { key: 'Report', nameKey: 'ribbonSourceReport' },
  { key: 'Misc', nameKey: 'ribbonSourceMisc' },
]

/** Create Source dialog for the citation manager */
function SourceModal({
  sources,
  onAdd,
  onClose,
}: {
  sources: SourceInfo[]
  onAdd: (source: SourceInfo) => void
  onClose: () => void
}) {
  const { t } = useI18n()
  const [type, setType] = useState('Book')
  const [author, setAuthor] = useState('')
  const [title, setTitle] = useState('')
  const [year, setYear] = useState('')
  const [publisher, setPublisher] = useState('')
  const [url, setUrl] = useState('')

  const add = () => {
    if (!title.trim()) return
    // unique tag derived from author/title initials + year
    const base = (author.trim() || title.trim()).replace(/[\s,，、]+/g, '').slice(0, 8) || 'Src'
    let tag = `${base}${year.trim()}`
    let n = 1
    while (sources.some((s) => s.tag === tag)) tag = `${base}${year.trim()}_${++n}`
    onAdd({
      tag,
      type,
      author: author.trim(),
      title: title.trim(),
      year: year.trim(),
      ...(publisher.trim() ? { publisher: publisher.trim() } : {}),
      ...(url.trim() ? { url: url.trim() } : {}),
    })
    onClose()
  }

  return (
    <div className="modal-backdrop" onMouseDown={(e) => e.target === e.currentTarget && onClose()}>
      <div className="modal">
        <h2>{t('ribbonSourceCreateTitle')}</h2>
        <label>
          {t('ribbonSourceType')}
          <Dropdown
            value={type}
            ariaLabel={t('ribbonSourceType')}
            options={SOURCE_TYPES.map((s) => ({ value: s.key, label: t(s.nameKey) }))}
            onPick={setType}
          />
        </label>
        <label>
          {t('ribbonSourceAuthor')}
          <input
            value={author}
            onChange={(e) => setAuthor(e.target.value)}
            placeholder={t('ribbonSourceAuthorPh')}
          />
        </label>
        <label>
          {t('ribbonSourceTitle')}
          <input
            value={title}
            onChange={(e) => setTitle(e.target.value)}
            placeholder={t('ribbonSourceTitlePh')}
          />
        </label>
        <label>
          {t('ribbonSourceYear')}
          <input value={year} onChange={(e) => setYear(e.target.value)} placeholder="2026" />
        </label>
        <label>
          {type === 'JournalArticle'
            ? t('ribbonSourceJournalName')
            : type === 'InternetSite'
              ? t('ribbonSourceSiteName')
              : t('ribbonSourcePublisher')}
          <input value={publisher} onChange={(e) => setPublisher(e.target.value)} />
        </label>
        {type === 'InternetSite' && (
          <label>
            URL
            <input value={url} onChange={(e) => setUrl(e.target.value)} placeholder="https://…" />
          </label>
        )}
        <div className="modal-actions">
          <button className="btn-ghost" onClick={onClose}>
            {t('ribbonCancel')}
          </button>
          <button className="btn-primary" disabled={!title.trim()} onClick={add}>
            {t('ribbonAdd')}
          </button>
        </div>
      </div>
    </div>
  )
}

interface ReferencesTabProps extends TabProps {
  blocks: Block[]
  styles?: ReadonlyMap<string, StyleInfo>
  onInsertNote: (kind: 'footnote' | 'endnote') => void
  onNavigateNextNote: (kind: 'footnote' | 'endnote') => void
  onNoteSettings: () => void
  sources: SourceInfo[]
  /** footnotes/endnotes hold Zotero citation fields the bridge cannot see yet */
  zoteroNoteFields?: boolean
  onAddSource: (source: SourceInfo) => void
  /** TOC page-number backfill: docHeadings in document order → real page numbers */
  headingPages?: (positions?: number[]) => string[] | null
}

export function ReferencesTab({
  editor,
  hasDoc,
  blocks,
  styles,
  dropdown,
  setDropdown,
  onInsertNote,
  onNavigateNextNote,
  onNoteSettings,
  sources,
  zoteroNoteFields,
  onAddSource,
  headingPages,
}: ReferencesTabProps) {
  const { t } = useI18n()
  const [captionOpen, setCaptionOpen] = useState(false)
  const [tocOptionsOpen, setTocOptionsOpen] = useState(false)
  const [tocUpdateOpen, setTocUpdateOpen] = useState(false)
  const [sourceOpen, setSourceOpen] = useState(false)
  const [zoteroBusy, setZoteroBusy] = useState<ZoteroCommand | null>(null)

  const runZotero = async (command: ZoteroCommand) => {
    if (zoteroBusy) return
    setDropdown(() => null)
    // every command lets Zotero rebuild the bibliography from the fields it can see;
    // note citations are not among them yet, so their works would silently drop out
    if (zoteroNoteFields) {
      window.alert(t('zoteroNoteFieldsUnsupported'))
      return
    }
    setZoteroBusy(command)
    try {
      const result = await window.desktop.zoteroCommand(command)
      if (!result.ok) {
        console.error('Zotero integration failed', result.error)
        window.alert(
          t(
            result.errorCode === 'connection-refused'
              ? 'zoteroConnectionError'
              : 'zoteroOperationError',
          ),
        )
      }
    } catch (error) {
      console.error('Zotero integration failed', error)
      window.alert(t('zoteroOperationError'))
    } finally {
      setZoteroBusy(null)
    }
  }

  const insertToc = () => {
    const entries = collectTocEntriesWithPages(
      editor,
      blocks,
      styles,
      DEFAULT_TOC_OPTIONS,
      headingPages,
    )
    if (entries.length === 0) {
      window.alert(t('ribbonTocNoHeadings'))
      return
    }
    editor.view.dispatch(closeHistory(editor.state.tr))
    editor
      .chain()
      .focus()
      .insertContent(tocFieldNodes(entries, DEFAULT_TOC_OPTIONS) as never)
      .run()
    editor.view.dispatch(closeHistory(editor.state.tr))
  }

  const applyCustomToc = (options: TocOptions) => {
    const entries = collectTocEntriesWithPages(editor, blocks, styles, options, headingPages)
    if (entries.length === 0) {
      window.alert(t('ribbonTocNoHeadings'))
      return
    }
    const region = findTocField(editor, blocks)
    const nodes = tocFieldNodes(entries, options, tocInstruction(options))
    if (region) {
      if (!tocOptionsFromField(region.instruction, region.nodes[0]?.xml ?? '')) {
        window.alert(t('ribbonTocUnsupported'))
        return
      }
      if (!anchorTocReplacement(nodes, region, blocks)) {
        window.alert(t('ribbonTocUnsupported'))
        return
      }
      if (region.keepPageBreak)
        nodes.push({
          type: 'docProtected',
          attrs: {
            docxIndex: null,
            blockType: 'passthrough',
            label: t('ribbonPageBreak'),
            genXml: PAGE_BREAK_PARAGRAPH_XML,
            fieldDisplay: { kind: 'pageBreak' },
          },
        })
      editor.view.dispatch(closeHistory(editor.state.tr))
      editor
        .chain()
        .focus()
        .deleteRange({ from: region.from, to: region.to })
        .insertContentAt(region.from, nodes as never)
        .run()
      editor.view.dispatch(closeHistory(editor.state.tr))
    } else {
      editor.view.dispatch(closeHistory(editor.state.tr))
      editor
        .chain()
        .focus()
        .insertContent(nodes as never)
        .run()
      editor.view.dispatch(closeHistory(editor.state.tr))
    }
    setTocOptionsOpen(false)
  }

  const insertCitation = (source: SourceInfo) => {
    editor
      .chain()
      .focus()
      .insertContent({ type: 'text', text: citationText(source) } as never)
      .run()
    setDropdown(() => null)
  }

  const insertBibliography = () => {
    if (sources.length === 0) {
      window.alert(t('ribbonNoSources'))
      return
    }
    const nodes: Array<Record<string, unknown>> = [
      {
        type: 'docHeading',
        attrs: { docxIndex: null, styleId: null, aiChanged: false, level: 1 },
        content: [{ type: 'text', text: t('ribbonBibliographyHeading') }],
      },
      ...sources.map((s) => ({
        type: 'docParagraph',
        attrs: { docxIndex: null, styleId: null, aiChanged: false },
        content: [{ type: 'text', text: bibliographyLine(s) }],
      })),
    ]
    editor
      .chain()
      .focus()
      .insertContent(nodes as never)
      .run()
    setDropdown(() => null)
  }

  const [xePrompt, setXePrompt] = useState<{ initial: string; at: number } | null>(null)

  const markIndexEntry = () => {
    const { from, to } = editor.state.selection
    setXePrompt({ initial: editor.state.doc.textBetween(from, to, ' ').trim(), at: to })
  }

  const submitIndexEntry = (term: string) => {
    if (!xePrompt) return
    editor
      .chain()
      .focus()
      .insertContentAt(Math.min(xePrompt.at, editor.state.doc.content.size), {
        type: 'docXeMark',
        attrs: { term },
      } as never)
      .run()
  }

  const insertIndex = () => {
    const terms: string[] = []
    editor.state.doc.descendants((node) => {
      if (node.type.name === 'docXeMark') terms.push(String(node.attrs.term))
    })
    if (terms.length === 0) {
      window.alert(t('ribbonNoIndexEntries'))
      return
    }
    const unique = [...new Set(terms.map((t) => t.trim()).filter(Boolean))].sort((a, b) =>
      a.localeCompare(b, 'zh-CN'),
    )
    const nodes = generateIndexFieldXml(unique).map((xml, i) => ({
      type: 'docProtected',
      attrs: {
        docxIndex: null,
        blockType: 'passthrough',
        label: t('ribbonIndexFieldLabel'),
        genXml: xml,
        fieldDisplay: { kind: 'tocLine', left: unique[i], right: '', level: 1 },
      },
    }))
    editor
      .chain()
      .focus()
      .insertContent(nodes as never)
      .run()
  }

  return (
    <>
      <RibbonFoldGroup
        label={t('zoteroGroup')}
        icon={<IconBook size={BIG} />}
        caret={<IconCaret />}
        priority={1}
        caption
      >
        <button
          className="rb-big"
          disabled={!hasDoc || zoteroBusy !== null}
          data-tip={t('zoteroCitationTip')}
          onClick={() => void runZotero('addEditCitation')}
        >
          <span className="rb-big-icon">
            <IconCitation size={BIG} />
          </span>
          <span>{t('zoteroCitation')}</span>
        </button>
        <button
          className="rb-big"
          disabled={!hasDoc || zoteroBusy !== null}
          data-tip={t('zoteroBibliographyTip')}
          onClick={() => void runZotero('addEditBibliography')}
        >
          <span className="rb-big-icon">
            <IconBook size={BIG} />
          </span>
          <span>{t('zoteroBibliography')}</span>
        </button>
        <button
          className="rb-big"
          disabled={!hasDoc || zoteroBusy !== null}
          data-tip={t('zoteroRefreshTip')}
          onClick={() => void runZotero('refresh')}
        >
          <span className="rb-big-icon">
            <IconRefresh size={BIG} />
          </span>
          <span>{t('zoteroRefresh')}</span>
        </button>
        <div className="rb-split-wrap">
          <button
            className="rb-big"
            disabled={!hasDoc || zoteroBusy !== null}
            data-tip={t('zoteroDocumentSettingsTip')}
            onClick={() => toggleDropdown(setDropdown, 'zotero-settings')}
          >
            <span className="rb-big-icon">
              <IconCitation size={BIG} />
              <IconCaret />
            </span>
            <span>{t('zoteroDocumentSettings')}</span>
          </button>
          {dropdown === 'zotero-settings' && (
            <div data-rb-panel="" className="layout-menu">
              <button onClick={() => void runZotero('setDocPrefs')}>
                {t('zoteroDocumentPreferences')}
              </button>
              <button onClick={() => void runZotero('removeCodes')}>
                {t('zoteroRemoveCodes')}
              </button>
            </div>
          )}
        </div>
      </RibbonFoldGroup>

      <div className="ribbon-sep" />

      <RibbonFoldGroup
        label={t('ribbonToc')}
        icon={<IconToc size={BIG} />}
        caret={<IconCaret />}
        priority={4}
        caption
      >
        <button
          className="rb-big"
          disabled={!hasDoc}
          data-tip={t('ribbonTocTip')}
          onClick={insertToc}
        >
          <span className="rb-big-icon">
            <IconToc size={BIG} />
          </span>
          <span>{t('ribbonToc')}</span>
        </button>
        <button
          className="rb-big"
          disabled={!hasDoc}
          data-tip={t('ribbonTocCustom')}
          onClick={() => {
            const region = findTocField(editor, blocks)
            if (region && !tocOptionsFromField(region.instruction, region.nodes[0]?.xml ?? '')) {
              window.alert(t('ribbonTocUnsupported'))
              return
            }
            setTocOptionsOpen(true)
          }}
        >
          <span className="rb-big-icon">
            <IconToc size={BIG} />
          </span>
          <span>{t('ribbonTocCustom')}</span>
        </button>
        <button
          className="rb-big"
          disabled={!hasDoc}
          data-tip={t('ribbonTocUpdateTip')}
          onClick={() => setTocUpdateOpen(true)}
        >
          <span className="rb-big-icon">
            <IconRefresh size={BIG} />
          </span>
          <span>{t('ribbonTocUpdate')}</span>
        </button>
      </RibbonFoldGroup>

      {tocOptionsOpen && (
        <TocOptionsModal
          initial={(() => {
            const region = findTocField(editor, blocks)
            return region
              ? (tocOptionsFromField(region.instruction, region.nodes[0]?.xml ?? '') ??
                  DEFAULT_TOC_OPTIONS)
              : DEFAULT_TOC_OPTIONS
          })()}
          styles={styles}
          onApply={applyCustomToc}
          onClose={() => setTocOptionsOpen(false)}
        />
      )}
      {tocUpdateOpen && (
        <TocUpdateModal
          onApply={(mode) => {
            updateTocField(editor, blocks, styles, headingPages, mode)
            setTocUpdateOpen(false)
          }}
          onClose={() => setTocUpdateOpen(false)}
        />
      )}

      <div className="ribbon-sep" />

      <RibbonFoldGroup
        label={t('ribbonGroupFootnotes')}
        icon={<IconFootnote size={BIG} />}
        caret={<IconCaret />}
        priority={3}
        caption
      >
        <button
          className="rb-big"
          disabled={!hasDoc}
          data-tip={t('ribbonFootnoteTip')}
          onClick={() => onInsertNote('footnote')}
        >
          <span className="rb-big-icon">
            <IconFootnote size={BIG} />
          </span>
          <span>{t('ribbonFootnote')}</span>
        </button>
        <button
          className="rb-big"
          disabled={!hasDoc}
          data-tip={t('ribbonEndnoteTip')}
          onClick={() => onInsertNote('endnote')}
        >
          <span className="rb-big-icon">
            <IconEndnote size={BIG} />
          </span>
          <span>{t('ribbonEndnote')}</span>
        </button>
        <button className="rb-big" disabled={!hasDoc} onClick={onNoteSettings}>
          <span>{t('ribbonNoteSettings')}</span>
        </button>
        <button
          className="rb-big"
          disabled={!hasDoc}
          onClick={() => onNavigateNextNote('footnote')}
        >
          <span>{t('ribbonNextFootnote')}</span>
        </button>
        <button className="rb-big" disabled={!hasDoc} onClick={() => onNavigateNextNote('endnote')}>
          <span>{t('ribbonNextEndnote')}</span>
        </button>
      </RibbonFoldGroup>

      <div className="ribbon-sep" />

      <RibbonFoldGroup
        label={t('ribbonGroupCitationsIndex')}
        icon={<IconCitation size={BIG} />}
        caret={<IconCaret />}
        priority={2}
        caption
      >
        <div className="rb-split-wrap">
          <button
            className="rb-big"
            disabled={!hasDoc}
            data-tip={t('ribbonCitationTip')}
            onClick={() => toggleDropdown(setDropdown, 'citation')}
          >
            <span className="rb-big-icon">
              <IconCitation size={BIG} />
              <IconCaret />
            </span>
            <span>{t('ribbonCitation')}</span>
          </button>
          {dropdown === 'citation' && (
            <div data-rb-panel="" className="layout-menu">
              {sources.map((s) => (
                <button key={s.tag} data-tip={s.title} onClick={() => insertCitation(s)}>
                  {citationText(s)} {s.title.slice(0, 12)}
                </button>
              ))}
              <button
                onClick={() => {
                  setSourceOpen(true)
                  setDropdown(() => null)
                }}
              >
                {t('ribbonAddNewSource')}
              </button>
            </div>
          )}
        </div>
        <button
          className="rb-big"
          disabled={!hasDoc}
          data-tip={t('ribbonBibliographyTip')}
          onClick={insertBibliography}
        >
          <span className="rb-big-icon">
            <IconBook size={BIG} />
          </span>
          <span>{t('ribbonBibliography')}</span>
        </button>
        <button
          className="rb-big"
          disabled={!hasDoc}
          data-tip={t('ribbonCaptionTip')}
          onClick={() => setCaptionOpen(true)}
        >
          <span className="rb-big-icon">
            <IconCaption size={BIG} />
          </span>
          <span>{t('ribbonCaption')}</span>
        </button>
        <div className="rb-split-wrap">
          <button
            className="rb-big"
            disabled={!hasDoc}
            data-tip={t('ribbonIndexTip')}
            onClick={() => toggleDropdown(setDropdown, 'index')}
          >
            <span className="rb-big-icon">
              <IconIndex size={BIG} />
              <IconCaret />
            </span>
            <span>{t('ribbonIndex')}</span>
          </button>
          {dropdown === 'index' && (
            <div data-rb-panel="" className="layout-menu">
              <button
                onClick={() => {
                  markIndexEntry()
                  setDropdown(() => null)
                }}
              >
                {t('ribbonMarkEntry')}
              </button>
              <button
                onClick={() => {
                  insertIndex()
                  setDropdown(() => null)
                }}
              >
                {t('ribbonInsertIndex')}
              </button>
            </div>
          )}
        </div>
      </RibbonFoldGroup>

      {captionOpen && (
        <CaptionModal editor={editor} blocks={blocks} onClose={() => setCaptionOpen(false)} />
      )}
      {sourceOpen && (
        <SourceModal sources={sources} onAdd={onAddSource} onClose={() => setSourceOpen(false)} />
      )}
      {xePrompt && (
        <PromptModal
          title={t('ribbonMarkEntryTitle')}
          placeholder={t('ribbonMarkEntryPh')}
          initial={xePrompt.initial}
          onSubmit={submitIndexEntry}
          onClose={() => setXePrompt(null)}
        />
      )}
    </>
  )
}

/* ================= Review ================= */
