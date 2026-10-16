import { useState } from 'react'
import type { NoteProps } from '@genoffice/docx-engine'
import { useI18n } from '../i18n/locale'
import {
  effectiveNoteSettings,
  sectionNoteSettings,
  type NoteKind,
  type NoteSettingsChange,
  type NoteSettingsState,
} from '../editor/note-settings'
import { useModalKeys } from './modal-keys'

const FORMATS = [
  ['decimal', '1, 2, 3'],
  ['lowerRoman', 'i, ii, iii'],
  ['upperRoman', 'I, II, III'],
  ['lowerLetter', 'a, b, c'],
  ['upperLetter', 'A, B, C'],
  ['chicago', '*, †, ‡'],
] as const

type Draft = Record<NoteKind, NoteProps>
type Touched = Record<NoteKind, NoteSettingsChange>

export function NoteSettingsDialog({
  value,
  sectionIndex,
  onApply,
  onClose,
}: {
  value: NoteSettingsState
  sectionIndex: number | null
  onApply: (
    changes: Partial<Record<NoteKind, NoteSettingsChange>>,
    scope: 'document' | 'section',
  ) => void
  onClose: () => void
}) {
  const { t } = useI18n()
  const [kind, setKind] = useState<NoteKind>('footnote')
  const [scope, setScope] = useState<'document' | 'section'>('document')
  const [draft, setDraft] = useState<Draft>({
    footnote: effectiveNoteSettings(value, 'footnote'),
    endnote: effectiveNoteSettings(value, 'endnote'),
  })
  const [touched, setTouched] = useState<Touched>({ footnote: {}, endnote: {} })
  const [startText, setStartText] = useState({
    footnote: String(effectiveNoteSettings(value, 'footnote').numStart ?? 1),
    endnote: String(effectiveNoteSettings(value, 'endnote').numStart ?? 1),
  })
  const modalKeys = useModalKeys(onClose)
  const current = draft[kind]
  const defaults =
    kind === 'footnote'
      ? { pos: 'pageBottom', numFmt: 'decimal' }
      : { pos: 'docEnd', numFmt: 'lowerRoman' }
  const update = <K extends keyof NoteProps>(key: K, next: NoteProps[K]) => {
    setDraft((prev) => ({ ...prev, [kind]: { ...prev[kind], [key]: next } }))
    setTouched((prev) => ({ ...prev, [kind]: { ...prev[kind], [key]: next } }))
  }
  const validStarts = (['footnote', 'endnote'] as const).every((item) => {
    if (touched[item].numStart === undefined) return true
    const number = Number(startText[item])
    return (
      /^\d+$/.test(startText[item]) && Number.isInteger(number) && number >= 1 && number <= 32767
    )
  })
  const apply = () => {
    if (!validStarts) return
    const changes = Object.fromEntries(
      (['footnote', 'endnote'] as const)
        .filter((key) => Object.keys(touched[key]).length > 0)
        .map((key) => [key, touched[key]]),
    ) as Partial<Record<NoteKind, NoteSettingsChange>>
    if (Object.keys(changes).length) onApply(changes, scope)
    onClose()
  }
  const changeScope = (next: 'document' | 'section') => {
    if (next === 'section' && sectionIndex === null) return
    setScope(next)
    const read = (noteKind: NoteKind) =>
      next === 'section' && sectionIndex !== null
        ? sectionNoteSettings(value, noteKind, sectionIndex)
        : effectiveNoteSettings(value, noteKind)
    const footnote = read('footnote')
    const endnote = read('endnote')
    setDraft({ footnote, endnote })
    setStartText({
      footnote: String(footnote.numStart ?? 1),
      endnote: String(endnote.numStart ?? 1),
    })
    setTouched({ footnote: {}, endnote: {} })
  }
  return (
    <div
      className="modal-backdrop"
      ref={modalKeys.ref}
      onKeyDown={modalKeys.onKeyDown}
      onMouseDown={(event) => event.target === event.currentTarget && onClose()}
    >
      <div
        className="modal note-settings-dialog"
        role="dialog"
        aria-label={t('appNoteSettingsTitle')}
      >
        <h2>{t('appNoteSettingsTitle')}</h2>
        <div className="note-settings-kinds" role="tablist" aria-label={t('appNoteType')}>
          {(['footnote', 'endnote'] as const).map((item) => (
            <button
              key={item}
              type="button"
              role="tab"
              aria-selected={kind === item}
              onClick={() => setKind(item)}
            >
              {t(item === 'footnote' ? 'appNoteFootnote' : 'appNoteEndnote')}
            </button>
          ))}
        </div>
        <label>
          {t('appNotePosition')}
          <select
            value={current.pos ?? defaults.pos}
            disabled={kind === 'endnote' && scope === 'section'}
            onChange={(event) => update('pos', event.target.value as NoteProps['pos'])}
          >
            {kind === 'footnote' ? (
              <>
                <option value="pageBottom">{t('appNotePosPageBottom')}</option>
                <option value="beneathText">{t('appNotePosBeneathText')}</option>
              </>
            ) : (
              <>
                <option value="docEnd">{t('appNotePosDocEnd')}</option>
                <option value="sectEnd">{t('appNotePosSectEnd')}</option>
              </>
            )}
          </select>
        </label>
        <label>
          {t('appNumberFormat')}
          <select
            value={current.numFmt ?? defaults.numFmt}
            onChange={(event) => update('numFmt', event.target.value)}
          >
            {FORMATS.map(([value, label]) => (
              <option key={value} value={value}>
                {label}
              </option>
            ))}
            {current.numFmt && !FORMATS.some(([value]) => value === current.numFmt) && (
              <option value={current.numFmt}>{current.numFmt}</option>
            )}
          </select>
        </label>
        <label>
          {t('appStartAt')}
          <input
            type="number"
            min={1}
            max={32767}
            value={startText[kind]}
            onChange={(event) => {
              setStartText((prev) => ({ ...prev, [kind]: event.target.value }))
              update('numStart', Number(event.target.value))
            }}
          />
        </label>
        <label>
          {t('appNoteNumbering')}
          <select
            value={current.numRestart ?? 'continuous'}
            onChange={(event) =>
              update('numRestart', event.target.value as NoteProps['numRestart'])
            }
          >
            <option value="continuous">{t('appNoteRestartContinuous')}</option>
            <option value="eachSect">{t('appNoteRestartSection')}</option>
            {kind === 'footnote' && <option value="eachPage">{t('appNoteRestartPage')}</option>}
          </select>
        </label>
        <label>
          {t('appNoteScopeLabel')}
          <select
            value={scope}
            onChange={(event) => changeScope(event.target.value as typeof scope)}
          >
            <option value="document">{t('appNoteScopeWholeDocument')}</option>
            <option value="section" disabled={sectionIndex === null}>
              {t('appNoteScopeThisSection')}
            </option>
          </select>
        </label>
        <div className="modal-actions">
          <button className="btn-ghost" onClick={onClose}>
            {t('appCancel')}
          </button>
          <button className="btn-primary" disabled={!validStarts} onClick={apply}>
            {t('appOk')}
          </button>
        </div>
      </div>
    </div>
  )
}
