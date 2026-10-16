/**
 * Custom show management dialog — modeled on PowerPoint "Slide Show → Custom Shows".
 *
 * Left column: show list (new/delete/play); right side: the selected show's name, included-page
 * checkboxes and order adjustment. Changes are committed as one native undoable transaction.
 */
import React, { useState } from 'react'
import { useI18n } from '../i18n/locale'
import type { CustomShow } from '../slideshow-utils'

/** Allocate an OOXML unsigned-int identity without changing imported IDs. */
function newShowId(shows: CustomShow[]): string {
  const used = new Set(shows.map((show) => show.id))
  for (let id = 0; id <= 0xffffffff; id++) if (!used.has(String(id))) return String(id)
  throw new Error('No custom-show IDs remain')
}

export function CustomShowDialog({
  shows,
  slideCount,
  onCommit,
  onPlay,
  onClose,
}: {
  shows: CustomShow[]
  /** Document page count (checkbox list; stale out-of-range indexes filtered at render/show time) */
  slideCount: number
  onCommit: (shows: CustomShow[]) => Promise<boolean>
  onPlay: (show: CustomShow) => void
  onClose: () => void
}) {
  const { t } = useI18n()
  const [draft, setDraft] = useState(() =>
    shows.map((show) => ({ ...show, slideIndices: [...show.slideIndices] })),
  )
  const [saving, setSaving] = useState(false)
  const [selId, setSelId] = useState<string | null>(shows[0]?.id ?? null)
  const sel = draft.find((s) => s.id === selId) ?? null
  // Out-of-range indexes left over after page deletion aren't shown (the show entry filters them too)
  const selIndices = sel ? sel.slideIndices.filter((i) => i >= 0 && i < slideCount) : []

  const patchSel = (patch: Partial<CustomShow>) => {
    if (sel) setDraft(draft.map((s) => (s.id === sel.id ? { ...s, ...patch } : s)))
  }

  const addShow = () => {
    const show: CustomShow = {
      id: newShowId(draft),
      name: t('paneCsdDefaultName', { n: draft.length + 1 }),
      slideIndices: [],
    }
    setDraft([...draft, show])
    setSelId(show.id)
  }

  const removeShow = () => {
    if (!sel) return
    const rest = draft.filter((s) => s.id !== sel.id)
    setDraft(rest)
    setSelId(rest[0]?.id ?? null)
  }

  /** Check/uncheck a page: checking appends to the end of the playback order */
  const togglePage = (i: number) => {
    if (!sel) return
    patchSel({
      slideIndices: selIndices.includes(i) ? selIndices.filter((x) => x !== i) : [...selIndices, i],
    })
  }

  /** Move an item up/down in the playback order */
  const movePage = (posInShow: number, dir: -1 | 1) => {
    const next = posInShow + dir
    if (!sel || next < 0 || next >= selIndices.length) return
    const arr = selIndices.slice()
    ;[arr[posInShow], arr[next]] = [arr[next]!, arr[posInShow]!]
    patchSel({ slideIndices: arr })
  }

  const finish = async (play = false) => {
    if (saving || draft.some((show) => !show.name.trim())) return
    setSaving(true)
    const ok = await onCommit(draft)
    setSaving(false)
    if (!ok) return
    if (play && sel) onPlay(sel)
    else onClose()
  }

  return (
    <div className="modal-backdrop" onClick={() => void finish()}>
      <div className="modal csd" onClick={(e) => e.stopPropagation()}>
        <h2>{t('paneCsdTitle')}</h2>
        <div className="csd-body">
          <div className="csd-shows">
            <div className="csd-list" role="listbox">
              {draft.length === 0 && <div className="csd-empty">{t('paneCsdEmpty')}</div>}
              {draft.map((s) => (
                <button
                  key={s.id}
                  data-testid={`custom-show-${s.id}`}
                  className={`csd-item${s.id === selId ? ' active' : ''}`}
                  onClick={() => setSelId(s.id)}
                >
                  {s.name}
                  <span className="csd-item-count">
                    {t('paneCsdPageCount', {
                      n: s.slideIndices.filter((i) => i < slideCount).length,
                    })}
                  </span>
                </button>
              ))}
            </div>
            <div className="csd-list-actions">
              <button data-testid="custom-show-new" disabled={saving} onClick={addShow}>
                {t('paneCsdNew')}
              </button>
              <button
                data-testid="custom-show-delete"
                disabled={!sel || saving}
                onClick={removeShow}
              >
                {t('paneCsdDelete')}
              </button>
            </div>
          </div>
          {sel ? (
            <div className="csd-detail">
              <label>
                {t('paneCsdName')}
                <input
                  type="text"
                  data-testid="custom-show-name"
                  value={sel.name}
                  onChange={(e) => patchSel({ name: e.target.value })}
                  placeholder={t('paneCsdNamePlaceholder')}
                />
              </label>
              <div className="csd-columns">
                <div className="csd-col">
                  <div className="csd-col-title">{t('paneCsdAllSlides')}</div>
                  <div className="csd-list csd-pages">
                    {Array.from({ length: slideCount }, (_, i) => (
                      <label key={i} className="csd-page">
                        <input
                          type="checkbox"
                          data-testid={`custom-show-slide-${i}`}
                          checked={selIndices.includes(i)}
                          onChange={() => togglePage(i)}
                        />
                        {t('paneCsdSlideN', { n: i + 1 })}
                      </label>
                    ))}
                  </div>
                </div>
                <div className="csd-col">
                  <div className="csd-col-title">{t('paneCsdOrder')}</div>
                  <div className="csd-list csd-pages">
                    {selIndices.length === 0 && (
                      <div className="csd-empty">{t('paneCsdOrderEmpty')}</div>
                    )}
                    {selIndices.map((slideIdx, posInShow) => (
                      <div key={`${slideIdx}-${posInShow}`} className="csd-order-item">
                        <span>
                          {posInShow + 1}. {t('paneCsdSlideN', { n: slideIdx + 1 })}
                        </span>
                        <span className="csd-order-btns">
                          <button
                            disabled={saving || posInShow === 0}
                            data-testid={`custom-show-up-${posInShow}`}
                            onClick={() => movePage(posInShow, -1)}
                            data-tip={t('paneMoveUp')}
                            aria-label={t('paneMoveUp')}
                          >
                            ↑
                          </button>
                          <button
                            disabled={saving || posInShow === selIndices.length - 1}
                            data-testid={`custom-show-down-${posInShow}`}
                            onClick={() => movePage(posInShow, 1)}
                            data-tip={t('paneMoveDown')}
                            aria-label={t('paneMoveDown')}
                          >
                            ↓
                          </button>
                        </span>
                      </div>
                    ))}
                  </div>
                </div>
              </div>
            </div>
          ) : (
            <div className="csd-detail csd-empty">{t('paneCsdCreateHint')}</div>
          )}
        </div>
        <div className="modal-actions">
          <button
            data-testid="custom-show-close"
            disabled={saving || draft.some((show) => !show.name.trim())}
            onClick={() => void finish()}
          >
            {t('paneCsdClose')}
          </button>
          <button
            className="primary"
            disabled={
              !sel || selIndices.length === 0 || saving || draft.some((show) => !show.name.trim())
            }
            data-testid="custom-show-play"
            onClick={() => void finish(true)}
          >
            {t('paneCsdPlay')}
          </button>
        </div>
      </div>
    </div>
  )
}
