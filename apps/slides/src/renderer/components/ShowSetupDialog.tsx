import React, { useState } from 'react'
import type { CustomShowInfo, ShowSetup } from '../../shared/ipc'
import { useI18n } from '../i18n/locale'

export function ShowSetupDialog({
  initial,
  slideCount,
  customShows,
  onCommit,
  onClose,
}: {
  initial: ShowSetup
  slideCount: number
  customShows: CustomShowInfo[]
  onCommit: (settings: ShowSetup) => Promise<boolean>
  onClose: () => void
}) {
  const { t } = useI18n()
  const [draft, setDraft] = useState(initial)
  const [saving, setSaving] = useState(false)
  const [rangeStart, setRangeStart] = useState(
    initial.slides.kind === 'range' ? initial.slides.start : 1,
  )
  const [rangeEnd, setRangeEnd] = useState(
    initial.slides.kind === 'range' ? initial.slides.end : slideCount,
  )
  const valid =
    draft.slides.kind !== 'range' ||
    (Number.isInteger(rangeStart) &&
      Number.isInteger(rangeEnd) &&
      rangeStart >= 1 &&
      rangeEnd >= rangeStart &&
      rangeEnd <= slideCount)
  const submit = async () => {
    if (saving || !valid) return
    setSaving(true)
    const next: ShowSetup =
      draft.slides.kind === 'range'
        ? { ...draft, slides: { kind: 'range', start: rangeStart, end: rangeEnd } }
        : draft
    const ok = await onCommit(next)
    setSaving(false)
    if (ok) onClose()
  }
  return (
    <div className="modal-backdrop" onClick={onClose}>
      <div
        className="modal show-setup-dialog"
        role="dialog"
        aria-modal="true"
        aria-label={t('paneShowSetupTitle')}
        onClick={(e) => e.stopPropagation()}
      >
        <h2>{t('paneShowSetupTitle')}</h2>
        <div className="show-setup-grid">
          <fieldset>
            <legend>{t('paneShowSetupType')}</legend>
            {(['speaker', 'browse', 'kiosk'] as const).map((mode) => (
              <label key={mode}>
                <input
                  type="radio"
                  name="show-type"
                  data-testid={`show-type-${mode}`}
                  checked={draft.mode === mode}
                  onChange={() => setDraft({ ...draft, mode })}
                />
                {t(
                  mode === 'speaker'
                    ? 'paneShowTypeSpeaker'
                    : mode === 'browse'
                      ? 'paneShowTypeBrowse'
                      : 'paneShowTypeKiosk',
                )}
              </label>
            ))}
          </fieldset>
          <fieldset>
            <legend>{t('paneShowSetupOptions')}</legend>
            <label>
              <input
                type="checkbox"
                data-testid="show-loop"
                checked={draft.loop}
                onChange={(e) => setDraft({ ...draft, loop: e.target.checked })}
              />
              {t('paneShowLoop')}
            </label>
            <label>
              <input
                type="checkbox"
                data-testid="show-no-narration"
                checked={!draft.showNarration}
                onChange={(e) => setDraft({ ...draft, showNarration: !e.target.checked })}
              />
              {t('paneShowNoNarration')}
            </label>
            <label>
              <input
                type="checkbox"
                data-testid="show-no-animation"
                checked={!draft.showAnimation}
                onChange={(e) => setDraft({ ...draft, showAnimation: !e.target.checked })}
              />
              {t('paneShowNoAnimation')}
            </label>
          </fieldset>
          <fieldset>
            <legend>{t('paneShowSetupSlides')}</legend>
            <label>
              <input
                type="radio"
                name="show-slides"
                data-testid="show-slides-all"
                checked={draft.slides.kind === 'all'}
                onChange={() => setDraft({ ...draft, slides: { kind: 'all' } })}
              />
              {t('paneShowSlidesAll')}
            </label>
            <label>
              <input
                type="radio"
                name="show-slides"
                data-testid="show-slides-range"
                checked={draft.slides.kind === 'range'}
                onChange={() =>
                  setDraft({
                    ...draft,
                    slides: { kind: 'range', start: rangeStart, end: rangeEnd },
                  })
                }
              />
              {t('paneShowSlidesRange')}
            </label>
            <div className="show-setup-range">
              <label>
                {t('paneShowFrom')}
                <input
                  type="number"
                  data-testid="show-range-start"
                  min={1}
                  max={slideCount}
                  value={rangeStart}
                  disabled={draft.slides.kind !== 'range'}
                  onChange={(e) => setRangeStart(Number(e.target.value))}
                />
              </label>
              <label>
                {t('paneShowTo')}
                <input
                  type="number"
                  data-testid="show-range-end"
                  min={1}
                  max={slideCount}
                  value={rangeEnd}
                  disabled={draft.slides.kind !== 'range'}
                  onChange={(e) => setRangeEnd(Number(e.target.value))}
                />
              </label>
            </div>
            <label>
              <input
                type="radio"
                name="show-slides"
                data-testid="show-slides-custom"
                disabled={customShows.length === 0}
                checked={draft.slides.kind === 'custom'}
                onChange={() =>
                  setDraft({ ...draft, slides: { kind: 'custom', id: customShows[0]!.id } })
                }
              />
              {t('paneShowSlidesCustom')}
            </label>
            {draft.slides.kind === 'custom' && (
              <select
                data-testid="show-custom-id"
                value={draft.slides.id}
                onChange={(e) =>
                  setDraft({ ...draft, slides: { kind: 'custom', id: e.target.value } })
                }
              >
                {customShows.map((show) => (
                  <option key={show.id} value={show.id}>
                    {show.name}
                  </option>
                ))}
              </select>
            )}
          </fieldset>
          <fieldset>
            <legend>{t('paneShowSetupAdvance')}</legend>
            <label>
              <input
                type="radio"
                name="show-advance"
                data-testid="show-manual"
                checked={!draft.useTimings}
                onChange={() => setDraft({ ...draft, useTimings: false })}
              />
              {t('paneShowManual')}
            </label>
            <label>
              <input
                type="radio"
                name="show-advance"
                data-testid="show-timings"
                checked={draft.useTimings}
                onChange={() => setDraft({ ...draft, useTimings: true })}
              />
              {t('paneShowUseTimings')}
            </label>
          </fieldset>
        </div>
        <div className="show-setup-actions">
          <button onClick={onClose}>{t('paneCancel')}</button>
          <button
            data-testid="show-setup-apply"
            disabled={!valid || saving}
            onClick={() => void submit()}
          >
            {t('paneOk')}
          </button>
        </div>
      </div>
    </div>
  )
}

const EMU_PER_CM = 360000

export function CustomSlideSizeDialog({
  initial,
  onCommit,
  onClose,
}: {
  initial: { cx: number; cy: number }
  onCommit: (cx: number, cy: number) => Promise<boolean>
  onClose: () => void
}) {
  const { t } = useI18n()
  const [width, setWidth] = useState(initial.cx / EMU_PER_CM)
  const [height, setHeight] = useState(initial.cy / EMU_PER_CM)
  const [saving, setSaving] = useState(false)
  const valid = [width, height].every(
    (value) => Number.isFinite(value) && value >= 2.54 && value <= 142.24,
  )
  const portrait = height > width
  const orient = (wantPortrait: boolean) => {
    if (wantPortrait !== portrait) {
      setWidth(height)
      setHeight(width)
    }
  }
  const submit = async () => {
    if (!valid || saving) return
    setSaving(true)
    const ok = await onCommit(Math.round(width * EMU_PER_CM), Math.round(height * EMU_PER_CM))
    setSaving(false)
    if (ok) onClose()
  }
  return (
    <div className="modal-backdrop" onClick={onClose}>
      <div
        className="modal custom-slide-size-dialog"
        role="dialog"
        aria-modal="true"
        aria-label={t('paneSlideSizeTitle')}
        onClick={(e) => e.stopPropagation()}
      >
        <h2>{t('paneSlideSizeTitle')}</h2>
        <label>
          {t('paneSlideWidthCm')}
          <input
            data-testid="slide-width-cm"
            type="number"
            min="2.54"
            max="142.24"
            step="0.01"
            value={width}
            onChange={(e) => setWidth(Number(e.target.value))}
          />
        </label>
        <label>
          {t('paneSlideHeightCm')}
          <input
            data-testid="slide-height-cm"
            type="number"
            min="2.54"
            max="142.24"
            step="0.01"
            value={height}
            onChange={(e) => setHeight(Number(e.target.value))}
          />
        </label>
        <fieldset>
          <legend>{t('paneSlideOrientation')}</legend>
          <label>
            <input
              type="radio"
              name="slide-orientation"
              data-testid="slide-landscape"
              checked={!portrait}
              onChange={() => orient(false)}
            />
            {t('paneSlideLandscape')}
          </label>
          <label>
            <input
              type="radio"
              name="slide-orientation"
              data-testid="slide-portrait"
              checked={portrait}
              onChange={() => orient(true)}
            />
            {t('paneSlidePortrait')}
          </label>
        </fieldset>
        <div className="show-setup-actions">
          <button onClick={onClose}>{t('paneCancel')}</button>
          <button
            data-testid="slide-size-apply"
            disabled={!valid || saving}
            onClick={() => void submit()}
          >
            {t('paneOk')}
          </button>
        </div>
      </div>
    </div>
  )
}
