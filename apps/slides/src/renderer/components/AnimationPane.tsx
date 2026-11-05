/**
 * Animation pane (right side) — the current page's animation order list: select/reorder/delete/change trigger.
 * The source of truth lives in the main process (the pptx <p:timing>); the list is owned by App
 * and written back wholesale (setAnimations overwrite-style); this only displays and sends back intents.
 */
import React, { useEffect, useRef } from 'react'
import type { AnimEffectKind, AnimationItem } from '../../shared/ipc'
import type { AnimationEditPatch } from '../animation-actions'
import { animationStepNumbers } from '../animation-authoring'
import { animClassOf } from '../animation-play'
import { useI18n, type StringKey } from '../i18n/locale'
import { IconSidebarCollapse } from './icons'

const EFFECT_KEY: Record<AnimEffectKind, StringKey> = {
  appear: 'paneAnimEffAppear',
  fade: 'paneAnimEffFade',
  flyIn: 'paneAnimEffFlyIn',
  wipe: 'paneAnimEffWipe',
  wipeDown: 'paneAnimEffWipeDown',
  splitIn: 'paneAnimEffSplitIn',
  bounce: 'paneAnimEffBounce',
  flipIn: 'paneAnimEffFlipIn',
  zoom: 'paneAnimEffZoom',
  pulse: 'paneAnimEffPulse',
  spin: 'paneAnimEffSpin',
  grow: 'paneAnimEffGrow',
  teeter: 'paneAnimEffTeeter',
  disappear: 'paneAnimEffDisappear',
  fadeOut: 'paneAnimEffFadeOut',
  flyOut: 'paneAnimEffFlyOut',
  wipeOut: 'paneAnimEffWipeOut',
  shrink: 'paneAnimEffShrink',
  zoomOut: 'paneAnimEffZoomOut',
  motionPath: 'paneAnimEffMotionPath',
}

const TRIGGER_GLYPH = { onClick: '🖱', withPrev: '⇉', afterPrev: '⏱' } as const
const TRIGGER_KEY = {
  onClick: 'paneAnimTrigOnClick',
  withPrev: 'paneAnimTrigWithPrev',
  afterPrev: 'paneAnimTrigAfterPrev',
} as const satisfies Record<string, StringKey>

interface Props {
  slideIndex: number
  items: AnimationItem[]
  /** Pane selection (the timing controls bind to it); -1 = none */
  selected: number
  onSelect: (idx: number) => void
  onPatch: (idx: number, patch: AnimationEditPatch) => void
  /** Move up/down one position */
  onMove: (idx: number, dir: -1 | 1) => void
  onDelete: (idx: number) => void
  /** Play preview */
  onPreview: () => void
  onCollapse: () => void
}

export function AnimationPane({
  slideIndex,
  items,
  selected,
  onSelect,
  onPatch,
  onMove,
  onDelete,
  onPreview,
  onCollapse,
}: Props) {
  const { t } = useI18n()
  const stepNos = animationStepNumbers(items)
  const listRef = useRef<HTMLDivElement>(null)
  const active = items[selected]
  useEffect(() => {
    listRef.current
      ?.querySelector<HTMLElement>(`[data-animation-index="${selected}"]`)
      ?.scrollIntoView?.({ block: 'nearest' })
  }, [selected, items])

  return (
    <aside
      className="comments-pane anim-pane"
      onKeyDown={(e) => {
        if (!e.metaKey && !e.ctrlKey) e.stopPropagation()
      }}
    >
      <div className="ai-panel-header">
        <span className="ai-panel-title">{t('paneAnimTitle', { n: slideIndex + 1 })}</span>
        <div className="ai-panel-header-actions">
          <button
            className="ai-header-btn"
            disabled={items.length === 0}
            onClick={onPreview}
            data-tip={t('paneAnimPreview')}
            aria-label={t('paneAnimPreview')}
          >
            ▶
          </button>
          <button
            className="ai-header-btn"
            onClick={onCollapse}
            data-tip={t('paneAnimCollapse')}
            aria-label={t('paneAnimCollapse')}
          >
            <IconSidebarCollapse size={15} />
          </button>
        </div>
      </div>

      <div
        ref={listRef}
        className="comments-list"
        role="listbox"
        aria-label={t('paneAnimTitle', { n: slideIndex + 1 })}
      >
        {items.length === 0 && (
          <div className="comments-empty">
            <p>{t('paneAnimEmpty')}</p>
            <p className="comments-empty-sub">{t('paneAnimEmptySub')}</p>
          </div>
        )}
        {items.map((it, i) => (
          <div
            key={i}
            className={`anim-row anim-${animClassOf(it.effect)} ${selected === i ? 'selected' : ''}`}
            data-testid={`animation-row-${i}`}
            data-animation-index={i}
            role="option"
            aria-selected={selected === i}
            tabIndex={selected === i || (selected < 0 && i === 0) ? 0 : -1}
            onClick={() => onSelect(i)}
            onKeyDown={(e) => {
              if (e.target !== e.currentTarget) return
              if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
                e.preventDefault()
                const next = Math.max(
                  0,
                  Math.min(items.length - 1, i + (e.key === 'ArrowDown' ? 1 : -1)),
                )
                onSelect(next)
                listRef.current
                  ?.querySelector<HTMLElement>(`[data-animation-index="${next}"]`)
                  ?.focus()
              } else if (e.key === 'Enter' || e.key === ' ') {
                e.preventDefault()
                onSelect(i)
              } else if (e.key === 'Delete' || e.key === 'Backspace') {
                e.preventDefault()
                onDelete(i)
              }
            }}
            data-tip={`${t(TRIGGER_KEY[it.trigger])} · ${t('paneAnimSeconds', { sec: (it.durationMs / 1000).toFixed(2) })}${it.delayMs ? ` · ${t('paneAnimDelaySeconds', { sec: (it.delayMs / 1000).toFixed(2) })}` : ''}`}
          >
            <span className="anim-row-no">{stepNos[i]}</span>
            <span className="anim-row-trigger" data-tip={t(TRIGGER_KEY[it.trigger])}>
              {TRIGGER_GLYPH[it.trigger]}
            </span>
            <span className="anim-row-main">
              <span className="anim-row-effect">
                {it.customPreset ? t('paneAnimImportedEffect') : t(EFFECT_KEY[it.effect])}
              </span>
              <span className="anim-row-target">
                {it.targetName}
                {it.paragraph != null && ` · ${t('paneAnimParaN', { n: it.paragraph + 1 })}`}
              </span>
            </span>
            <span className="anim-row-actions" onClick={(e) => e.stopPropagation()}>
              <button
                disabled={i === 0}
                data-tip={t('paneMoveUp')}
                aria-label={t('paneMoveUp')}
                onClick={() => onMove(i, -1)}
              >
                ▲
              </button>
              <button
                disabled={i === items.length - 1}
                data-tip={t('paneMoveDown')}
                aria-label={t('paneMoveDown')}
                onClick={() => onMove(i, 1)}
              >
                ▼
              </button>
              <button
                data-tip={t('paneAnimDelete')}
                aria-label={t('paneAnimDelete')}
                onClick={() => onDelete(i)}
              >
                ✕
              </button>
            </span>
          </div>
        ))}
      </div>
      <div className="anim-properties" data-testid="animation-properties">
        {!active ? (
          <p className="anim-select-hint">{t('paneAnimSelectHint')}</p>
        ) : (
          <>
            <div className="anim-properties-target">{active.targetName}</div>
            <details open>
              <summary>{t('ribbonTransEffectOptions')}</summary>
              <label>
                {t('ribbonGroupAnimation')}
                <select
                  aria-label={t('ribbonGroupAnimation')}
                  value={active.customPreset ? 'imported' : active.effect}
                  onChange={(e) => onPatch(selected, { effect: e.target.value as AnimEffectKind })}
                >
                  {active.customPreset && (
                    <option value="imported" disabled>
                      {t('paneAnimImportedEffect')}
                    </option>
                  )}
                  {(Object.keys(EFFECT_KEY) as AnimEffectKind[]).map((effect) => (
                    <option key={effect} value={effect}>
                      {t(EFFECT_KEY[effect])}
                    </option>
                  ))}
                </select>
              </label>
              <label>
                {t('ribbonTransEffectOptions')}
                <select
                  aria-label={t('ribbonTransEffectOptions')}
                  disabled={
                    active.customPreset ||
                    !['flyIn', 'flyOut', 'wipe', 'wipeOut'].includes(active.effect)
                  }
                  value={active.direction ?? 'bottom'}
                  onChange={(e) =>
                    onPatch(selected, { direction: e.target.value as AnimationItem['direction'] })
                  }
                >
                  {(['top', 'bottom', 'left', 'right'] as const).map((direction) => (
                    <option key={direction} value={direction}>
                      {t(
                        (
                          {
                            top: 'ribbonTransFromTop',
                            bottom: 'ribbonTransFromBottom',
                            left: 'ribbonTransFromLeft',
                            right: 'ribbonTransFromRight',
                          } as const
                        )[direction],
                      )}
                    </option>
                  ))}
                </select>
              </label>
              <label>
                {t('ribbonAnimAfterEffect')}
                <select
                  aria-label={t('ribbonAnimAfterEffect')}
                  disabled={active.customPreset}
                  value={active.after ?? 'stay'}
                  onChange={(e) =>
                    onPatch(selected, {
                      after:
                        e.target.value === 'stay'
                          ? null
                          : (e.target.value as AnimationItem['after']),
                    })
                  }
                >
                  <option value="stay">{t('ribbonAnimStayVisible')}</option>
                  <option value="hideAfter">{t('ribbonAnimHideAfter')}</option>
                  <option value="hideOnNextClick">{t('ribbonAnimHideNextClick')}</option>
                </select>
              </label>
            </details>
            <details open>
              <summary>{t('ribbonGroupTiming')}</summary>
              <label>
                {t('ribbonAnimStart')}
                <select
                  aria-label={t('ribbonAnimStart')}
                  value={active.trigger}
                  onChange={(e) =>
                    onPatch(selected, { trigger: e.target.value as AnimationItem['trigger'] })
                  }
                >
                  {(Object.keys(TRIGGER_KEY) as Array<keyof typeof TRIGGER_KEY>).map((trigger) => (
                    <option key={trigger} value={trigger}>
                      {t(TRIGGER_KEY[trigger])}
                    </option>
                  ))}
                </select>
              </label>
              {(['durationMs', 'delayMs'] as const).map((field) => (
                <label key={field}>
                  {t(field === 'durationMs' ? 'ribbonAnimDuration' : 'ribbonAnimDelay')}
                  <span className="anim-seconds-input">
                    <input
                      key={`${selected}:${active.nativeId ?? ''}:${active[field]}`}
                      type="number"
                      min={0}
                      step={0.1}
                      aria-label={t(
                        field === 'durationMs' ? 'ribbonAnimDuration' : 'ribbonAnimDelay',
                      )}
                      disabled={
                        field === 'durationMs' && ['appear', 'disappear'].includes(active.effect)
                      }
                      defaultValue={(active[field] / 1000).toFixed(2)}
                      onBlur={(e) => {
                        const value = Number(e.currentTarget.value)
                        if (
                          e.currentTarget.value.trim() &&
                          Number.isFinite(value) &&
                          value >= 0 &&
                          Math.round(value * 1000) !== active[field]
                        )
                          onPatch(selected, { [field]: Math.round(value * 1000) })
                        else e.currentTarget.value = (active[field] / 1000).toFixed(2)
                      }}
                      onKeyDown={(e) => {
                        if (e.key === 'Enter') e.currentTarget.blur()
                      }}
                    />
                    {t('ribbonSecondsUnit')}
                  </span>
                </label>
              ))}
            </details>
          </>
        )}
      </div>
    </aside>
  )
}
