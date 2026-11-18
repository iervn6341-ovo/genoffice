import React, { useMemo } from 'react'
import type { RenderSlide } from '@genoffice/pptx-render'
import type { AnimationItem } from '../../shared/ipc'
import { animationMarkers } from '../animation-authoring'
import { useI18n } from '../i18n/locale'

/** Editor chrome only: never included in playback, printing or exported slide content. */
export function AnimationOrderOverlay({
  slide,
  items,
  zoom,
  selected,
  onSelect,
  onDelete,
}: {
  slide: RenderSlide
  items: AnimationItem[]
  zoom: number
  selected: number
  onSelect: (index: number) => void
  onDelete: (index: number) => void
}) {
  const { t } = useI18n()
  const markers = useMemo(() => animationMarkers(slide, items, zoom), [slide, items, zoom])
  return (
    <div className="animation-order-overlay">
      {markers.map((marker) => (
        <button
          key={marker.index}
          type="button"
          className={`animation-order-badge${selected === marker.index ? ' selected' : ''}`}
          data-testid={`animation-order-${marker.index}`}
          aria-label={t('paneAnimOrderLabel', {
            n: marker.step,
            target: items[marker.index]!.targetName,
          })}
          aria-pressed={selected === marker.index}
          style={{ left: marker.x, top: marker.y, transform: `scale(${1 / Math.max(0.1, zoom)})` }}
          onPointerDown={(e) => e.stopPropagation()}
          onMouseDown={(e) => e.stopPropagation()}
          onKeyDown={(e) => {
            if (e.key === 'Delete' || e.key === 'Backspace') {
              e.preventDefault()
              e.stopPropagation()
              onDelete(marker.index)
            } else if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
              e.preventDefault()
              e.stopPropagation()
              onSelect(
                Math.max(
                  0,
                  Math.min(items.length - 1, marker.index + (e.key === 'ArrowDown' ? 1 : -1)),
                ),
              )
            }
          }}
          onClick={(e) => {
            e.stopPropagation()
            onSelect(marker.index)
          }}
        >
          {marker.step}
        </button>
      ))}
    </div>
  )
}
