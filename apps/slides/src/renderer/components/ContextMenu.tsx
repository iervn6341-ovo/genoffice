/**
 * Context menu — fixed positioning, kept within the viewport; closes on any click/Escape/scroll.
 * Items that are null render as separators.
 */
import React, { useEffect, useLayoutEffect, useRef, useState } from 'react'
import { platformShortcuts } from '@genoffice/i18n'
import { placeContextMenu } from '../context-menu-position'

export interface CtxItem {
  label: string
  /** Shortcut hint on the right */
  hint?: string
  disabled?: boolean
  danger?: boolean
  onClick?: () => void
  /** Row of color swatches rendered under the label instead of a clickable item ('none' = the clear square) */
  swatches?: string[]
  onSwatch?: (color: string) => void
}

interface Props {
  x: number
  y: number
  items: Array<CtxItem | null>
  onClose: () => void
}

export function ContextMenu({ x, y, items, onClose }: Props) {
  const ref = useRef<HTMLDivElement>(null)
  const [place, setPlace] = useState<{ x: number; y: number; maxHeight: number | null }>({
    x,
    y,
    maxHeight: null,
  })

  // Re-place whenever the anchor or the item list changes (the thumbnail menu gains/loses rows
  // after the clipboard probe resolves) and whenever the window is resized.
  useLayoutEffect(() => {
    const el = ref.current
    if (!el) return
    const measure = () => {
      // scrollHeight is the uncapped content height even when CSS max-height is clipping it
      const natural = el.scrollHeight + (el.offsetHeight - el.clientHeight)
      setPlace(
        placeContextMenu(x, y, el.offsetWidth, natural, window.innerWidth, window.innerHeight),
      )
    }
    measure()
    window.addEventListener('resize', measure)
    return () => window.removeEventListener('resize', measure)
  }, [x, y, items])

  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') onClose()
    }
    const onDown = (e: MouseEvent) => {
      if (!ref.current?.contains(e.target as Node)) onClose()
    }
    window.addEventListener('keydown', onKey)
    window.addEventListener('mousedown', onDown)
    // Wheeling over a menu that scrolls (taller than the window) scrolls it; anywhere else closes it.
    const onWheel = (e: WheelEvent) => {
      const el = ref.current
      if (el && el.contains(e.target as Node) && el.scrollHeight > el.clientHeight) return
      onClose()
    }
    window.addEventListener('wheel', onWheel)
    window.addEventListener('blur', onClose)
    return () => {
      window.removeEventListener('keydown', onKey)
      window.removeEventListener('mousedown', onDown)
      window.removeEventListener('wheel', onWheel)
      window.removeEventListener('blur', onClose)
    }
  }, [onClose])

  const swatchRow = (swatches: string[], onSwatch?: (color: string) => void) => (
    <div className="ctx-swatches-row">
      {swatches.map((c) => (
        <button
          key={c}
          className={`ctx-swatch ${c === 'none' ? 'ctx-swatch-none' : ''}`}
          style={c === 'none' ? undefined : { background: c }}
          data-tip={c}
          aria-label={c}
          onClick={() => {
            onClose()
            onSwatch?.(c)
          }}
        >
          {c === 'none' ? '✕' : ''}
        </button>
      ))}
    </div>
  )

  return (
    <div
      ref={ref}
      className="ctx-menu"
      style={{
        left: place.x,
        top: place.y,
        ...(place.maxHeight != null ? { maxHeight: place.maxHeight } : null),
      }}
      onContextMenu={(e) => e.preventDefault()}
    >
      {items.map((item, i) =>
        item === null ? (
          <div key={i} className="ctx-sep" />
        ) : item.swatches ? (
          <div key={i} className="ctx-swatches">
            <span className="ctx-swatches-label">{item.label}</span>
            {swatchRow(item.swatches, item.onSwatch)}
          </div>
        ) : (
          <button
            key={i}
            className={`ctx-item ${item.danger ? 'danger' : ''}`}
            disabled={item.disabled}
            onClick={() => {
              onClose()
              item.onClick?.()
            }}
          >
            <span>{item.label}</span>
            {item.hint && <span className="ctx-hint">{platformShortcuts(item.hint)}</span>}
          </button>
        ),
      )}
    </div>
  )
}
