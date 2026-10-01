import { useLayoutEffect, useRef, type ReactNode } from 'react'
import { useI18n } from '../i18n/locale'

export type InsertGalleryKey = 'chart' | 'smartart' | 'shapes' | 'icons' | 'wordart'

export function isInsertGallery(key: string): key is InsertGalleryKey {
  return ['chart', 'smartart', 'shapes', 'icons', 'wordart'].includes(key)
}

/** Non-modal gallery: the workspace reserves its width, leaving the slide editable. */
export function RibbonGalleryPane({
  title,
  children,
  onClose,
}: {
  title: string
  children: ReactNode
  onClose: () => void
}) {
  const { t } = useI18n()
  const bodyRef = useRef<HTMLDivElement>(null)
  useLayoutEffect(() => {
    bodyRef.current?.querySelector<HTMLButtonElement>('button:not(:disabled)')?.focus()
  }, [])

  return (
    <section
      className="ribbon-gallery-pane"
      aria-labelledby="ribbon-gallery-title"
      onMouseDown={(e) => e.stopPropagation()}
      onKeyDown={(e) => {
        // Gallery navigation must not move/delete slides via document shortcuts.
        e.stopPropagation()
        if (e.key === 'Escape') {
          e.preventDefault()
          onClose()
          return
        }
        if (!['ArrowDown', 'ArrowUp', 'ArrowLeft', 'ArrowRight', 'Home', 'End'].includes(e.key))
          return
        const buttons = Array.from(
          bodyRef.current?.querySelectorAll<HTMLButtonElement>('button:not(:disabled)') ?? [],
        )
        const index = buttons.indexOf(e.target as HTMLButtonElement)
        if (index < 0 || buttons.length === 0) return
        e.preventDefault()
        if (e.key === 'Home' || e.key === 'End') {
          buttons[e.key === 'Home' ? 0 : buttons.length - 1]?.focus()
          return
        }
        const target = buttons[index]!
        const grid = target.parentElement!
        const siblings = buttons.filter((button) => button.parentElement === grid)
        const style = getComputedStyle(grid)
        const columns = style.display === 'grid' ? style.gridTemplateColumns.split(' ').length : 1
        const step = e.key === 'ArrowDown' || e.key === 'ArrowUp' ? columns : 1
        const next =
          siblings.indexOf(target) + (e.key === 'ArrowUp' || e.key === 'ArrowLeft' ? -step : step)
        siblings[next]?.focus()
      }}
    >
      <header className="ribbon-gallery-header">
        <h2 id="ribbon-gallery-title">{title}</h2>
        <button type="button" aria-label={t('paneCsdClose')} onClick={onClose}>
          ×
        </button>
      </header>
      <div ref={bodyRef} className="ribbon-gallery-body">
        {children}
      </div>
    </section>
  )
}
