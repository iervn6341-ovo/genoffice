import { useEffect, useRef, useState, type ReactNode } from 'react'

/**
 * A ribbon group that folds into one dropdown button when the window is too narrow, like the
 * groups of Microsoft 365's ribbon. The fit hook (useRibbonLabelFit) decides which groups fold:
 * lowest `priority` first, and only after every text label is off. The children render once;
 * folded, they become the dropdown panel opened by the button, so no command is duplicated.
 *
 * The class names follow each app's own ribbon markup (Docs/Markdown/HTML: `ribbon-group` +
 * `ribbon-group-items` + `rb-big`; Sheets: `ribbon-group` + `ribbon-group-content` + its tool
 * button classes), so the group looks the same as its neighbours when unfolded.
 */
export function RibbonFoldGroup({
  label,
  icon,
  priority,
  children,
  className,
  itemsClassName = 'ribbon-group-items',
  buttonClassName = 'rb-big',
  iconClassName = 'rb-big-icon',
  caret,
  as = 'div',
  itemsRole,
  caption,
  preserveEditorFocus = false,
}: {
  /** group name: the folded button's caption and the accessible name */
  label: string
  /** the folded button's icon */
  icon: ReactNode
  /** fold order: lower folds first */
  priority: number
  children: ReactNode
  className?: string | undefined
  itemsClassName?: string | undefined
  buttonClassName?: string | undefined
  iconClassName?: string | undefined
  /** dropdown chevron drawn after the icon (the app's own caret glyph) */
  caret?: ReactNode | undefined
  as?: 'div' | 'section' | undefined
  /** ARIA role of the items container (e.g. a view switcher's `tablist`) */
  itemsRole?: string | undefined
  /** show the group name under the items while unfolded (Insert-style groups with a caption row) */
  caption?: boolean | undefined
  /** Keep a live editor's caret when this folded launcher opens its panel. */
  preserveEditorFocus?: boolean | undefined
}) {
  const [open, setOpen] = useState(false)
  const rootRef = useRef<HTMLElement | null>(null)
  useEffect(() => {
    if (!open) return
    const close = () => setOpen(false)
    const onDown = (e: MouseEvent) => {
      if (!rootRef.current?.contains(e.target as Node)) close()
    }
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') close()
    }
    document.addEventListener('mousedown', onDown)
    document.addEventListener('keydown', onKey)
    window.addEventListener('resize', close)
    return () => {
      document.removeEventListener('mousedown', onDown)
      document.removeEventListener('keydown', onKey)
      window.removeEventListener('resize', close)
    }
  }, [open])
  const Root = as
  return (
    <Root
      ref={rootRef as never}
      className={`ribbon-group rb-fold-group${open ? ' rb-fold-open' : ''}${className ? ` ${className}` : ''}`}
      data-fold-priority={priority}
      aria-label={label}
    >
      <button
        type="button"
        className={`${buttonClassName} rb-fold-btn${open ? ' active' : ''}`}
        data-tip={label}
        aria-expanded={open}
        onPointerDown={preserveEditorFocus ? (e) => e.preventDefault() : undefined}
        onMouseDown={(e) => {
          if (preserveEditorFocus) e.preventDefault()
          e.stopPropagation()
        }}
        onClick={() => setOpen((v) => !v)}
      >
        <span className={iconClassName}>
          {icon}
          {caret}
        </span>
        <span>{label}</span>
      </button>
      <div
        className={`${itemsClassName} rb-fold-items`}
        role={itemsRole}
        // Keep the group open while a child button opens a nested picker. A click
        // inside that picker is a completed command and closes the group.
        onClick={(e) => {
          if (!open) return
          const target = e.target as Element
          const button = target.closest('button')
          if (!button) return
          if (button.closest('[data-rb-keep-open]')) return
          if (button.closest('[data-rb-panel], .menu-select-drop, .gcp-palette')) {
            setOpen(false)
            return
          }
          // React flushes the child's dropdown state after the click handler.
          // Check the resulting DOM before treating its launcher as a leaf.
          window.setTimeout(() => {
            const root = rootRef.current
            if (!root?.contains(button)) return
            if (button.getAttribute('aria-expanded') === 'true') return
            const panels = root.querySelectorAll('[data-rb-panel], .menu-select-drop, .gcp-palette')
            if (![...panels].some((panel) => getComputedStyle(panel).display !== 'none')) {
              setOpen(false)
            }
          }, 0)
        }}
      >
        {children}
      </div>
      {caption && <div className="ribbon-group-label rb-fold-caption">{label}</div>}
    </Root>
  )
}
