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
        onMouseDown={(e) => e.stopPropagation()}
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
        // a command picked from the panel closes it (the panel is the group's own items)
        onClick={(e) => {
          if (open && (e.target as Element).closest('button')) setOpen(false)
        }}
      >
        {children}
      </div>
    </Root>
  )
}
