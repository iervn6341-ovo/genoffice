import { useState, type ReactNode } from 'react'

/**
 * One collapsible group of the print pane (Office's Printer / Settings / Page Setup / Options
 * headings): keeps the sheet short while every setting stays one click away.
 */
export function PrintSection({
  title,
  defaultOpen = true,
  children,
}: {
  title: string
  defaultOpen?: boolean
  children: ReactNode
}) {
  const [open, setOpen] = useState(defaultOpen)
  return (
    <section className={`print-section${open ? ' open' : ''}`}>
      <button
        type="button"
        className="print-section-head"
        aria-expanded={open}
        onClick={() => setOpen((v) => !v)}
      >
        <svg width="10" height="10" viewBox="0 0 10 10" aria-hidden="true">
          <path d="M3 1.5 7 5 3 8.5" />
        </svg>
        {title}
      </button>
      {open && <div className="print-section-body">{children}</div>}
    </section>
  )
}

/** A labeled row inside a print section: label on the left, the control on the right. */
export function PrintRow({ label, children }: { label: string; children: ReactNode }) {
  return (
    <label className="ps-field">
      <span>{label}</span>
      {children}
    </label>
  )
}

export function PrintCheck({
  label,
  checked,
  disabled,
  onChange,
}: {
  label: string
  checked: boolean
  disabled?: boolean
  onChange: (next: boolean) => void
}) {
  return (
    <label className="ps-check">
      <input
        type="checkbox"
        checked={checked}
        disabled={disabled === true}
        onChange={(e) => onChange(e.target.checked)}
      />
      {label}
    </label>
  )
}
