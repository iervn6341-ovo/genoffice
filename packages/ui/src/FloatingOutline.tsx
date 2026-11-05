import { useEffect, useId, useMemo, useRef, useState } from 'react'

export interface FloatingOutlineItem {
  id: string
  text: string
  level: number
  page?: string
}

/** An overlay navigator: no document reflow or layout measurement on resize. */
export function FloatingOutline({
  items,
  title,
  searchLabel,
  emptyLabel,
  closeLabel,
  note,
  onSelect,
  activeId,
  open,
  onOpenChange,
}: {
  items: FloatingOutlineItem[]
  title: string
  searchLabel: string
  emptyLabel: string
  closeLabel: string
  note?: string
  onSelect: (id: string) => void
  activeId: string | null
  open?: boolean
  onOpenChange?: (open: boolean) => void
}) {
  const panelId = useId()
  const trigger = useRef<HTMLButtonElement>(null)
  const [hovered, setHovered] = useState(false)
  const [pinned, setPinned] = useState(false)
  const [dismissed, setDismissed] = useState(false)
  const [search, setSearch] = useState('')
  const list = useRef<HTMLElement>(null)
  const isPinned = open ?? pinned
  const setPinnedOpen = (value: boolean) => {
    setPinned(value)
    onOpenChange?.(value)
  }
  const expanded = (hovered || isPinned) && (!dismissed || isPinned)
  const filtered = useMemo(() => {
    const query = search.trim().toLocaleLowerCase()
    return query ? items.filter((item) => item.text.toLocaleLowerCase().includes(query)) : items
  }, [items, search])
  const miniItems = useMemo(() => {
    const active = items.findIndex((item) => item.id === activeId)
    const start = Math.max(0, Math.min(active - 12, items.length - 24))
    return items.slice(start, start + 24)
  }, [items, activeId])
  const close = () => {
    setPinnedOpen(false)
    setHovered(false)
    setDismissed(true)
    trigger.current?.focus({ preventScroll: true })
  }

  useEffect(() => {
    const panel = list.current
    const item = panel?.querySelector<HTMLElement>('[aria-current="location"]')
    if (!panel || !item) return
    // Move only the list; scrollIntoView would also move the document viewport.
    const top = item.offsetTop - panel.offsetTop
    if (top < panel.scrollTop) panel.scrollTop = top
    else if (top + item.offsetHeight > panel.scrollTop + panel.clientHeight) {
      panel.scrollTop = top + item.offsetHeight - panel.clientHeight
    }
  }, [activeId, expanded, filtered])

  if (items.length === 0) return null
  return (
    <div className="floating-document-outline-host">
      <div
        className={`floating-document-outline${expanded ? ' is-expanded' : ''}`}
        onMouseEnter={() => {
          setHovered(true)
          setDismissed(false)
        }}
        onMouseLeave={() => setHovered(false)}
        onKeyDown={(event) => {
          if (event.key === 'Escape') {
            event.stopPropagation()
            close()
          }
        }}
      >
        <button
          ref={trigger}
          type="button"
          className="floating-document-outline-trigger"
          aria-label={title}
          title={title}
          aria-expanded={expanded}
          aria-controls={expanded ? panelId : undefined}
          onClick={() => {
            setPinnedOpen(!isPinned)
            setDismissed(isPinned)
          }}
        >
          <svg
            width="16"
            height="16"
            viewBox="0 0 24 24"
            fill="none"
            stroke="currentColor"
            strokeWidth="1.5"
            aria-hidden="true"
          >
            <path d="M8 6h13M8 12h13M8 18h13M3 6h1M3 12h1M3 18h1" />
          </svg>
          <span className="floating-document-outline-tracks" aria-hidden="true">
            {miniItems.map((item) => (
              <span
                key={item.id}
                className={item.id === activeId ? 'is-active' : undefined}
                style={{ width: Math.max(5, 17 - item.level * 3) }}
              />
            ))}
          </span>
        </button>
        {expanded && (
          <aside id={panelId} className="floating-document-outline-card" aria-label={title}>
            <header>
              <strong>{title}</strong>
              <span className="floating-document-outline-count">{items.length}</span>
              <button type="button" aria-label={closeLabel} title={closeLabel} onClick={close}>
                ×
              </button>
            </header>
            {note && <div className="floating-document-outline-note">{note}</div>}
            {items.length > 5 && (
              <input
                type="search"
                aria-label={searchLabel}
                placeholder={searchLabel}
                value={search}
                onChange={(event) => setSearch(event.target.value)}
              />
            )}
            <nav ref={list} className="floating-document-outline-list" aria-label={title}>
              {filtered.map((item) => (
                <button
                  type="button"
                  key={item.id}
                  className="floating-document-outline-item"
                  aria-current={activeId === item.id ? 'location' : undefined}
                  style={{
                    paddingInlineStart: 12 + (Math.min(6, Math.max(1, item.level)) - 1) * 12,
                  }}
                  title={item.text}
                  onClick={() => {
                    onSelect(item.id)
                  }}
                >
                  <span className="floating-document-outline-text">{item.text}</span>
                  {item.page && <span className="floating-document-outline-page">{item.page}</span>}
                </button>
              ))}
              {filtered.length === 0 && (
                <div className="floating-document-outline-note" role="status">
                  {emptyLabel}
                </div>
              )}
            </nav>
          </aside>
        )}
      </div>
    </div>
  )
}
