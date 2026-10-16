import { useEffect, useMemo, useRef, useState, type ReactNode } from 'react'
import type { Lang } from '@genoffice/i18n'
import { backstageLabels } from './strings-backstage'

export interface BackstageExport {
  id: string
  label: string
  disabled?: boolean
  run: () => void
}

export interface BackstageRecent {
  path: string
  name: string
}

export interface BackstageProps {
  lang: Lang
  /** which blank-file tile the New pane offers */
  kind: 'doc' | 'book' | 'deck'
  docName: string
  filePath?: string | null
  /** a document is open (Save / Save As / Print / Export need one) */
  hasDoc: boolean
  canSave: boolean
  /** extra rows (pages, words, …) already translated by the host app */
  infoRows?: ReadonlyArray<{ label: string; value: string }>
  recent: readonly BackstageRecent[]
  exports: readonly BackstageExport[]
  onClose: () => void
  onNew: () => void
  onOpen: () => void
  onOpenRecent: (path: string) => void
  onSave: () => void
  onSaveAs: () => void
  onPrint?: () => void
  /** When set, Print opens a pane listing these (e.g. print scopes) instead of running onPrint */
  printItems?: readonly BackstageExport[] | undefined
}

/** keep in step with the bs-slide-out duration in backstage.css */
const CLOSE_MS = 140

type Pane = 'new' | 'open' | 'info' | 'export' | 'print'

interface NavItem {
  id: string
  label: string
  icon: ReactNode
  needsDoc?: boolean
  action?: 'save' | 'saveAs' | 'print'
}

const svg = (children: ReactNode) => (
  <svg width="18" height="18" viewBox="0 0 20 20" aria-hidden="true">
    {children}
  </svg>
)

const ICONS = {
  new: svg(<path d="M5 2.5h6.5L15 6v11.5H5zM11 2.5V6.5h4M10 9.5v5M7.5 12h5" />),
  open: svg(<path d="M2.5 15.5v-10h5l1.5 2h8v8zM2.5 15.5l2.5-6h13" />),
  info: svg(
    <>
      <circle cx="10" cy="10" r="7" />
      <path d="M10 9v5M10 6.2v.1" />
    </>,
  ),
  save: svg(<path d="M3.5 3.5h10l3 3v10h-13zM6.5 3.5v4h6v-4M6.5 16.5v-5h7v5" />),
  saveAs: svg(<path d="M3.5 3.5h10l3 3v5M6.5 3.5v4h6v-4M3.5 3.5v13h6M13 17l4-4-1.5-1.5-4 4V17z" />),
  print: svg(<path d="M6 7V3h8v4M6 14H3.5V7h13v7H14M6 11.5h8V17H6z" />),
  export: svg(<path d="M10 12.5V3M6.5 6.5 10 3l3.5 3.5M3.5 11v6h13v-6" />),
  back: svg(<path d="M16 10H4M9 5l-5 5 5 5" />),
}

function dirOf(path: string): string {
  const i = Math.max(path.lastIndexOf('/'), path.lastIndexOf('\\'))
  return i > 0 ? path.slice(0, i) : path
}

/**
 * The File tab's full-window view (Office "backstage"): a left rail of file commands and a content
 * pane. Home/New/Open/Info/Export show a pane; Save/Save As/Print run their command and return to
 * the document.
 */
export function Backstage(props: BackstageProps) {
  const l = backstageLabels(props.lang)
  const [pane, setPane] = useState<Pane | null>(null)
  const rootRef = useRef<HTMLDivElement>(null)
  // the pressed button unmounts when a pane swaps in: keep focus inside so Esc still reaches us
  useEffect(() => {
    rootRef.current?.focus()
  }, [pane])

  // closing plays the slide-out before the host unmounts us (no wait under reduced motion)
  const [closing, setClosing] = useState(false)
  const closeTimer = useRef<number | undefined>(undefined)
  useEffect(() => () => window.clearTimeout(closeTimer.current), [])
  const requestClose = () => {
    if (closing) return
    setClosing(true)
    const still = window.matchMedia('(prefers-reduced-motion: reduce)').matches
    closeTimer.current = window.setTimeout(props.onClose, still ? 0 : CLOSE_MS)
  }
  const act = (fn: () => void) => () => {
    requestClose()
    fn()
  }
  const blankLabel =
    props.kind === 'book' ? l.blankBook : props.kind === 'deck' ? l.blankDeck : l.blankDoc

  const nav = useMemo<Array<NavItem | 'sep'>>(
    () => [
      { id: 'new', label: l.navNew, icon: ICONS.new },
      { id: 'open', label: l.navOpen, icon: ICONS.open },
      'sep',
      { id: 'info', label: l.navInfo, icon: ICONS.info, needsDoc: true },
      { id: 'save', label: l.navSave, icon: ICONS.save, needsDoc: true, action: 'save' as const },
      {
        id: 'saveAs',
        label: l.navSaveAs,
        icon: ICONS.saveAs,
        needsDoc: true,
        action: 'saveAs' as const,
      },
      ...(props.onPrint
        ? [
            {
              id: 'print',
              label: l.navPrint,
              icon: ICONS.print,
              needsDoc: true,
              ...(props.printItems?.length ? {} : { action: 'print' as const }),
            },
          ]
        : []),
      ...(props.exports.length
        ? [{ id: 'export', label: l.navExport, icon: ICONS.export, needsDoc: true }]
        : []),
    ],
    [l, props.onPrint, props.printItems?.length, props.exports.length],
  )

  const runAction = (action: string) => {
    if (action === 'save') act(props.onSave)()
    else if (action === 'saveAs') act(props.onSaveAs)()
    else if (action === 'print' && props.onPrint) act(props.onPrint)()
  }

  const recentList = (
    <ul className="bs-recent">
      {props.recent.length === 0 && <li className="bs-empty">{l.noRecent}</li>}
      {props.recent.map((r) => (
        <li key={r.path}>
          <button
            type="button"
            className="bs-recent-item"
            onClick={act(() => props.onOpenRecent(r.path))}
          >
            <span className="bs-recent-name">{r.name}</span>
            <span className="bs-recent-dir">{dirOf(r.path)}</span>
          </button>
        </li>
      ))}
    </ul>
  )

  const newTiles = (
    <div className="bs-tiles">
      <button type="button" className="bs-tile" onClick={act(props.onNew)}>
        <span className="bs-tile-thumb" aria-hidden="true">
          <span className="bs-tile-plus">+</span>
        </span>
        <span className="bs-tile-label">{blankLabel}</span>
      </button>
    </div>
  )

  const infoRows = [
    { label: l.propName, value: props.docName },
    {
      label: l.propLocation,
      value: props.filePath ? dirOf(props.filePath) : l.propUnsaved,
    },
    ...(props.infoRows ?? []),
  ]

  const paneTitle = pane ? nav.find((n) => n !== 'sep' && n.id === pane) : null

  return (
    <div
      className={`backstage${closing ? ' closing' : ''}`}
      ref={rootRef}
      tabIndex={-1}
      role="dialog"
      aria-modal="true"
      onKeyDown={(e) => {
        if (e.key === 'Escape') {
          e.stopPropagation()
          if (pane) setPane(null)
          else requestClose()
        }
      }}
    >
      <aside className="bs-drawer">
        {pane === null ? (
          <nav className="bs-rail" aria-label={l.navOpen}>
            <button
              type="button"
              className="bs-back"
              onClick={requestClose}
              aria-label={l.back}
              title={l.back}
            >
              {ICONS.back}
            </button>
            {nav.map((item, i) =>
              item === 'sep' ? (
                <hr key={`sep${i}`} className="bs-sep" />
              ) : (
                <button
                  key={item.id}
                  type="button"
                  className="bs-nav"
                  disabled={item.needsDoc ? !props.hasDoc : false}
                  onClick={() => (item.action ? runAction(item.action) : setPane(item.id as Pane))}
                >
                  {item.icon}
                  <span>{item.label}</span>
                </button>
              ),
            )}
          </nav>
        ) : (
          <div className="bs-body">
            <button type="button" className="bs-subhead" onClick={() => setPane(null)}>
              {ICONS.back}
              <span>{paneTitle && paneTitle !== 'sep' ? paneTitle.label : ''}</span>
            </button>
            {pane === 'new' && newTiles}
            {pane === 'open' && (
              <>
                <button type="button" className="bs-primary" onClick={act(props.onOpen)}>
                  {l.browse}
                </button>
                <h2 className="bs-h">{l.secRecent}</h2>
                {recentList}
              </>
            )}
            {pane === 'info' && (
              <dl className="bs-info">
                {infoRows.map((r) => (
                  <div key={r.label} className="bs-info-row">
                    <dt>{r.label}</dt>
                    <dd>{r.value}</dd>
                  </div>
                ))}
              </dl>
            )}
            {pane === 'print' && (
              <div className="bs-exports">
                {props.printItems?.map((x) => (
                  <button
                    key={x.id}
                    type="button"
                    className="bs-export"
                    disabled={x.disabled || !props.hasDoc}
                    onClick={act(x.run)}
                  >
                    {x.label}
                  </button>
                ))}
              </div>
            )}
            {pane === 'export' && (
              <div className="bs-exports">
                {props.exports.map((x) => (
                  <button
                    key={x.id}
                    type="button"
                    className="bs-export"
                    disabled={x.disabled || !props.hasDoc}
                    onClick={act(x.run)}
                  >
                    {x.label}
                  </button>
                ))}
              </div>
            )}
          </div>
        )}
      </aside>
      <div className="bs-scrim" onMouseDown={requestClose} />
    </div>
  )
}
