/**
 * PowerPoint's slide show controls, shared by presenter view and the full-screen show:
 * Pointer Options (laser pointer, pen, highlighter, eraser, erase all ink, ink colours), Zoom,
 * Black Screen and the "…" Slide Show Options menu with its submenus.
 */
import React, { useEffect, useRef, useState } from 'react'
import type { RenderNode, RenderSlide, ShapeRenderNode } from '@genoffice/pptx-render'
import { useI18n, type TFunc } from '../i18n/locale'
import { INK_COLORS, type InkTool } from './ShowInk'

// ── Icons (PowerPoint for Mac's show glyphs, outline style) ────────────────────────

function Glyph({ size = 22, children }: { size?: number; children: React.ReactNode }) {
  return (
    <svg
      width={size}
      height={size}
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth={1.5}
      strokeLinecap="round"
      strokeLinejoin="round"
      aria-hidden
    >
      {children}
    </svg>
  )
}

export const ShowIcon = {
  endShow: (s?: number) => (
    <Glyph size={s}>
      <circle cx="12" cy="12" r="9" fill="currentColor" stroke="none" />
      <path d="M9 9l6 6M15 9l-6 6" stroke="var(--slides-show-black)" strokeWidth={2} />
    </Glyph>
  ),
  swap: (s?: number) => (
    <Glyph size={s}>
      <path d="M4 8h14l-3-3M20 16H6l3 3" />
    </Glyph>
  ),
  slideShow: (s?: number) => (
    <Glyph size={s}>
      <rect x="4" y="7" width="14" height="11" rx="1.5" fill="currentColor" />
      <path d="M7 4h11.5A1.5 1.5 0 0 1 20 5.5V15" />
    </Glyph>
  ),
  pen: (s?: number) => (
    <Glyph size={s}>
      <path d="M4 20l1.2-4.5L16 4.7a1.8 1.8 0 0 1 2.6 0l.7.7a1.8 1.8 0 0 1 0 2.6L8.5 18.8z" />
      <path d="M14.5 6.2l3.3 3.3" />
    </Glyph>
  ),
  laser: (s?: number) => (
    <Glyph size={s}>
      <path d="M5 19l9.5-9.5" strokeWidth={2.4} />
      <circle cx="17.5" cy="6.5" r="2.6" fill="currentColor" stroke="none" />
    </Glyph>
  ),
  highlighter: (s?: number) => (
    <Glyph size={s}>
      <path d="M13 5l6 6-7 7H8v-4z" fill="currentColor" />
      <path d="M4 20h7" />
    </Glyph>
  ),
  eraser: (s?: number) => (
    <Glyph size={s}>
      <path
        d="M8 19l-4-4 9.5-9.5a1.6 1.6 0 0 1 2.3 0l3.7 3.7a1.6 1.6 0 0 1 0 2.3L12 19z"
        fill="currentColor"
      />
      <path d="M12 19h8" />
    </Glyph>
  ),
  eraseAll: (s?: number) => (
    <Glyph size={s}>
      <rect x="3" y="5" width="15" height="11" rx="1.5" />
      <path
        d="M14 19l-2-2 5-5a1 1 0 0 1 1.4 0l1.6 1.6a1 1 0 0 1 0 1.4L16 19z"
        fill="currentColor"
      />
    </Glyph>
  ),
  zoom: (s?: number) => (
    <Glyph size={s}>
      <circle cx="10.5" cy="10.5" r="6" />
      <path d="M15 15l5 5M8 10.5h5M10.5 8v5" />
    </Glyph>
  ),
  black: (s?: number) => (
    <Glyph size={s}>
      <rect x="3" y="4" width="18" height="11" rx="1" />
      <path d="M12 15v5M8 20h8M3 4l18 11" />
    </Glyph>
  ),
  more: (s?: number) => (
    <Glyph size={s}>
      <circle cx="12" cy="12" r="9" />
      <circle cx="8" cy="12" r="0.9" fill="currentColor" />
      <circle cx="12" cy="12" r="0.9" fill="currentColor" />
      <circle cx="16" cy="12" r="0.9" fill="currentColor" />
    </Glyph>
  ),
  camera: (s?: number) => (
    <Glyph size={s}>
      <rect x="3" y="7" width="12" height="10" rx="2" />
      <path d="M15 11l6-3v8l-6-3" />
    </Glyph>
  ),
  prev: (s?: number) => (
    <Glyph size={s}>
      <path d="M15 5l-7 7 7 7" strokeWidth={2} />
    </Glyph>
  ),
  next: (s?: number) => (
    <Glyph size={s}>
      <path d="M9 5l7 7-7 7" strokeWidth={2} />
    </Glyph>
  ),
}

// ── Slide titles for "By Title" ─────────────────────────────────────────────────────

function nodeText(n: RenderNode): string {
  const text = (n as ShapeRenderNode).text
  if (!text) return ''
  return text.lines
    .map((l) =>
      l.runs
        .filter((r) => !r.isBullet)
        .map((r) => r.text)
        .join(''),
    )
    .join(' ')
    .replace(/\s+/g, ' ')
    .trim()
}

/** A slide's title: its title placeholder, else its topmost text (PowerPoint's "By Title" list) */
export function slideTitle(slide: RenderSlide): string {
  const texts = slide.nodes.filter(
    (n) => (n.type === 'shape' || n.type === 'text') && !n.decoration && nodeText(n),
  )
  const title = texts.find((n) => /title/i.test((n as ShapeRenderNode).placeholder ?? ''))
  const top = title ?? [...texts].sort((a, b) => a.box.y - b.box.y)[0]
  return top ? nodeText(top).slice(0, 60) : ''
}

// ── Menus ──────────────────────────────────────────────────────────────────────────

export type ShowMenuItem =
  | 'sep'
  | {
      label: string
      onSelect?: () => void
      disabled?: boolean
      checked?: boolean
      submenu?: ShowMenuItem[]
    }

/** A PowerPoint-style menu (dark, native-looking) with hover-opened submenus */
export function ShowMenu({
  items,
  onDone,
  className,
}: {
  items: ShowMenuItem[]
  /** After a command runs: close the whole menu */
  onDone: () => void
  className?: string
}): React.JSX.Element {
  const [open, setOpen] = useState<number | null>(null)
  return (
    <div className={`ssc-menu${className ? ` ${className}` : ''}`} role="menu">
      {items.map((it, i) =>
        it === 'sep' ? (
          <div key={i} className="ssc-menu-sep" role="separator" />
        ) : (
          <div
            key={i}
            className="ssc-menu-row"
            onPointerEnter={() => setOpen(it.submenu ? i : null)}
          >
            <button
              type="button"
              role="menuitem"
              className={`ssc-menu-item${open === i ? ' ssc-open' : ''}`}
              disabled={it.disabled || (it.submenu && !it.submenu.length)}
              aria-haspopup={it.submenu ? 'menu' : undefined}
              aria-checked={it.checked}
              onClick={() => {
                if (it.submenu) {
                  setOpen(open === i ? null : i)
                  return
                }
                it.onSelect?.()
                onDone()
              }}
            >
              <span className="ssc-menu-check">{it.checked ? '✓' : ''}</span>
              <span className="ssc-menu-label">{it.label}</span>
              {it.submenu && <span className="ssc-menu-chev">›</span>}
            </button>
            {it.submenu && open === i && (
              <ShowMenu items={it.submenu} onDone={onDone} className="ssc-submenu" />
            )}
          </div>
        ),
      )}
    </div>
  )
}

/** The Pointer Options popover: tools, Erase All Ink, and the ink colours */
export function PointerOptions({
  tool,
  color,
  canErase,
  onTool,
  onColor,
  onEraseAll,
}: {
  tool: InkTool
  color: string
  canErase: boolean
  onTool: (tool: InkTool) => void
  onColor: (hex: string) => void
  onEraseAll: () => void
}): React.JSX.Element {
  const { t } = useI18n()
  const toolRow = (id: Exclude<InkTool, 'none'>, label: string, icon: React.ReactNode) => (
    <button
      type="button"
      role="menuitemcheckbox"
      aria-checked={tool === id}
      className={`ssc-pop-item${tool === id ? ' ssc-on' : ''}`}
      disabled={id === 'eraser' && !canErase}
      onClick={() => onTool(tool === id ? 'none' : id)}
    >
      {icon}
      <span>{label}</span>
    </button>
  )
  return (
    <div className="ssc-pop" role="menu" aria-label={t('panePresenterPointerOptions')}>
      {toolRow('laser', t('panePresenterLaserPointer'), ShowIcon.laser(18))}
      {toolRow('pen', t('panePresenterPenTool'), ShowIcon.pen(18))}
      {toolRow('highlighter', t('panePresenterHighlighter'), ShowIcon.highlighter(18))}
      {toolRow('eraser', t('panePresenterEraser'), ShowIcon.eraser(18))}
      <button
        type="button"
        role="menuitem"
        className="ssc-pop-item"
        disabled={!canErase}
        onClick={onEraseAll}
      >
        {ShowIcon.eraseAll(18)}
        <span>{t('panePresenterEraseInk')}</span>
      </button>
      <div className="ssc-swatches" role="group">
        {INK_COLORS.map((c) => (
          <button
            key={c.key}
            type="button"
            className={`ssc-swatch${c.hex === color ? ' ssc-on' : ''}`}
            style={{ background: c.hex }}
            aria-label={t(INK_COLOR_KEYS[c.key])}
            data-tip={t(INK_COLOR_KEYS[c.key])}
            onClick={() => onColor(c.hex)}
          />
        ))}
      </div>
    </div>
  )
}

const INK_COLOR_KEYS = {
  white: 'paneInkWhite',
  black: 'paneInkBlack',
  darkRed: 'paneInkDarkRed',
  red: 'paneInkRed',
  orange: 'paneInkOrange',
  yellow: 'paneInkYellow',
  lightGreen: 'paneInkLightGreen',
  darkGreen: 'paneInkDarkGreen',
  lightBlue: 'paneInkLightBlue',
  mediumBlue: 'paneInkMediumBlue',
  darkBlue: 'paneInkDarkBlue',
  purple: 'paneInkPurple',
} as const

/** Which popover of a toolbar is open */
type OpenPanel = 'pointer' | 'more' | null

/**
 * The tool row: Pointer Options, Zoom, Black Screen, Slide Show Options (…). `showNav` adds
 * the previous / next arrows the full-screen control bar starts with.
 */
export function ShowToolbar({
  tool,
  inkColor,
  canErase,
  zoomOn,
  black,
  cameraOn,
  menuItems,
  menuUp,
  showNav,
  onTool,
  onColor,
  onEraseAll,
  onZoom,
  onBlack,
  onCamera,
  onPrev,
  onNext,
  onOpenChange,
}: {
  tool: InkTool
  inkColor: string
  canErase: boolean
  zoomOn: boolean
  black: boolean
  cameraOn?: boolean
  menuItems: ShowMenuItem[]
  /** Open the menus upward (the full-screen bar sits at the bottom of the screen) */
  menuUp?: boolean
  showNav?: boolean
  onTool: (tool: InkTool) => void
  onColor: (hex: string) => void
  onEraseAll: () => void
  onZoom: () => void
  onBlack: () => void
  /** PowerPoint's Camera: live camera bubble on the slide */
  onCamera?: () => void
  onPrev?: () => void
  onNext?: () => void
  /** A popover opened or closed (Esc closes it before it ends the show) */
  onOpenChange?: (open: boolean) => void
}): React.JSX.Element {
  const { t } = useI18n()
  const [open, setOpenState] = useState<OpenPanel>(null)
  const rootRef = useRef<HTMLDivElement>(null)
  const setOpen = (p: OpenPanel) => {
    setOpenState(p)
    onOpenChange?.(p !== null)
  }
  const setOpenRef = useRef(setOpen)
  setOpenRef.current = setOpen
  useEffect(() => {
    if (!open) return
    const onDown = (e: PointerEvent) => {
      if (!rootRef.current?.contains(e.target as Node)) setOpenRef.current(null)
    }
    document.addEventListener('pointerdown', onDown, true)
    return () => document.removeEventListener('pointerdown', onDown, true)
  }, [open])
  // Esc from the show keys closes the open popover (see closeShowPopovers)
  useEffect(() => {
    const close = () => setOpenRef.current(null)
    window.addEventListener(CLOSE_POPOVERS_EVENT, close)
    return () => window.removeEventListener(CLOSE_POPOVERS_EVENT, close)
  }, [])

  const btn = (
    label: string,
    icon: React.ReactNode,
    onClick: () => void,
    on = false,
    extra: React.ButtonHTMLAttributes<HTMLButtonElement> = {},
  ) => (
    <button
      type="button"
      className={`ssc-btn${on ? ' ssc-on' : ''}`}
      aria-label={label}
      data-tip={label}
      onClick={onClick}
      {...extra}
    >
      {icon}
    </button>
  )

  return (
    <div
      ref={rootRef}
      className={`ssc-bar${menuUp ? ' ssc-up' : ''}`}
      onClick={(e) => e.stopPropagation()}
      onPointerDown={(e) => e.stopPropagation()}
    >
      {showNav && onPrev && btn(t('panePresenterPrevTip'), ShowIcon.prev(20), onPrev)}
      {showNav && onNext && btn(t('panePresenterNextTip'), ShowIcon.next(20), onNext)}
      <div className="ssc-anchor">
        {btn(
          t('panePresenterPointerOptions'),
          tool === 'laser'
            ? ShowIcon.laser()
            : tool === 'highlighter'
              ? ShowIcon.highlighter()
              : tool === 'eraser'
                ? ShowIcon.eraser()
                : ShowIcon.pen(),
          () => setOpen(open === 'pointer' ? null : 'pointer'),
          tool !== 'none' || open === 'pointer',
          { 'aria-haspopup': 'menu', 'aria-expanded': open === 'pointer' },
        )}
        {open === 'pointer' && (
          <PointerOptions
            tool={tool}
            color={inkColor}
            canErase={canErase}
            onTool={(v) => {
              onTool(v)
              setOpen(null)
            }}
            onColor={(hex) => {
              onColor(hex)
              setOpen(null)
            }}
            onEraseAll={() => {
              onEraseAll()
              setOpen(null)
            }}
          />
        )}
      </div>
      {btn(t('panePresenterZoom'), ShowIcon.zoom(), onZoom, zoomOn, { 'aria-pressed': zoomOn })}
      {btn(t('panePresenterBlackTip'), ShowIcon.black(), onBlack, black, { 'aria-pressed': black })}
      {onCamera &&
        btn(t('panePresenterCamera'), ShowIcon.camera(), onCamera, !!cameraOn, {
          'aria-pressed': !!cameraOn,
        })}
      <div className="ssc-anchor">
        {btn(
          t('panePresenterShowOptions'),
          ShowIcon.more(),
          () => setOpen(open === 'more' ? null : 'more'),
          open === 'more',
          { 'aria-haspopup': 'menu', 'aria-expanded': open === 'more' },
        )}
        {open === 'more' && (
          <ShowMenu items={menuItems} onDone={() => setOpen(null)} className="ssc-more" />
        )}
      </div>
    </div>
  )
}

const CLOSE_POPOVERS_EVENT = 'genoffice-show-close-popovers'

/** Close any open show toolbar popover (Esc closes a menu before it ends the show) */
export function closeShowPopovers(): void {
  window.dispatchEvent(new Event(CLOSE_POPOVERS_EVENT))
}

/** Build the "…" Slide Show Options menu shared by presenter view and the full-screen show */
export function showOptionsMenu(
  t: TFunc,
  o: {
    slides: RenderSlide[]
    order: number[]
    pos: number
    ended: boolean
    lastViewed: number | null
    cover: 'none' | 'black' | 'white'
    tool: InkTool
    canErase: boolean
    presenter: boolean
    canSwap: boolean
    paused?: boolean
    onNext: () => void
    onPrev: () => void
    onGoto: (orderPos: number) => void
    onCover: (c: 'black' | 'white') => void
    onSwap?: () => void
    onToggleView?: () => void
    onTool: (tool: InkTool) => void
    onEraseAll: () => void
    onPause?: () => void
    onEnd: () => void
  },
): ShowMenuItem[] {
  const pointer: ShowMenuItem[] = [
    {
      label: t('panePresenterLaserPointer'),
      checked: o.tool === 'laser',
      onSelect: () => o.onTool(o.tool === 'laser' ? 'none' : 'laser'),
    },
    {
      label: t('panePresenterPenTool'),
      checked: o.tool === 'pen',
      onSelect: () => o.onTool(o.tool === 'pen' ? 'none' : 'pen'),
    },
    {
      label: t('panePresenterHighlighter'),
      checked: o.tool === 'highlighter',
      onSelect: () => o.onTool(o.tool === 'highlighter' ? 'none' : 'highlighter'),
    },
    {
      label: t('panePresenterEraser'),
      checked: o.tool === 'eraser',
      disabled: !o.canErase,
      onSelect: () => o.onTool(o.tool === 'eraser' ? 'none' : 'eraser'),
    },
    'sep',
    { label: t('panePresenterEraseInk'), disabled: !o.canErase, onSelect: o.onEraseAll },
  ]
  const byTitle: ShowMenuItem[] = o.order.map((idx, p) => {
    const s = o.slides[idx]
    const title = s ? slideTitle(s) : ''
    return {
      label: `${p + 1} ${title || t('panePresenterFilmSlideN', { n: idx + 1 })}`,
      checked: p === o.pos && !o.ended,
      onSelect: () => o.onGoto(p),
    }
  })
  return [
    { label: t('paneShowMenuNext'), onSelect: o.onNext, disabled: o.ended },
    { label: t('paneShowMenuPrev'), onSelect: o.onPrev, disabled: o.pos === 0 && !o.ended },
    {
      label: t('paneShowMenuLastViewed'),
      disabled: o.lastViewed == null,
      onSelect: () => o.lastViewed != null && o.onGoto(o.lastViewed),
    },
    'sep',
    { label: t('paneShowMenuByTitle'), submenu: byTitle },
    'sep',
    {
      label: t('paneShowMenuScreen'),
      submenu: [
        {
          label: t('paneShowMenuBlackScreen'),
          checked: o.cover === 'black',
          onSelect: () => o.onCover('black'),
        },
        {
          label: t('paneShowMenuWhiteScreen'),
          checked: o.cover === 'white',
          onSelect: () => o.onCover('white'),
        },
      ],
    },
    ...(o.onSwap
      ? [{ label: t('panePresenterSwap'), disabled: !o.canSwap, onSelect: o.onSwap }]
      : []),
    ...(o.onToggleView
      ? [
          {
            label: o.presenter ? t('panePresenterUseShow') : t('paneShowMenuUsePresenterView'),
            onSelect: o.onToggleView,
          },
        ]
      : []),
    { label: t('panePresenterPointerOptions'), submenu: pointer },
    'sep',
    ...(o.onPause
      ? [
          {
            label: o.paused ? t('paneShowMenuResume') : t('paneShowMenuPause'),
            onSelect: o.onPause,
          },
        ]
      : []),
    { label: t('panePresenterEndShow'), onSelect: o.onEnd },
  ]
}

/**
 * PowerPoint's pointer shortcuts during a show: Ctrl+L laser pointer, Ctrl+P pen, Ctrl+I
 * highlighter, Ctrl+E eraser, Ctrl+A arrow, and E erases the slide's ink. The ⌘ variants stay
 * with the app menu (⌘P prints).
 */
export function usePointerShortcuts(setTool: (tool: InkTool) => void, eraseAll: () => void): void {
  const ref = useRef({ setTool, eraseAll })
  ref.current = { setTool, eraseAll }
  useEffect(() => {
    const pick: Record<string, InkTool> = {
      l: 'laser',
      p: 'pen',
      i: 'highlighter',
      e: 'eraser',
      a: 'none',
    }
    const onKey = (e: KeyboardEvent) => {
      if (e.altKey || e.metaKey || e.shiftKey) return
      if (!e.ctrlKey) {
        if (e.key === 'e' || e.key === 'E') {
          e.preventDefault()
          ref.current.eraseAll()
        }
        return
      }
      const next = pick[e.key.toLowerCase()]
      if (!next) return
      e.preventDefault()
      ref.current.setTool(next)
    }
    window.addEventListener('keydown', onKey, true)
    return () => window.removeEventListener('keydown', onKey, true)
  }, [])
}
