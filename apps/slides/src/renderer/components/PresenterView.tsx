/**
 * Presenter view for slide shows.
 *
 * Multi-screen: on entry the main process detects displays and opens a full-screen audience
 * show window (AudienceView) on the external screen. This view is the only driver — page
 * number/animation cursor/blackout/ink are all broadcast over IPC (absolute state, the audience
 * side seeks idempotently), and audience clicks/keys are sent back here for arbitration.
 * Single screen: this window only.
 *
 * Layout: top bar (end show/swap displays/switch to normal show),
 * left column timer (pause/reset) + clock → big frame → toolbar (pen/laser/clear ink/blackout)
 * + page progress navigation, right column next-slide preview + notes (adjustable font size),
 * bottom full-width thumbnail strip with click-to-jump.
 * Keyboard (PowerPoint's): N/→/↓/space/enter/PgDn next step; P/←/↑/Backspace/PgUp previous page; Home/End first/last;
 * B or . black screen, W or , white screen; a number then Enter jumps to that slide; Esc exit.
 */
import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react'
import type { RenderSlide } from '@genoffice/pptx-render'
import type { AnimationItem, NotesParagraphView, ShowSyncState } from '../../shared/ipc'
import { AnimatedSlideStage, useAnimPlayer } from './AnimatedSlide'
import { useI18n } from '../i18n/locale'
import { SlideThumb } from '../SlideThumb'
import {
  DEFAULT_HIGHLIGHTER_COLOR,
  DEFAULT_PEN_COLOR,
  applyInkEvent,
  InkBoard,
  InkLayer,
  useInkPointer,
  type InkTool,
} from './ShowInk'
import {
  closeShowPopovers,
  ShowIcon,
  ShowToolbar,
  showOptionsMenu,
  usePointerShortcuts,
} from './ShowControls'
import { NotesView } from './NotesEditor'
import { CameraBubble } from './ShowCamera'
import { readShowMonitor } from '../show-monitor'
import { liftShowCurtain } from '../show-actions'
import { toggleShowScreen, type ShowKeyCommand, type ShowScreen } from '../slideshow-utils'
import { useShowKeys } from '../use-show-keys'

/** Layout constants (aligned with styles.css) */
const IS_MAC = navigator.platform.toLowerCase().includes('mac')
/** Side column (next slide + notes) share of the screen: PowerPoint's default split, draggable */
const SIDE_FRAC_DEFAULT = 0.34
const SIDE_FRAC_MIN = 0.2
const SIDE_FRAC_MAX = 0.6
const SIDE_FRAC_KEY = 'genoffice.slides.presenterSideFrac'
const TOP_H = 50
const TIMER_H = 40
const BOTTOM_H = 56
const FILM_H = 118

/** Elapsed show seconds → mm:ss (past 1 hour minutes keep accumulating) */
function fmtElapsed(sec: number): string {
  const m = Math.floor(sec / 60)
  const s = sec % 60
  return `${String(m).padStart(2, '0')}:${String(s).padStart(2, '0')}`
}

function fmtWallClock(): string {
  const d = new Date()
  return `${d.getHours()}:${String(d.getMinutes()).padStart(2, '0')}`
}

export function PresenterView({
  slides,
  images,
  startAt,
  onExit,
  onUseSlideShow,
}: {
  slides: RenderSlide[]
  images: Map<string, HTMLImageElement>
  /** Start page (original index) */
  startAt: number
  /** Exit presenter view; lastIndex is the original index of the page dwelt on (for locating back in the edit view) */
  onExit: (lastIndex: number) => void
  /** "Switch to normal show": exit this view and start a normal show from lastIndex */
  onUseSlideShow?: (lastIndex: number) => void
}) {
  const { t } = useI18n()
  // Playback sequence (original indexes): hidden pages skipped; the start page kept even if hidden (consistent with SlideShowView)
  const order = useMemo(() => {
    const o = slides.map((_, i) => i).filter((i) => !slides[i]!.hidden || i === startAt)
    return o.length > 0 ? o : [startAt]
  }, [slides, startAt])
  const [pos, setPos] = useState(() => Math.max(0, order.indexOf(startAt)))
  const [ended, setEnded] = useState(false)
  const [cover, setCover] = useState<ShowScreen>('none')
  const black = cover === 'black'
  const white = cover === 'white'
  const [size, setSize] = useState({ w: window.innerWidth, h: window.innerHeight })
  /** Per-page animation lists + notes (prefetched once on entry, zero IPC on page turns) */
  const [allAnims, setAllAnims] = useState<AnimationItem[][] | null>(null)
  const [allNotes, setAllNotes] = useState<NotesParagraphView[][]>([])
  /** How the current page was entered: forward = initial state playing step by step, others = all-finished state */
  const navModeRef = useRef<'fresh' | 'all'>('fresh')
  /** Whether an external audience window was opened (swap-displays button availability) */
  const [hasAudience, setHasAudience] = useState(false)

  // ── Timer (pausable/resettable) and clock ──────────────────────────────────
  const [elapsed, setElapsed] = useState(0)
  const [paused, setPaused] = useState(false)
  const [clock, setClock] = useState(fmtWallClock)
  useEffect(() => {
    if (paused) return
    const t = window.setInterval(() => setElapsed((s) => s + 1), 1000)
    return () => window.clearInterval(t)
  }, [paused])
  useEffect(() => {
    const t = window.setInterval(() => setClock(fmtWallClock()), 10_000)
    return () => window.clearInterval(t)
  }, [])

  // ── Notes font size ──────────────────────────────────────────────────
  const [noteSize, setNoteSize] = useState(15)

  // ── Side column width: drag the divider (remembered on this machine) ─────────
  const [sideFrac, setSideFrac] = useState(() => {
    try {
      const v = parseFloat(localStorage.getItem(SIDE_FRAC_KEY) ?? '')
      return v >= SIDE_FRAC_MIN && v <= SIDE_FRAC_MAX ? v : SIDE_FRAC_DEFAULT
    } catch {
      return SIDE_FRAC_DEFAULT
    }
  })
  const sideW = Math.round(Math.max(260, size.w * sideFrac))
  const startSideDrag = (e: React.PointerEvent) => {
    e.preventDefault()
    const clamp = (x: number) =>
      Math.min(SIDE_FRAC_MAX, Math.max(SIDE_FRAC_MIN, (window.innerWidth - x) / window.innerWidth))
    let last = sideFrac
    const move = (ev: PointerEvent) => {
      last = clamp(ev.clientX)
      setSideFrac(last)
    }
    const up = () => {
      window.removeEventListener('pointermove', move)
      window.removeEventListener('pointerup', up)
      try {
        localStorage.setItem(SIDE_FRAC_KEY, String(last))
      } catch {
        // storage unavailable: the split lasts for this show
      }
    }
    window.addEventListener('pointermove', move)
    window.addEventListener('pointerup', up)
  }

  // ── Pointer Options: laser / pen / highlighter / eraser ─────────────────────────
  // Ink lives on the board, not in React state: a pointer move must not re-render this view
  const [tool, setToolState] = useState<InkTool>('none')
  const [penColor, setPenColor] = useState<string>(DEFAULT_PEN_COLOR)
  const [hlColor, setHlColor] = useState<string>(DEFAULT_HIGHLIGHTER_COLOR)
  const inkColor = tool === 'highlighter' ? hlColor : penColor
  const board = useMemo(() => new InkBoard(), [])
  const [canErase, setCanErase] = useState(false)
  useEffect(() => board.subscribe(() => setCanErase(board.strokes.length > 0)), [board])
  const stageboxRef = useRef<HTMLDivElement>(null)
  const ink = useInkPointer(board, tool, inkColor, stageboxRef, (ev) =>
    window.slidesApi.presenterInk(ev),
  )

  // ── Zoom (PowerPoint's magnifier): arm, then click the slide to magnify that spot 2× ──
  const [zoomArmed, setZoomArmed] = useState(false)
  const [zoom, setZoom] = useState<{ x: number; y: number } | null>(null)
  const zoomOn = zoomArmed || zoom !== null
  const toggleZoom = useCallback(() => {
    if (zoomArmed || zoom) {
      setZoomArmed(false)
      setZoom(null)
      return
    }
    setToolState('none')
    setZoomArmed(true)
  }, [zoomArmed, zoom])
  const setTool = useCallback((v: InkTool) => {
    setToolState(v)
    setZoomArmed(false)
    setZoom(null)
  }, [])
  const pickColor = useCallback(
    (hex: string) => {
      // PowerPoint: a colour applies to the highlighter when it is active, otherwise to the pen
      if (tool === 'highlighter') setHlColor(hex)
      else {
        setPenColor(hex)
        if (tool !== 'pen') setTool('pen')
      }
    },
    [tool, setTool],
  )
  // Ink made with the mouse on the audience screen (laser, pen…) shows here too
  useEffect(() => window.slidesApi.onShowInk((ev) => applyInkEvent(board, ev)), [board])
  /** Camera (PowerPoint's show toolbar): live camera bubble on both screens */
  const [camera, setCamera] = useState(false)
  /** A toolbar popover is open: Esc closes it instead of ending the show */
  const popoverOpenRef = useRef(false)
  /** The slide shown before the current one ("Last Viewed") */
  const [lastViewed, setLastViewed] = useState<number | null>(null)

  useEffect(() => {
    let cancelled = false
    void Promise.all(slides.map((_, i) => window.slidesApi.getAnimations(i))).then((lists) => {
      if (!cancelled) setAllAnims(lists)
    })
    void Promise.all(slides.map((_, i) => window.slidesApi.getNotesRich(i))).then((notes) => {
      if (!cancelled) setAllNotes(notes)
    })
    return () => {
      cancelled = true
    }
  }, [slides])

  const slide = slides[order[pos]!]
  const player = useAnimPlayer(slide?.heightPx ?? 540, slide?.widthPx ?? 960)

  // Load the page's animations when the page changes/prefetch completes (forward = initial state, back/jump = finished state);
  // loadedRef ticks with player.epoch; the broadcast effect uses it to read the "in place" page number
  const loadedRef = useRef<{ idx: number; fresh: boolean }>({ idx: order[pos]!, fresh: true })
  useEffect(() => {
    loadedRef.current = { idx: order[pos]!, fresh: navModeRef.current === 'fresh' }
    player.load(allAnims?.[order[pos]!] ?? [], navModeRef.current)
  }, [allAnims, pos, order, player.load]) // eslint-disable-line react-hooks/exhaustive-deps

  // ── Multi-screen: open the audience window on entry, close on exit ─────────────
  useEffect(() => {
    let disposed = false
    void window.slidesApi.presenterStart({ monitor: readShowMonitor() }).then((r) => {
      if (!disposed) setHasAudience(r.audience)
    })
    return () => {
      disposed = true
      void window.slidesApi.presenterEnd()
    }
  }, [])

  // Broadcast show state: epoch change = the new page's animations are loaded and loadedRef matches the player cursor
  useEffect(() => {
    const state: ShowSyncState = {
      idx: loadedRef.current.idx,
      fresh: loadedRef.current.fresh,
      played: player.played,
      playing: player.playing,
      ended,
      black,
      white,
      zoom,
      tool,
      inkColor,
      camera,
    }
    window.slidesApi.presenterSync(state)
  }, [
    player.epoch,
    player.played,
    player.playing,
    ended,
    black,
    white,
    zoom,
    tool,
    inkColor,
    camera,
  ])

  // Clear ink on page turn (the audience side clears in sync)
  useEffect(() => {
    board.clear()
    window.slidesApi.presenterInk({ type: 'clear' })
    setZoom(null)
    setZoomArmed(false)
  }, [pos, board])

  const exitRef = useRef(() => {})
  exitRef.current = () => onExit(order[Math.min(pos, order.length - 1)] ?? startAt)

  // System full screen: requested on entry; kept when switching to normal show (SlideShowView takes over seamlessly)
  const keepFsRef = useRef(false)
  /** Children stay hidden (black root only) until the snap's resize settled — same
   *  windowed-flash guard as SlideShowView's covered gate */
  const [covered, setCovered] = useState(false)
  useEffect(() => {
    // Same cover mechanics as SlideShowView: one main-side call bleeds the view over
    // the tab strip and snaps the window (macOS simpleFullScreen, no Space animation);
    // HTML fullscreen is only used off-macOS where it is instant.
    let alive = true
    const snapped = window.slidesApi.setShowFullScreen?.(true) ?? Promise.resolve()
    void snapped
      .catch(() => {})
      .then(() => {
        if (!IS_MAC) void document.documentElement.requestFullscreen?.().catch(() => {})
        // Same settle condition as SlideShowView: viewport spans the whole
        // screen in BOTH dimensions — no bleed-only or full-width-only
        // intermediate passes. Deadline covers stale preloads that never snap.
        const deadline = performance.now() + 500
        const reveal = () => {
          if (!alive) return
          const w = window.innerWidth
          const h = window.innerHeight
          const settled = w >= screen.width && h >= screen.height
          if (!settled && performance.now() < deadline) {
            requestAnimationFrame(reveal)
            return
          }
          setSize({ w, h })
          setCovered(true)
        }
        requestAnimationFrame(reveal)
      })
    return () => {
      alive = false
      if (!keepFsRef.current) {
        if (document.fullscreenElement) void document.exitFullscreen().catch(() => {})
        void window.slidesApi.setShowFullScreen?.(false)
      }
      liftShowCurtain()
    }
  }, [])

  // Curtain from startPresenterView: needed only until the presenter reveals
  useEffect(() => {
    if (covered) liftShowCurtain()
  }, [covered])

  useEffect(() => {
    const onResize = () => setSize({ w: window.innerWidth, h: window.innerHeight })
    window.addEventListener('resize', onResize)
    return () => window.removeEventListener('resize', onResize)
  }, [])

  const posRef = useRef(pos)
  posRef.current = pos
  const goTo = useCallback((nextPos: number, fresh: boolean) => {
    navModeRef.current = fresh ? 'fresh' : 'all'
    if (nextPos !== posRef.current) setLastViewed(posRef.current)
    setPos(nextPos)
  }, [])

  const next = useCallback(() => {
    if (ended) {
      exitRef.current()
      return
    }
    // Advance in-page animations first; turn the page only when this page's animations are done
    if (player.advance()) return
    if (pos >= order.length - 1) setEnded(true)
    else goTo(pos + 1, true)
  }, [ended, pos, order.length, goTo, player.advance]) // eslint-disable-line react-hooks/exhaustive-deps

  const prev = useCallback(() => {
    if (ended) {
      setEnded(false)
      return
    }
    if (pos > 0) goTo(pos - 1, false)
  }, [ended, pos, goTo])

  // Navigation sent back from the audience window (clicks/keys)
  const nextRef = useRef(next)
  nextRef.current = next
  const prevRef = useRef(prev)
  prevRef.current = prev
  useEffect(
    () =>
      window.slidesApi.onAudienceNav((action) => {
        if (action === 'next') nextRef.current()
        else if (action === 'prev') prevRef.current()
        else if (action === 'exit') exitRef.current()
        else if (action.startsWith('goto:')) gotoRef.current(Number(action.slice(5)))
        else commandRef.current(action as ShowKeyCommand)
      }),
    [],
  )

  // Keyboard: PowerPoint's show shortcuts, shared with the show and audience windows
  const runCommand = useCallback(
    (command: ShowKeyCommand) => {
      if (command === 'exit') {
        // Esc closes an open menu, then leaves Zoom, then ends the show
        if (popoverOpenRef.current) return closeShowPopovers()
        if (zoomArmed || zoom) {
          setZoomArmed(false)
          setZoom(null)
          return
        }
        return exitRef.current()
      }
      if (command === 'black' || command === 'white')
        return setCover((c) => toggleShowScreen(c, command))
      setCover('none')
      if (command === 'next') return next()
      if (command === 'prev') return prev()
      setEnded(false)
      goTo(command === 'first' ? 0 : order.length - 1, false)
    },
    [next, prev, goTo, order.length, zoomArmed, zoom],
  )
  const gotoNumber = useCallback(
    (n: number) => {
      const p = order.indexOf(n - 1)
      if (p < 0) return
      setCover('none')
      setEnded(false)
      goTo(p, false)
    },
    [order, goTo],
  )
  useShowKeys({ onCommand: runCommand, onGoto: gotoNumber })
  const commandRef = useRef(runCommand)
  commandRef.current = runCommand
  const gotoRef = useRef(gotoNumber)
  gotoRef.current = gotoNumber

  // Scroll the current thumbnail into view
  const filmRef = useRef<HTMLDivElement>(null)
  useEffect(() => {
    filmRef.current
      ?.querySelector('.pv-film-cur')
      ?.scrollIntoView({ block: 'nearest', inline: 'nearest', behavior: 'smooth' })
  }, [pos])

  const clearInk = useCallback(() => {
    board.clear()
    window.slidesApi.presenterInk({ type: 'clear' })
  }, [board])

  usePointerShortcuts(setTool, clearInk)

  if (!slide) return null

  // Left main area auto-fit: subtract the right column/top bar/timer row/toolbar/thumbnail strip, then take the max fit by aspect ratio
  const ar = slide.widthPx / slide.heightPx
  const mainW = size.w - sideW
  const mainH = size.h - TOP_H - TIMER_H - BOTTOM_H - FILM_H
  const fitW = Math.max(80, Math.round(Math.min(mainW - 48, (mainH - 16) * ar)))
  const fitH = Math.round(fitW / ar)

  const nextSlide = ended ? null : (slides[order[pos + 1] ?? -1] ?? null)
  const notes = allNotes[order[pos]!] ?? []

  const useShow = () => {
    keepFsRef.current = true
    onUseSlideShow?.(order[Math.min(pos, order.length - 1)] ?? startAt)
  }

  return (
    <div className={`presenter${covered ? '' : ' pv-uncovered'}`}>
      <div className="pv-top">
        <button
          className="pv-top-btn pv-top-exit"
          onClick={() => exitRef.current()}
          data-tip={t('panePresenterEndTip')}
        >
          <span className="pv-top-ico">{ShowIcon.endShow(20)}</span>
          <span>{t('panePresenterEndShow')}</span>
        </button>
        <button
          className="pv-top-btn"
          disabled={!hasAudience}
          onClick={() => void window.slidesApi.presenterSwap()}
          data-tip={hasAudience ? t('panePresenterSwapTip') : t('panePresenterNoSecond')}
        >
          <span className="pv-top-ico">{ShowIcon.swap(20)}</span>
          <span>{t('panePresenterSwap')}</span>
        </button>
        {onUseSlideShow && (
          <button className="pv-top-btn" onClick={useShow} data-tip={t('panePresenterUseShowTip')}>
            <span className="pv-top-ico">{ShowIcon.slideShow(20)}</span>
            <span>{t('panePresenterUseShow')}</span>
          </button>
        )}
        <div className="pv-top-spacer" />
        {!hasAudience && <span className="pv-top-hint">{t('panePresenterSingleHint')}</span>}
      </div>
      <div className="pv-body">
        <div className="pv-left">
          <div className="pv-timer-row">
            <span className="pv-timer" data-tip={t('panePresenterElapsed')}>
              {fmtElapsed(elapsed)}
            </span>
            <button
              className="pv-mini-btn"
              onClick={() => setPaused((v) => !v)}
              data-tip={paused ? t('panePresenterResume') : t('panePresenterPause')}
              aria-label={paused ? t('panePresenterResume') : t('panePresenterPause')}
            >
              {paused ? '▶' : '❚❚'}
            </button>
            <button
              className="pv-mini-btn"
              onClick={() => {
                setElapsed(0)
                setPaused(false)
              }}
              data-tip={t('panePresenterRestart')}
              aria-label={t('panePresenterRestart')}
            >
              ↻
            </button>
            <div className="pv-top-spacer" />
            <span className="pv-clock" data-tip={t('panePresenterClock')}>
              {clock}
            </span>
          </div>
          <div
            className="pv-stage"
            onClick={(e) => {
              if (tool !== 'none') return
              if (zoomArmed) {
                // magnify the clicked spot (PowerPoint's Zoom); a later click zooms back out
                const r = stageboxRef.current?.getBoundingClientRect()
                if (r && r.width && r.height)
                  setZoom({
                    x: Math.min(1, Math.max(0, (e.clientX - r.left) / r.width)),
                    y: Math.min(1, Math.max(0, (e.clientY - r.top) / r.height)),
                  })
                setZoomArmed(false)
                return
              }
              if (zoom) {
                setZoom(null)
                return
              }
              next()
            }}
          >
            {ended ? (
              <div className="pv-end">{t('panePresenterEnded')}</div>
            ) : (
              <div
                ref={stageboxRef}
                className={`pv-stagebox${tool !== 'none' ? ` pv-tool-${tool}` : ''}${zoomArmed ? ' pv-zoom-armed' : ''}`}
                style={{ width: fitW, height: fitH }}
                {...ink}
              >
                <div
                  className="show-zoom"
                  style={
                    zoom
                      ? {
                          transform: 'scale(2)',
                          transformOrigin: `${zoom.x * 100}% ${zoom.y * 100}%`,
                        }
                      : undefined
                  }
                >
                  <AnimatedSlideStage
                    slide={slide}
                    images={images}
                    width={fitW}
                    states={player.states}
                  />
                </div>
                <CameraBubble on={camera} width={fitW} onError={() => setCamera(false)} />
                <InkLayer board={board} width={fitW} height={fitH} />
                {black && <div className="pv-black" data-tip={t('panePresenterBlackOn')} />}
                {white && <div className="pv-white" data-tip={t('panePresenterBlackOn')} />}
              </div>
            )}
          </div>
          <div className="pv-bottomrow">
            <ShowToolbar
              tool={tool}
              inkColor={inkColor}
              canErase={canErase}
              zoomOn={zoomOn}
              black={black}
              cameraOn={camera}
              onCamera={() => setCamera((v) => !v)}
              menuItems={showOptionsMenu(t, {
                slides,
                order,
                pos,
                ended,
                lastViewed,
                cover,
                tool,
                canErase,
                presenter: true,
                canSwap: hasAudience,
                paused,
                onNext: next,
                onPrev: prev,
                onGoto: (p) => {
                  setCover('none')
                  setEnded(false)
                  goTo(p, false)
                },
                onCover: (c) => setCover((cur) => toggleShowScreen(cur, c)),
                onSwap: () => void window.slidesApi.presenterSwap(),
                ...(onUseSlideShow ? { onToggleView: useShow } : {}),
                onTool: setTool,
                onEraseAll: clearInk,
                onPause: () => setPaused((v) => !v),
                onEnd: () => exitRef.current(),
              })}
              onTool={setTool}
              onColor={pickColor}
              onEraseAll={clearInk}
              onZoom={toggleZoom}
              onBlack={() => setCover((c) => toggleShowScreen(c, 'black'))}
              onOpenChange={(open) => (popoverOpenRef.current = open)}
              // the tool row sits above the thumbnail strip: open the menu upward
              menuUp
            />
            <div className="pv-nav">
              <button
                className="pv-round"
                disabled={!ended && pos === 0}
                onClick={prev}
                data-tip={t('panePresenterPrevTip')}
                aria-label={t('panePresenterPrevTip')}
              >
                ‹
              </button>
              <div className="pv-nav-mid">
                <div className="pv-nav-label">
                  {t('panePresenterSlideOf', { cur: pos + 1, total: order.length })}
                </div>
                <div className="pv-progress">
                  <div style={{ width: `${Math.round(((pos + 1) / order.length) * 100)}%` }} />
                </div>
              </div>
              <button
                className="pv-round"
                onClick={next}
                data-tip={t('panePresenterNextTip')}
                aria-label={t('panePresenterNextTip')}
              >
                ›
              </button>
            </div>
            <div aria-hidden />
          </div>
        </div>
        <div
          className="pv-splitter"
          role="separator"
          aria-orientation="vertical"
          onPointerDown={startSideDrag}
        />
        <div className="pv-side" style={{ width: sideW }}>
          <div className="pv-section-label">{t('panePresenterNextSlide')}</div>
          <div className="pv-next">
            {nextSlide ? (
              <SlideThumb slide={nextSlide} images={images} width={sideW - 28} />
            ) : (
              <div className="pv-next-none">{ended ? '—' : t('panePresenterLastSlide')}</div>
            )}
          </div>
          <div className="pv-section-label">{t('panePresenterNotes')}</div>
          <div className="pv-notes" style={{ fontSize: noteSize }}>
            {notes.some((p) => p.runs.some((r) => r.text.trim())) ? (
              <NotesView paragraphs={notes} />
            ) : (
              t('panePresenterNoNotes')
            )}
          </div>
          <div className="pv-notes-size">
            <button
              className="pv-mini-btn"
              onClick={() => setNoteSize((s) => Math.min(28, s + 2))}
              data-tip={t('panePresenterNotesBigger')}
            >
              A⁺
            </button>
            <button
              className="pv-mini-btn"
              onClick={() => setNoteSize((s) => Math.max(11, s - 2))}
              data-tip={t('panePresenterNotesSmaller')}
            >
              A⁻
            </button>
          </div>
        </div>
      </div>
      <div className="pv-film" ref={filmRef}>
        {order.map((idx, i) => (
          <div
            key={idx}
            className={`pv-film-item${i === pos ? ' pv-film-cur' : ''}`}
            onClick={() => {
              setEnded(false)
              goTo(i, false)
            }}
            data-tip={t('panePresenterFilmSlideN', { n: i + 1 })}
          >
            <SlideThumb slide={slides[idx]!} images={images} width={150} />
          </div>
        ))}
      </div>
    </div>
  )
}
