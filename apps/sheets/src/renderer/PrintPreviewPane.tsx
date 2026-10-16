import { useEffect, useLayoutEffect, useRef, useState } from 'react'
import type { PDFDocumentProxy } from 'pdfjs-dist'

import { useI18n } from './i18n/locale'
import type { PrintTargets } from './page-layout-actions'
import {
  PREVIEW_LOOKAHEAD,
  applyFetch,
  clampPage,
  createDebouncedLatest,
  emptyTimeline,
  isTimelineComplete,
  knownPages,
  locatePage,
  nextFetch,
  needsMore,
  pageCountLabel,
  type PreviewTimeline,
} from './print-preview-plan'
import { drawPreviewPage, openPreviewPdf } from './print-preview-render'

/** Settings changes wait this long before the preview is rebuilt. */
const PREVIEW_DEBOUNCE_MS = 350
/** Pages loaded in the background after the first segment, so the page count settles. */
const BACKGROUND_PAGES = 40

interface Segment {
  readonly unit: number
  /** First page of this segment within its unit (1-based). */
  readonly start: number
  readonly count: number
  readonly doc: PDFDocumentProxy
}

interface Session {
  readonly id: string
  readonly targets: Extract<PrintTargets, { ok: true }>
  readonly segments: Segment[]
  timeline: PreviewTimeline
  payload: {
    unit: number
    request: Awaited<ReturnType<Extract<PrintTargets, { ok: true }>['payloadAt']>>
  } | null
  lock: Promise<void>
  disposed: boolean
}

let sessionCounter = 0

function disposeSession(session: Session | null): void {
  if (!session) return
  session.disposed = true
  for (const segment of session.segments)
    void segment.doc.loadingTask.destroy().catch(() => undefined)
  session.segments.length = 0
  void window.desktopApi.releasePrintPreview(session.id).catch(() => undefined)
}

/**
 * The print sheet's preview: the same laid-out HTML and page setup the real print uses, rendered
 * to PDF by the main process in short segments and drawn page by page. `jobKey` changes whenever
 * something that alters the printout changes; the preview then rebuilds after a short pause.
 */
export function PrintPreviewPane({
  getTargets,
  jobKey,
}: {
  readonly getTargets: () => PrintTargets
  readonly jobKey: string
}): React.JSX.Element {
  const { t } = useI18n()
  const [timeline, setTimeline] = useState<PreviewTimeline>(() => emptyTimeline(0))
  const [page, setPage] = useState(1)
  const [phase, setPhase] = useState<'loading' | 'ready' | 'error'>('loading')
  const [error, setError] = useState('')
  const [fetching, setFetching] = useState(false)
  const [size, setSize] = useState({ width: 0, height: 0 })
  const sessionRef = useRef<Session | null>(null)
  const targetsRef = useRef(getTargets)
  targetsRef.current = getTargets
  const frameRef = useRef<HTMLDivElement | null>(null)
  const canvasRef = useRef<HTMLCanvasElement | null>(null)

  const fetchNext = (session: Session): Promise<void> => {
    const step = async (): Promise<void> => {
      if (session.disposed) return
      const next = nextFetch(session.timeline)
      if (!next) return
      if (session.payload?.unit !== next.unit) {
        session.payload = { unit: next.unit, request: await session.targets.payloadAt(next.unit) }
      }
      if (session.disposed) return
      const id = `${session.id}:${next.unit}`
      const request = session.payload.request
      let result = await window.desktopApi.printPreview(
        id,
        next.from === 1 ? request : undefined,
        next.from,
        next.to,
      )
      if (!result.ok && result.expired === true) {
        result = await window.desktopApi.printPreview(id, request, next.from, next.to)
      }
      if (!result.ok) throw new Error(result.error ?? t('appPrintPreviewFailed'))
      let returned = 0
      if (result.pdf !== null) {
        const doc = await openPreviewPdf(result.pdf)
        if (session.disposed) {
          void doc.loadingTask.destroy().catch(() => undefined)
          return
        }
        returned = doc.numPages
        if (returned > 0) {
          session.segments.push({
            unit: next.unit,
            start: result.full ? 1 : next.from,
            count: returned,
            doc,
          })
        } else {
          void doc.loadingTask.destroy().catch(() => undefined)
        }
      }
      session.timeline = applyFetch(session.timeline, next, returned, result.full)
      setTimeline(session.timeline)
    }
    session.lock = session.lock.then(step)
    // later steps keep running after one failed; the failure itself reaches the caller
    const mine = session.lock
    session.lock = mine.catch(() => undefined)
    return mine
  }

  /** Loads segments until `targetPage` exists or everything is loaded. */
  const ensurePage = async (session: Session, targetPage: number): Promise<void> => {
    while (
      !session.disposed &&
      !isTimelineComplete(session.timeline) &&
      knownPages(session.timeline) < targetPage
    ) {
      await fetchNext(session)
    }
  }

  const schedulerRef = useRef<ReturnType<typeof createDebouncedLatest> | null>(null)
  schedulerRef.current ??= createDebouncedLatest(PREVIEW_DEBOUNCE_MS, async (isStale) => {
    disposeSession(sessionRef.current)
    sessionRef.current = null
    setPhase('loading')
    setError('')
    const targets = targetsRef.current()
    if (!targets.ok) {
      setTimeline(emptyTimeline(0))
      setError(targets.error)
      setPhase('error')
      return
    }
    sessionRef.current = null
    const session: Session = {
      id: `print-preview-${Date.now()}-${++sessionCounter}`,
      targets,
      segments: [],
      timeline: emptyTimeline(targets.count),
      payload: null,
      lock: Promise.resolve(),
      disposed: false,
    }
    sessionRef.current = session
    try {
      await ensurePage(session, 1)
      if (isStale() || session.disposed) return
      setTimeline(session.timeline)
      setPage(1)
      setPhase('ready')
      // let the page count settle without making the reader wait for it
      while (
        !isStale() &&
        !session.disposed &&
        !isTimelineComplete(session.timeline) &&
        knownPages(session.timeline) < BACKGROUND_PAGES
      ) {
        await fetchNext(session)
      }
    } catch (failure: unknown) {
      if (isStale() || session.disposed) return
      setError(failure instanceof Error ? failure.message : t('appPrintPreviewFailed'))
      setPhase('error')
    }
  })

  useEffect(() => {
    schedulerRef.current?.schedule()
  }, [jobKey])
  useEffect(
    () => () => {
      schedulerRef.current?.cancel()
      disposeSession(sessionRef.current)
      sessionRef.current = null
    },
    [],
  )

  useLayoutEffect(() => {
    const frame = frameRef.current
    if (!frame) return
    const measure = (): void => setSize({ width: frame.clientWidth, height: frame.clientHeight })
    measure()
    const observer = new ResizeObserver(measure)
    observer.observe(frame)
    return () => observer.disconnect()
  }, [])

  // draw the viewed page
  useEffect(() => {
    const session = sessionRef.current
    const canvas = canvasRef.current
    if (!session || !canvas || size.width < 40 || size.height < 40) return
    const where = locatePage(timeline, page)
    const segment = where
      ? session.segments.find(
          (s) => s.unit === where.unit && where.page >= s.start && where.page < s.start + s.count,
        )
      : undefined
    if (!where || !segment) return
    const draw = drawPreviewPage(
      segment.doc,
      where.page - segment.start + 1,
      canvas,
      size.width - 32,
      size.height - 32,
    )
    draw.done.catch(() => undefined)
    return () => draw.cancel()
  }, [timeline, page, size])

  // paging forward past what is loaded keeps fetching
  useEffect(() => {
    const session = sessionRef.current
    if (!session || phase !== 'ready' || !needsMore(timeline, page, PREVIEW_LOOKAHEAD)) return
    setFetching(true)
    void ensurePage(session, page + PREVIEW_LOOKAHEAD)
      .catch((failure: unknown) => {
        setError(failure instanceof Error ? failure.message : t('appPrintPreviewFailed'))
        setPhase('error')
      })
      .finally(() => setFetching(false))
    // ensurePage is rebuilt every render; its inputs are the timeline and the page
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [timeline, page, phase])

  const known = knownPages(timeline)
  const complete = isTimelineComplete(timeline)
  const current = clampPage(page, timeline)
  const goTo = (target: number): void => {
    const session = sessionRef.current
    if (!session) return
    if (target <= known || complete) {
      setPage(clampPage(target, timeline))
      return
    }
    setFetching(true)
    void ensurePage(session, target)
      .then(() => setPage(clampPage(target, session.timeline)))
      .catch(() => undefined)
      .finally(() => setFetching(false))
  }
  const empty = phase === 'ready' && complete && known === 0
  const loading = phase === 'loading' || fetching

  return (
    <div
      className="print-preview-pane"
      aria-label={t('appPrintPreviewLabel')}
      data-testid="print-preview"
    >
      <div className="print-preview-frame" ref={frameRef}>
        {phase === 'error' ? (
          <div className="print-preview-msg" role="alert">
            {error}
          </div>
        ) : empty ? (
          <div className="print-preview-msg">{t('appPrintPreviewEmpty')}</div>
        ) : (
          <canvas
            ref={canvasRef}
            className={`print-preview-canvas${phase === 'loading' ? ' stale' : ''}`}
            data-testid="print-preview-canvas"
          />
        )}
        {loading && phase !== 'error' && (
          <div className="print-preview-busy" role="status">
            {t('appPrintPreviewLoading')}
          </div>
        )}
      </div>
      <div className="print-preview-nav">
        <button
          type="button"
          aria-label={t('appPrintPrevPage')}
          disabled={phase !== 'ready' || current <= 1}
          onClick={() => goTo(current - 1)}
        >
          ‹
        </button>
        <span data-testid="print-preview-pages">
          {phase === 'ready' && known > 0 ? `${current} / ${pageCountLabel(timeline)}` : '–'}
        </span>
        <button
          type="button"
          aria-label={t('appPrintNextPage')}
          disabled={phase !== 'ready' || (complete && current >= known)}
          onClick={() => goTo(current + 1)}
        >
          ›
        </button>
        {phase === 'ready' && known > 0 && (
          <span className="print-preview-total">
            {t('appPrintPreviewPages', { count: pageCountLabel(timeline) })}
          </span>
        )}
      </div>
    </div>
  )
}
