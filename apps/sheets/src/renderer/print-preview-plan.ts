/**
 * Pure planning for the print sheet's preview. The preview shows what will print, one print
 * "unit" (a laid-out sheet) at a time, and fetches PDF pages in segments so a huge sheet never
 * blocks the UI: the first segment appears quickly and further ones load as the reader pages
 * forward. Nothing here touches the DOM or Electron.
 */

/** Pages rendered per request. */
export const PREVIEW_SEGMENT_PAGES = 8
/** Fetch more once the viewed page is this close to the end of what is loaded. */
export const PREVIEW_LOOKAHEAD = 2

export interface PreviewUnitState {
  /** Pages loaded so far (all of them once `complete`). */
  readonly pages: number
  readonly complete: boolean
}

export interface PreviewTimeline {
  /** How many print units (sheets) the job has. */
  readonly unitCount: number
  /** Units fetched so far, in order. */
  readonly units: readonly PreviewUnitState[]
}

export interface PreviewFetch {
  readonly unit: number
  /** First page of the segment, 1-based within the unit. */
  readonly from: number
  /** Last page of the segment (inclusive). */
  readonly to: number
}

export function emptyTimeline(unitCount: number): PreviewTimeline {
  return { unitCount, units: [] }
}

export function knownPages(timeline: PreviewTimeline): number {
  return timeline.units.reduce((sum, unit) => sum + unit.pages, 0)
}

export function isTimelineComplete(timeline: PreviewTimeline): boolean {
  if (timeline.unitCount === 0) return true
  const last = timeline.units[timeline.units.length - 1]
  return timeline.units.length === timeline.unitCount && last !== undefined && last.complete
}

/** The next segment to request, or null when everything is loaded. */
export function nextFetch(
  timeline: PreviewTimeline,
  segment: number = PREVIEW_SEGMENT_PAGES,
): PreviewFetch | null {
  if (timeline.unitCount === 0) return null
  const last = timeline.units[timeline.units.length - 1]
  if (last === undefined) return { unit: 0, from: 1, to: segment }
  if (!last.complete) {
    return { unit: timeline.units.length - 1, from: last.pages + 1, to: last.pages + segment }
  }
  if (timeline.units.length < timeline.unitCount) {
    return { unit: timeline.units.length, from: 1, to: segment }
  }
  return null
}

/**
 * Records a fetched segment. A segment that came back shorter than asked ends its unit; `full`
 * marks a result that already holds the whole unit (header/footer variants cannot be paged).
 */
export function applyFetch(
  timeline: PreviewTimeline,
  fetch: PreviewFetch,
  returnedPages: number,
  full = false,
): PreviewTimeline {
  const requested = fetch.to - fetch.from + 1
  const units = timeline.units.slice()
  const previous = units[fetch.unit]
  const base = previous?.pages ?? 0
  const pages = full ? returnedPages : base + returnedPages
  units[fetch.unit] = { pages, complete: full || returnedPages < requested }
  return { unitCount: timeline.unitCount, units }
}

/** Where a 1-based global page lives, or null while it is not loaded. */
export function locatePage(
  timeline: PreviewTimeline,
  globalPage: number,
): { readonly unit: number; readonly page: number } | null {
  if (globalPage < 1) return null
  let remaining = globalPage
  for (let unit = 0; unit < timeline.units.length; unit++) {
    const pages = timeline.units[unit]!.pages
    if (remaining <= pages) return { unit, page: remaining }
    remaining -= pages
  }
  return null
}

/** True when the viewed page is near the end of the loaded pages and more can be fetched. */
export function needsMore(
  timeline: PreviewTimeline,
  currentPage: number,
  lookahead: number = PREVIEW_LOOKAHEAD,
): boolean {
  return !isTimelineComplete(timeline) && currentPage + lookahead > knownPages(timeline)
}

/** "N" once everything is loaded, "N+" while more may follow. */
export function pageCountLabel(timeline: PreviewTimeline): string {
  const known = knownPages(timeline)
  return isTimelineComplete(timeline) ? String(known) : `${known}+`
}

export function clampPage(page: number, timeline: PreviewTimeline): number {
  const known = knownPages(timeline)
  if (known < 1) return 1
  return Math.min(Math.max(1, Math.floor(page) || 1), known)
}

/** Print-job key: when it changes the preview is stale and has to be rebuilt. */
export function previewKey(parts: unknown): string {
  return JSON.stringify(parts)
}

export interface DebouncedLatest {
  /** (Re)starts the countdown; an earlier pending or running run becomes stale. */
  schedule(): void
  /** Drops anything pending or running. */
  cancel(): void
}

/**
 * Debounced "latest wins" runner: after `delayMs` of quiet it calls `run`, handing it an
 * `isStale` probe that turns true as soon as a newer schedule() or cancel() happens, so the run
 * can stop early and discard its result.
 */
export function createDebouncedLatest(
  delayMs: number,
  run: (isStale: () => boolean) => Promise<void>,
  timers: {
    readonly set: (fn: () => void, ms: number) => unknown
    readonly clear: (handle: unknown) => void
  } = {
    set: (fn, ms) => setTimeout(fn, ms),
    clear: (handle) => clearTimeout(handle as ReturnType<typeof setTimeout>),
  },
): DebouncedLatest {
  let generation = 0
  let handle: unknown = null
  const drop = (): void => {
    generation++
    if (handle !== null) timers.clear(handle)
    handle = null
  }
  return {
    schedule() {
      drop()
      const mine = generation
      handle = timers.set(() => {
        handle = null
        void run(() => mine !== generation).catch(() => undefined)
      }, delayMs)
    },
    cancel: drop,
  }
}
