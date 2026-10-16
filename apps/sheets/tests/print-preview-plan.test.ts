import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'

import {
  PREVIEW_SEGMENT_PAGES,
  applyFetch,
  clampPage,
  createDebouncedLatest,
  emptyTimeline,
  isTimelineComplete,
  knownPages,
  locatePage,
  needsMore,
  nextFetch,
  pageCountLabel,
  previewKey,
} from '../src/renderer/print-preview-plan'

const SEG = PREVIEW_SEGMENT_PAGES

describe('preview timeline', () => {
  it('starts with the first segment of the first unit', () => {
    expect(nextFetch(emptyTimeline(2))).toEqual({ unit: 0, from: 1, to: SEG })
    expect(nextFetch(emptyTimeline(0))).toBeNull()
    expect(isTimelineComplete(emptyTimeline(0))).toBe(true)
    expect(isTimelineComplete(emptyTimeline(1))).toBe(false)
  })

  it('a short segment ends its unit and moves on to the next', () => {
    let timeline = emptyTimeline(2)
    timeline = applyFetch(timeline, nextFetch(timeline)!, 3)
    expect(timeline.units[0]).toEqual({ pages: 3, complete: true })
    expect(nextFetch(timeline)).toEqual({ unit: 1, from: 1, to: SEG })
    timeline = applyFetch(timeline, nextFetch(timeline)!, 2)
    expect(nextFetch(timeline)).toBeNull()
    expect(isTimelineComplete(timeline)).toBe(true)
    expect(knownPages(timeline)).toBe(5)
    expect(pageCountLabel(timeline)).toBe('5')
  })

  it('a full segment asks for the next one of the same unit', () => {
    let timeline = emptyTimeline(1)
    timeline = applyFetch(timeline, nextFetch(timeline)!, SEG)
    expect(pageCountLabel(timeline)).toBe(`${SEG}+`)
    expect(nextFetch(timeline)).toEqual({ unit: 0, from: SEG + 1, to: SEG * 2 })
    // exactly SEG pages in total: the follow-up segment comes back empty
    timeline = applyFetch(timeline, nextFetch(timeline)!, 0)
    expect(timeline.units[0]).toEqual({ pages: SEG, complete: true })
    expect(isTimelineComplete(timeline)).toBe(true)
  })

  it('a full render (header/footer variants) completes the unit at once', () => {
    const timeline = applyFetch(emptyTimeline(1), nextFetch(emptyTimeline(1))!, 30, true)
    expect(timeline.units[0]).toEqual({ pages: 30, complete: true })
    expect(nextFetch(timeline)).toBeNull()
  })

  it('maps global pages onto units', () => {
    let timeline = emptyTimeline(2)
    timeline = applyFetch(timeline, nextFetch(timeline)!, 3)
    timeline = applyFetch(timeline, nextFetch(timeline)!, 4)
    expect(locatePage(timeline, 1)).toEqual({ unit: 0, page: 1 })
    expect(locatePage(timeline, 3)).toEqual({ unit: 0, page: 3 })
    expect(locatePage(timeline, 4)).toEqual({ unit: 1, page: 1 })
    expect(locatePage(timeline, 7)).toEqual({ unit: 1, page: 4 })
    expect(locatePage(timeline, 8)).toBeNull()
    expect(locatePage(timeline, 0)).toBeNull()
  })

  it('asks for more only near the end of an incomplete timeline', () => {
    let timeline = emptyTimeline(1)
    timeline = applyFetch(timeline, nextFetch(timeline)!, SEG)
    expect(needsMore(timeline, 1)).toBe(false)
    expect(needsMore(timeline, SEG - 1)).toBe(true)
    expect(needsMore(timeline, SEG)).toBe(true)
    timeline = applyFetch(timeline, nextFetch(timeline)!, 1)
    expect(needsMore(timeline, SEG + 1)).toBe(false)
  })

  it('clamps the viewed page into the loaded range', () => {
    const timeline = applyFetch(emptyTimeline(1), { unit: 0, from: 1, to: SEG }, 3)
    expect(clampPage(0, timeline)).toBe(1)
    expect(clampPage(99, timeline)).toBe(3)
    expect(clampPage(2, timeline)).toBe(2)
    expect(clampPage(5, emptyTimeline(1))).toBe(1)
  })

  it('keys the preview by the settings that change the printout', () => {
    expect(previewKey({ a: 1, b: [2] })).toBe(previewKey({ a: 1, b: [2] }))
    expect(previewKey({ a: 1 })).not.toBe(previewKey({ a: 2 }))
  })
})

describe('createDebouncedLatest', () => {
  beforeEach(() => vi.useFakeTimers())
  afterEach(() => vi.useRealTimers())

  it('runs once after a quiet period', async () => {
    const run = vi.fn(async () => undefined)
    const job = createDebouncedLatest(300, run)
    job.schedule()
    job.schedule()
    job.schedule()
    vi.advanceTimersByTime(299)
    expect(run).not.toHaveBeenCalled()
    vi.advanceTimersByTime(1)
    expect(run).toHaveBeenCalledTimes(1)
  })

  it('marks a running render stale when a newer one is scheduled', async () => {
    const stale: boolean[] = []
    let release: () => void = () => undefined
    const job = createDebouncedLatest(100, async (isStale) => {
      await new Promise<void>((resolve) => (release = resolve))
      stale.push(isStale())
    })
    job.schedule()
    vi.advanceTimersByTime(100)
    job.schedule() // settings changed while the first render is in flight
    release()
    await vi.advanceTimersByTimeAsync(0)
    expect(stale).toEqual([true])
  })

  it('cancel drops a pending run and stales a running one', async () => {
    const run = vi.fn(async () => undefined)
    const job = createDebouncedLatest(100, run)
    job.schedule()
    job.cancel()
    vi.advanceTimersByTime(500)
    expect(run).not.toHaveBeenCalled()
  })

  it('swallows a failing run instead of leaking an unhandled rejection', async () => {
    const job = createDebouncedLatest(10, async () => {
      throw new Error('boom')
    })
    job.schedule()
    await vi.advanceTimersByTimeAsync(20)
  })
})
