import { useCallback, useEffect, useMemo, useRef, useState } from 'react'
import type { AnimationItem } from '../shared/ipc'
import { buildSteps, computeNodeStates, type AnimStep, type NodeAnimState } from './animation-play'

export interface AnimPlayer {
  /** Animation metadata is loaded for the page being rendered. */
  ready: boolean
  /** animStateKey (sourceId or sourceId+paragraph index) → current visual state (targets without animation aren't in the Map) */
  states: Map<string, NodeAnimState>
  /** Whether the current step is playing */
  playing: boolean
  /** Whether there are steps left to play */
  pending: boolean
  /** Number of steps played (broadcast by the presenter to the audience window) */
  played: number
  /** Generation number of load/seek: incremented after a page loads; broadcasters use it to tell "the new page is in place" */
  epoch: number
  /** Advance: playing → complete the current step immediately; otherwise play the next step. Returns false when nothing to play */
  advance: () => boolean
  /** Load a page's animations; 'fresh' = initial state (auto steps auto-play), 'all' = everything-played state */
  load: (items: AnimationItem[], mode: 'fresh' | 'all') => void
  /** Jump to an absolute cursor (audience window mirroring the presenter); playing=true plays the current step pointed by played from its start */
  seek: (items: AnimationItem[], played: number, playing: boolean) => void
}

export interface AnimPageSource {
  /** A new visit resets playback even when returning to the same slide. */
  key: string | number
  items: AnimationItem[] | null
  /** Audience playback follows the presenter's absolute cursor. */
  cursor?: { played: number; playing: boolean }
}

interface PlayerState {
  steps: AnimStep[]
  played: number
  startedAt: number | null
  epoch: number
  source?: AnimPageSource
}

function pageState(source: AnimPageSource, epoch: number): PlayerState {
  const steps = buildSteps(source.items ?? [])
  const played = source.cursor ? Math.max(0, Math.min(source.cursor.played, steps.length)) : 0
  const playing = source.cursor?.playing ?? !!steps[0]?.auto
  const instant = playing && played < steps.length && steps[played]!.totalMs <= 0
  return {
    source,
    steps,
    played: instant ? played + 1 : played,
    startedAt: playing && played < steps.length && !instant ? performance.now() : null,
    epoch,
  }
}

export function useAnimPlayer(
  canvasHpx: number,
  canvasWpx?: number,
  source?: AnimPageSource,
): AnimPlayer {
  const [state, setState] = useState<PlayerState>(() =>
    source ? pageState(source, 0) : { steps: [], played: 0, startedAt: null, epoch: 0 },
  )
  const [elapsed, setElapsed] = useState(0)
  // Reset during this component's render, before children can commit a new slide
  // with the preceding page's cursor (or an empty map exposing every entrance).
  // React retries this render immediately; no mismatched frame reaches the canvas.
  if (
    source &&
    (source.key !== state.source?.key ||
      source.items !== state.source?.items ||
      source.cursor?.played !== state.source?.cursor?.played ||
      source.cursor?.playing !== state.source?.cursor?.playing)
  ) {
    setState(pageState(source, state.epoch + 1))
    setElapsed(0)
  }
  const stateRef = useRef(state)
  stateRef.current = state

  // rAF drives the current step's progress; when finished, stops at step end automatically (awaiting the next trigger)
  useEffect(() => {
    if (state.startedAt == null) return
    let raf = 0
    const tick = () => {
      if (stateRef.current !== state) return
      const t = performance.now() - state.startedAt!
      const total = state.steps[state.played]?.totalMs ?? 0
      if (t >= total) {
        setState((s) => ({ ...s, played: s.played + 1, startedAt: null }))
      } else {
        setElapsed(t)
        raf = requestAnimationFrame(tick)
      }
    }
    raf = requestAnimationFrame(tick)
    return () => cancelAnimationFrame(raf)
  }, [state])

  const load = useCallback((items: AnimationItem[], mode: 'fresh' | 'all') => {
    const steps = buildSteps(items)
    setElapsed(0)
    const s = stateRef.current
    const next =
      mode === 'all'
        ? { steps, played: steps.length, startedAt: null, epoch: s.epoch + 1 }
        : {
            steps,
            played: steps[0]?.auto && steps[0].totalMs <= 0 ? 1 : 0,
            startedAt: steps[0]?.auto && steps[0].totalMs > 0 ? performance.now() : null,
            epoch: s.epoch + 1,
          }
    stateRef.current = next
    setState(next)
  }, [])

  const seek = useCallback((items: AnimationItem[], played: number, playing: boolean) => {
    const steps = buildSteps(items)
    setElapsed(0)
    const s = stateRef.current
    const p = Math.max(0, Math.min(played, steps.length))
    const instant = playing && p < steps.length && steps[p]!.totalMs <= 0
    const next = {
      steps,
      played: instant ? p + 1 : p,
      startedAt: playing && p < steps.length && !instant ? performance.now() : null,
      epoch: s.epoch + 1,
    }
    stateRef.current = next
    setState(next)
  }, [])

  const advance = useCallback((): boolean => {
    const s = stateRef.current
    if (s.startedAt != null) {
      // Playing: complete the current step immediately
      const next = { ...s, played: s.played + 1, startedAt: null }
      stateRef.current = next
      setState(next)
      return true
    }
    if (s.played < s.steps.length) {
      setElapsed(0)
      // Appear/Disappear must commit immediately: a background/occluded presenter
      // may not receive an animation frame while the audience has already moved on.
      const instant = s.steps[s.played]!.totalMs <= 0
      const next = {
        ...s,
        played: instant ? s.played + 1 : s.played,
        startedAt: instant ? null : performance.now(),
      }
      stateRef.current = next
      setState(next)
      return true
    }
    return false
  }, [])

  const states = useMemo(
    () =>
      computeNodeStates(
        state.steps,
        state.played,
        state.startedAt != null ? elapsed : null,
        canvasHpx,
        canvasWpx ?? (canvasHpx * 16) / 9,
      ),
    [state, elapsed, canvasHpx, canvasWpx],
  )

  return {
    ready: !source || source.items != null,
    states,
    playing: state.startedAt != null,
    pending: state.played < state.steps.length,
    played: state.played,
    epoch: state.epoch,
    advance,
    load,
    seek,
  }
}
