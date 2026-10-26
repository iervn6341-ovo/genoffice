/**
 * Pure helpers behind the slide show's video control bar (PowerPoint shows a play/pause button,
 * a seekable timeline, elapsed/total time and a volume control over a video while the pointer
 * moves on it, and hides them again a moment after the pointer rests).
 */

/** The bar hides itself this long after the pointer last moved over the video. */
export const VIDEO_CONTROLS_IDLE_MS = 2500

/** Format seconds as m:ss (h:mm:ss from one hour on); unknown / negative / NaN values read as 0:00. */
export function formatMediaTime(sec: number): string {
  const total = Number.isFinite(sec) && sec > 0 ? Math.floor(sec) : 0
  const h = Math.floor(total / 3600)
  const m = Math.floor((total % 3600) / 60)
  const s = total % 60
  const ss = String(s).padStart(2, '0')
  if (h > 0) return `${h}:${String(m).padStart(2, '0')}:${ss}`
  return `${m}:${ss}`
}

/** A playback position clamped to 0..duration (duration unknown → 0). */
export function clampMediaTime(sec: number, duration: number): number {
  const d = Number.isFinite(duration) && duration > 0 ? duration : 0
  if (!Number.isFinite(sec) || sec < 0) return 0
  return Math.min(sec, d)
}

/** Fraction 0..1 of `value` within `duration` (unknown / zero duration → 0). */
export function mediaFraction(value: number, duration: number): number {
  if (!Number.isFinite(duration) || duration <= 0) return 0
  return Math.min(1, Math.max(0, value / duration))
}

/** The minimal slice of TimeRanges used to find how far the video is buffered past `time`. */
export interface RangesLike {
  length: number
  start(i: number): number
  end(i: number): number
}

/** End (seconds) of the buffered range containing `time`; nothing buffered there → `time`. */
export function bufferedEnd(ranges: RangesLike | null | undefined, time: number): number {
  if (!ranges) return time
  for (let i = 0; i < ranges.length; i++) {
    if (time >= ranges.start(i) - 0.05 && time <= ranges.end(i)) return ranges.end(i)
  }
  return time
}

/** Whether the bar is shown: only while the pointer is active over the video (or a drag is running). */
export function controlsVisible(s: {
  hovering: boolean
  idle: boolean
  dragging: boolean
}): boolean {
  return s.dragging || (s.hovering && !s.idle)
}

/** Volume glyph: muted (or zero) / low / high. */
export function volumeLevel(volume: number, muted: boolean): 'mute' | 'low' | 'high' {
  if (muted || !(volume > 0)) return 'mute'
  return volume < 0.5 ? 'low' : 'high'
}

/** The mute button's result: unmuting from a 0 volume restores an audible level. */
export function toggleMute(volume: number, muted: boolean): { volume: number; muted: boolean } {
  if (muted) return { volume: volume > 0 ? volume : 0.5, muted: false }
  return { volume, muted: true }
}

/** Dragging the volume slider: 0 means muted, anything above un-mutes. */
export function applyVolumeInput(value: number): { volume: number; muted: boolean } {
  const v = Math.min(1, Math.max(0, Number.isFinite(value) ? value : 0))
  return { volume: v, muted: v === 0 }
}

export interface VideoBarLayout {
  /** bar height in px */
  height: number
  /** show the elapsed / total readout */
  showTime: boolean
  /** show the volume slider next to the mute button */
  showVolumeSlider: boolean
  /** show the mute button at all */
  showVolume: boolean
  /** the box is too small to host a bar */
  hidden: boolean
}

/**
 * Size the bar for a video box of boxW x boxH px: it must always fit inside the box (so inside
 * the slide), shedding the volume slider and the time readout before the timeline itself.
 */
export function videoBarLayout(boxW: number, boxH: number): VideoBarLayout {
  const w = Math.max(0, boxW)
  const h = Math.max(0, boxH)
  if (w < 90 || h < 44) {
    return { height: 0, showTime: false, showVolumeSlider: false, showVolume: false, hidden: true }
  }
  const height = Math.round(Math.min(40, Math.max(28, h * 0.16)))
  return {
    height,
    showTime: w >= 260,
    showVolumeSlider: w >= 380,
    showVolume: w >= 170,
    hidden: false,
  }
}
