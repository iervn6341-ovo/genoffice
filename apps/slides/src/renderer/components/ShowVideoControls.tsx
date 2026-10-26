/**
 * Slide-show video player: the <video> plus PowerPoint's playback bar (play/pause, a seekable
 * timeline with buffered progress, elapsed / total time, mute + volume). The bar floats over the
 * bottom of the video while the pointer moves on it and hides after it rests; clicking the video
 * itself plays / pauses (it never turns the page); Space does the same while the pointer is on
 * the video, otherwise Space keeps turning pages.
 */
import React, { useCallback, useEffect, useRef, useState } from 'react'
import { useI18n } from '../i18n/locale'
import { registerShowVideoSpaceClaimer } from '../show-video-space'
import {
  VIDEO_CONTROLS_IDLE_MS,
  applyVolumeInput,
  bufferedEnd,
  controlsVisible,
  formatMediaTime,
  mediaFraction,
  toggleMute,
  videoBarLayout,
  volumeLevel,
} from '../show-video-state'
import '../show-video-controls.css'

// The last volume the presenter chose carries over to the next video of the show.
let rememberedVolume = { volume: 1, muted: false }

const Icon = ({ children }: { children: React.ReactNode }) => (
  <svg
    width="18"
    height="18"
    viewBox="0 0 24 24"
    fill="currentColor"
    stroke="none"
    aria-hidden
    focusable="false"
  >
    {children}
  </svg>
)

const PlayIcon = () => (
  <Icon>
    <path d="M7 4.5v15l12-7.5z" />
  </Icon>
)
const PauseIcon = () => (
  <Icon>
    <path d="M6 4.5h4v15H6zM14 4.5h4v15h-4z" />
  </Icon>
)
const VolumeIcon = ({ level }: { level: 'mute' | 'low' | 'high' }) => (
  <svg
    width="18"
    height="18"
    viewBox="0 0 24 24"
    fill="currentColor"
    stroke="currentColor"
    strokeWidth="1.8"
    strokeLinecap="round"
    aria-hidden
    focusable="false"
  >
    <path d="M4 9.5h3.5L12 5.5v13l-4.5-4H4z" stroke="none" />
    {level === 'mute' ? (
      <path d="M15.5 9.5l5 5M20.5 9.5l-5 5" fill="none" />
    ) : (
      <>
        <path d="M15.2 9.3a4 4 0 0 1 0 5.4" fill="none" />
        {level === 'high' && <path d="M17.8 6.8a7.5 7.5 0 0 1 0 10.4" fill="none" />}
      </>
    )}
  </svg>
)

export function ShowVideoPlayer({
  src,
  style,
  onPlayingChange,
  autoPlay,
  loop,
  trim,
}: {
  src: string
  /** absolute box of the video on the slide (left/top/width/height in px) */
  style: React.CSSProperties
  onPlayingChange?: (playing: boolean) => void
  autoPlay?: boolean
  loop?: boolean
  trim?: { startMs?: number; endMs?: number }
}) {
  const { t } = useI18n()
  const wrapRef = useRef<HTMLDivElement>(null)
  const videoRef = useRef<HTMLVideoElement>(null)
  const idleTimer = useRef<ReturnType<typeof setTimeout> | undefined>(undefined)
  const hoverRef = useRef(false)
  const kbFocusRef = useRef(false)
  const [hovering, setHovering] = useState(false)
  const [idle, setIdle] = useState(false)
  const [dragging, setDragging] = useState(false)
  const [paused, setPaused] = useState(true)
  const [time, setTime] = useState(0)
  const [duration, setDuration] = useState(0)
  const [buffered, setBuffered] = useState(0)
  const [vol, setVol] = useState(rememberedVolume)

  const boxW = Number(style.width) || 0
  const boxH = Number(style.height) || 0
  const layout = videoBarLayout(boxW, boxH)

  const toggle = useCallback(() => {
    const v = videoRef.current
    if (!v) return
    if (v.paused) void v.play().catch(() => undefined)
    else v.pause()
  }, [])

  // Space plays / pauses this video while the pointer is on it (or its bar has keyboard focus)
  useEffect(() => {
    return registerShowVideoSpaceClaimer((e) => {
      if (!hoverRef.current && !kbFocusRef.current) return false
      e.preventDefault()
      toggle()
      return true
    })
  }, [toggle])

  useEffect(() => () => clearTimeout(idleTimer.current), [])

  // the element's own volume state mirrors the remembered choice from the first frame
  useEffect(() => {
    const v = videoRef.current
    if (!v) return
    v.volume = rememberedVolume.volume
    v.muted = rememberedVolume.muted
  }, [])

  const wake = () => {
    hoverRef.current = true
    setHovering(true)
    setIdle(false)
    clearTimeout(idleTimer.current)
    idleTimer.current = setTimeout(() => setIdle(true), VIDEO_CONTROLS_IDLE_MS)
  }
  const leave = () => {
    hoverRef.current = false
    clearTimeout(idleTimer.current)
    setHovering(false)
    setIdle(true)
  }

  const syncTime = () => {
    const v = videoRef.current
    if (!v) return
    if (trim?.endMs != null && v.currentTime >= trim.endMs / 1000) {
      if (loop) {
        v.currentTime = (trim.startMs ?? 0) / 1000
        void v.play().catch(() => undefined)
      } else {
        v.pause()
        v.currentTime = (trim.startMs ?? 0) / 1000
        setPaused(true)
      }
    }
    setTime(v.currentTime)
    setBuffered(bufferedEnd(v.buffered, v.currentTime))
  }
  const syncDuration = () => {
    const v = videoRef.current
    if (v && Number.isFinite(v.duration)) setDuration(v.duration)
  }

  const visible = controlsVisible({ hovering, idle, dragging }) && !layout.hidden
  const level = volumeLevel(vol.volume, vol.muted)
  const frac = mediaFraction(time, duration)
  const bufFrac = mediaFraction(buffered, duration)

  const setVolume = (next: { volume: number; muted: boolean }) => {
    rememberedVolume = next
    setVol(next)
    const v = videoRef.current
    if (v) {
      v.volume = next.volume
      v.muted = next.muted
    }
  }

  // Controls must not reach the show's click (next page) or ink / link handlers
  const swallow = {
    onClick: (e: React.MouseEvent) => e.stopPropagation(),
    onDoubleClick: (e: React.MouseEvent) => e.stopPropagation(),
    onContextMenu: (e: React.MouseEvent) => e.stopPropagation(),
  }
  const blurAfter = (e: React.SyntheticEvent<HTMLElement>) => e.currentTarget.blur()

  return (
    <div
      ref={wrapRef}
      className="ss-video"
      style={style}
      tabIndex={0}
      data-testid="show-video"
      data-controls={visible ? 'on' : 'off'}
      onPointerMove={wake}
      onPointerDown={wake}
      onPointerLeave={leave}
      onFocus={(e) => {
        if (e.target.matches(':focus-visible')) kbFocusRef.current = true
      }}
      onBlur={(e) => {
        if (!e.currentTarget.contains(e.relatedTarget as Node | null)) kbFocusRef.current = false
      }}
    >
      <video
        ref={videoRef}
        className="ss-video-el"
        src={src}
        playsInline
        preload="auto"
        loop={!!loop && !trim?.endMs}
        onClick={(e) => {
          e.stopPropagation()
          toggle()
        }}
        onContextMenu={(e) => e.stopPropagation()}
        onPlay={() => {
          setPaused(false)
          onPlayingChange?.(true)
        }}
        onPause={() => {
          setPaused(true)
          onPlayingChange?.(false)
        }}
        onTimeUpdate={syncTime}
        onSeeked={syncTime}
        onProgress={syncTime}
        onLoadedMetadata={() => {
          syncDuration()
          syncTime()
          const v = videoRef.current
          if (v) {
            if (trim?.startMs != null) {
              v.currentTime = trim.startMs / 1000
            }
            if (autoPlay) {
              void v.play().catch(() => undefined)
            }
          }
        }}
        onDurationChange={syncDuration}
        onEnded={() => {
          const v = videoRef.current
          if (loop && v) {
            v.currentTime = (trim?.startMs ?? 0) / 1000
            void v.play().catch(() => undefined)
          } else {
            setPaused(true)
          }
        }}
      />
      {!layout.hidden && (
        <div
          className={`ss-video-bar${visible ? ' ss-video-bar-on' : ''}`}
          style={{ height: layout.height }}
          role="group"
          data-testid="show-video-controls"
          {...swallow}
        >
          <button
            type="button"
            className="ss-video-btn"
            data-testid="show-video-play"
            aria-label={paused ? t('paneShowVideoPlay') : t('paneShowVideoPause')}
            data-tip={paused ? t('paneShowVideoPlay') : t('paneShowVideoPause')}
            onClick={(e) => {
              toggle()
              blurAfter(e)
            }}
          >
            {paused ? <PlayIcon /> : <PauseIcon />}
          </button>
          <div className="ss-video-track">
            <div className="ss-video-buf" style={{ width: `${bufFrac * 100}%` }} />
            <div className="ss-video-played" style={{ width: `${frac * 100}%` }} />
            <input
              type="range"
              className="ss-video-seek"
              data-testid="show-video-seek"
              aria-label={t('paneShowVideoSeek')}
              aria-valuetext={`${formatMediaTime(time)} / ${formatMediaTime(duration)}`}
              min={0}
              max={duration > 0 ? duration : 1}
              step="any"
              value={Math.min(time, duration > 0 ? duration : 1)}
              disabled={!(duration > 0)}
              onPointerDown={() => setDragging(true)}
              onPointerUp={(e) => {
                setDragging(false)
                blurAfter(e)
              }}
              onPointerCancel={() => setDragging(false)}
              onChange={(e) => {
                const v = videoRef.current
                const next = Number(e.currentTarget.value)
                if (!v || !Number.isFinite(next)) return
                v.currentTime = next
                setTime(next)
              }}
            />
          </div>
          {layout.showTime && (
            <span className="ss-video-time" data-testid="show-video-time">
              {formatMediaTime(time)} / {formatMediaTime(duration)}
            </span>
          )}
          {layout.showVolume && (
            <button
              type="button"
              className="ss-video-btn"
              data-testid="show-video-mute"
              aria-label={level === 'mute' ? t('paneShowVideoUnmute') : t('paneShowVideoMute')}
              data-tip={level === 'mute' ? t('paneShowVideoUnmute') : t('paneShowVideoMute')}
              onClick={(e) => {
                setVolume(toggleMute(vol.volume, vol.muted))
                blurAfter(e)
              }}
            >
              <VolumeIcon level={level} />
            </button>
          )}
          {layout.showVolumeSlider && (
            <input
              type="range"
              className="ss-video-vol"
              data-testid="show-video-volume"
              aria-label={t('paneShowVideoVolume')}
              min={0}
              max={1}
              step={0.05}
              value={vol.muted ? 0 : vol.volume}
              onPointerDown={() => setDragging(true)}
              onPointerUp={(e) => {
                setDragging(false)
                blurAfter(e)
              }}
              onPointerCancel={() => setDragging(false)}
              onChange={(e) => setVolume(applyVolumeInput(Number(e.currentTarget.value)))}
            />
          )}
        </div>
      )}
    </div>
  )
}
