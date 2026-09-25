/**
 * PowerPoint's Camera button during a show: the presenter's live camera in a round bubble on
 * the slide. Each screen (presenter view, audience window, full-screen show) opens its own
 * stream from the same device and releases it when the camera is turned off or the show ends.
 */
import React, { useEffect, useRef } from 'react'

export function CameraBubble({
  on,
  width,
  onError,
}: {
  on: boolean
  /** Slide frame width (px): the bubble is sized relative to it */
  width: number
  /** No camera, or access was refused */
  onError?: () => void
}): React.JSX.Element | null {
  const videoRef = useRef<HTMLVideoElement>(null)
  const onErrorRef = useRef(onError)
  onErrorRef.current = onError
  useEffect(() => {
    if (!on) return
    let stream: MediaStream | null = null
    let cancelled = false
    const media = navigator.mediaDevices
    if (!media?.getUserMedia) {
      onErrorRef.current?.()
      return
    }
    media.getUserMedia({ video: { width: 640, height: 480 }, audio: false }).then(
      (s) => {
        if (cancelled) {
          s.getTracks().forEach((tr) => tr.stop())
          return
        }
        stream = s
        const v = videoRef.current
        if (v) {
          v.srcObject = s
          void v.play().catch(() => {})
        }
      },
      () => onErrorRef.current?.(),
    )
    return () => {
      cancelled = true
      stream?.getTracks().forEach((tr) => tr.stop())
    }
  }, [on])
  if (!on) return null
  const d = Math.round(width * 0.18)
  return (
    <div className="show-camera" style={{ width: d, height: d }}>
      <video ref={videoRef} muted playsInline />
    </div>
  )
}
