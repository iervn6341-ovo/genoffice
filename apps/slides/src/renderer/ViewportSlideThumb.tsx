import React, { useEffect, useRef, useState } from 'react'
import type { RenderSlide } from '@genoffice/pptx-render'
import { SlideThumb } from './SlideThumb'

/** Keep list geometry and drag targets, but allocate canvases only near the viewport.
 * Exports continue using SlideThumb directly, so every exported page renders eagerly.
 */
export const ViewportSlideThumb = React.memo(function ViewportSlideThumb({
  slide,
  images,
  width,
  current = false,
}: {
  slide: RenderSlide
  images: Map<string, HTMLImageElement>
  width: number
  current?: boolean
}) {
  const host = useRef<HTMLDivElement>(null)
  const [visible, setVisible] = useState(() => typeof IntersectionObserver === 'undefined')
  useEffect(() => {
    const el = host.current
    if (!el) return
    if (typeof IntersectionObserver === 'undefined') return
    const observer = new IntersectionObserver(
      ([entry]) => setVisible(entry?.isIntersecting ?? false),
      { rootMargin: '200px' },
    )
    observer.observe(el)
    return () => observer.disconnect()
  }, [])
  return (
    <div ref={host} style={{ width, height: (width * slide.heightPx) / slide.widthPx }}>
      {(visible || current) && <SlideThumb slide={slide} images={images} width={width} />}
    </div>
  )
})
