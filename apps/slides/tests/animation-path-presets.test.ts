import { describe, expect, it } from 'vitest'
import { buildTimingXml, isSupportedMotionPath, readSlideTimingXml } from '@genoffice/pptx-engine'
import { MORE_MOTION_PATHS } from '../src/renderer/components/Ribbon'
import { pointAtPath, samplePathPoints } from '../src/renderer/animation-play'

describe('additional motion paths', () => {
  it('all authored presets use the native and preview-supported path subset', () => {
    expect(MORE_MOTION_PATHS.map((preset) => preset.group)).toContain('complex')
    for (const preset of MORE_MOTION_PATHS) {
      expect(isSupportedMotionPath(preset.path), preset.label).toBe(true)
      const points = samplePathPoints(preset.path)
      expect(points.length, preset.label).toBeGreaterThan(2)
      expect(points.every((point) => Number.isFinite(point.x) && Number.isFinite(point.y))).toBe(
        true,
      )
      expect(pointAtPath(points, 0)).toEqual({ x: 0, y: 0 })
      const timing = buildTimingXml([
        {
          spid: 8,
          effect: 'motionPath',
          trigger: 'onClick',
          durationMs: 1500,
          delayMs: 0,
          motionPath: preset.path,
        },
      ])
      expect(readSlideTimingXml(timing)[0]!.motionPath).toBe(preset.path)
    }
  })

  it('rejects unpreviewable commands and XML markup rather than saving a malformed animation', () => {
    for (const path of [
      'M 0 0 Q 1 1 2 2',
      'M 0 0 L "bad" 1',
      'M 0 0 L Infinity 0',
      'M 0 0',
      'L 0 0',
      'M 0 0 L 11 0',
    ])
      expect(isSupportedMotionPath(path), path).toBe(false)
  })
})
