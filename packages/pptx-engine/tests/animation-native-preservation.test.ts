import { describe, expect, it } from 'vitest'
import {
  buildTimingXml,
  patchSlideTimingIncrementalXml,
  patchSlideTimingXml,
  readSlideTimingXml,
  type SlideAnimation,
} from '../src/animation'

const fade: SlideAnimation = {
  spid: 4,
  effect: 'fade',
  trigger: 'onClick',
  durationMs: 500,
  delayMs: 0,
}
const fly: SlideAnimation = {
  spid: 5,
  effect: 'flyIn',
  trigger: 'onClick',
  durationMs: 700,
  delayMs: 0,
}
const suffix = (timing: string) => `${timing}</p:sld>`

function importedTiming(): string {
  const base = buildTimingXml([fade, fly])
  const parsed = readSlideTimingXml(suffix(base), true)
  const second = parsed[1]!.nativeXml!
  const rich = second
    .replace('presetID="2"', 'presetID="77"')
    .replace(
      '</p:childTnLst>',
      '<p:extLst><p:ext uri="keep-effect-child"/></p:extLst></p:childTnLst>',
    )
  return base.replace(second, rich)
}

describe('native animation effect preservation', () => {
  it('editing a different effect leaves an imported unsupported effect and its IDs byte-identical', () => {
    const timing = importedTiming()
    const before = readSlideTimingXml(suffix(timing), true)
    expect(before[1]!.preset).toMatchObject({ id: 77 })
    const imported = before[1]!.nativeXml!
    const edited = [{ ...before[0]!, durationMs: 900 }, before[1]!]
    const output = patchSlideTimingIncrementalXml(suffix(timing), edited)
    expect(output).not.toBeNull()
    expect(output).toContain(imported)
    expect(readSlideTimingXml(output!, true)[1]!.nativeId).toBe(before[1]!.nativeId)
  })

  it('editing imported timing scales native behavior durations and keeps unknown child XML and IDs', () => {
    const timing = importedTiming()
    const before = readSlideTimingXml(suffix(timing), true)
    const imported = before[1]!
    const edited = [before[0]!, { ...imported, durationMs: 1400, delayMs: 150 }]
    const output = patchSlideTimingIncrementalXml(suffix(timing), edited)!
    const result = readSlideTimingXml(output, true)[1]!
    expect(result).toMatchObject({ durationMs: 1400, delayMs: 150, nativeId: imported.nativeId })
    expect(result.nativeXml).toContain('<p:ext uri="keep-effect-child"/>')
    expect(result.nativeXml).toContain('presetID="77"')
  })

  it('a trigger regroup keeps imported effect IDs and unsupported child XML', () => {
    const timing = importedTiming()
    const before = readSlideTimingXml(suffix(timing), true)
    const output = patchSlideTimingXml(suffix(timing), [
      before[0]!,
      { ...before[1]!, trigger: 'afterPrev', delayMs: 200 },
    ])
    const after = readSlideTimingXml(output, true)
    expect(after[1]).toMatchObject({
      trigger: 'afterPrev',
      delayMs: 200,
      nativeId: before[1]!.nativeId,
    })
    expect(after[1]!.nativeXml).toContain('<p:ext uri="keep-effect-child"/>')
    const ids = [...output.matchAll(/<p:cTn\b[^>]*\bid="(\d+)"/g)].map((m) => m[1])
    expect(new Set(ids).size).toBe(ids.length)
  })
})
