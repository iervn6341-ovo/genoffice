import { describe, expect, it } from 'vitest'
import {
  createBlankPptx,
  materializeSlide,
  openPptx,
  savePptx,
  type PictureElement,
} from '@genoffice/pptx-engine'
import { runTxn } from '@genoffice/pptx-ops'
import {
  NARRATION_NAME,
  bytesToBase64,
  encodeWav,
  mixToMono,
  withoutNarration,
} from '../src/renderer/narration'
import { narrationNodes } from '../src/renderer/narration-actions'

describe('narration audio encoding', () => {
  it('writes a valid 16-bit mono PCM WAV header and clamps samples', () => {
    const wav = encodeWav(Float32Array.from([0, 1, -1, 2, -2]), 22050)
    const view = new DataView(wav.buffer)
    const tag = (at: number) => String.fromCharCode(...wav.subarray(at, at + 4))
    expect([tag(0), tag(8), tag(12), tag(36)]).toEqual(['RIFF', 'WAVE', 'fmt ', 'data'])
    expect(view.getUint32(4, true)).toBe(36 + 10)
    expect(view.getUint16(22, true)).toBe(1)
    expect(view.getUint32(24, true)).toBe(22050)
    expect(view.getUint32(40, true)).toBe(10)
    expect(view.getInt16(46, true)).toBe(32767)
    expect(view.getInt16(48, true)).toBe(-32768)
    expect(view.getInt16(50, true)).toBe(32767)
    expect(view.getInt16(52, true)).toBe(-32768)
  })

  it('mixes channels to mono and base64 round-trips', () => {
    expect([...mixToMono([Float32Array.from([1, 0]), Float32Array.from([0, 1])])]).toEqual([
      0.5, 0.5,
    ])
    expect(Buffer.from(bytesToBase64(Uint8Array.from([1, 2, 3, 250])), 'base64')).toEqual(
      Buffer.from([1, 2, 3, 250]),
    )
  })
})

describe('narration on a slide', () => {
  it('is an embedded wav audio shape that survives save/reopen and can be removed', async () => {
    const opened = await openPptx(await createBlankPptx())
    const wav = encodeWav(new Float32Array(2000), 8000)
    const added = runTxn(opened, {
      ops: [
        {
          op: 'addMedia',
          target: { slide: 0 },
          kind: 'audio',
          bytes: wav,
          ext: 'wav',
          name: NARRATION_NAME,
          offset: { x: 8000000, y: 5000000, cx: 457200, cy: 457200 },
        },
      ],
    })
    expect(added.applied).toBe(true)
    const reopened = await openPptx(await savePptx(opened))
    const pic = reopened.deck.slides[0]!.elements.find(
      (el) => el.type === 'picture',
    ) as PictureElement
    expect(pic.name).toBe(NARRATION_NAME)
    expect(pic.media?.kind).toBe('audio')
    expect([...reopened.archive.entries.keys()].some((k) => /ppt\/media\/.*\.wav$/.test(k))).toBe(
      true,
    )
    const removed = runTxn(reopened, {
      ops: [{ op: 'deleteElement', target: { slide: 0, el: pic.id } }],
    })
    expect(removed.applied).toBe(true)
    expect(materializeSlide(reopened, 0)!.elements.some((el) => el.type === 'picture')).toBe(false)
  })

  it('finds narration nodes and strips them for presenting', () => {
    const slide = {
      widthPx: 10,
      heightPx: 10,
      scale: 1,
      background: {},
      nodes: [
        { type: 'picture', media: 'audio', name: NARRATION_NAME, sourceId: 'a' },
        { type: 'picture', media: 'audio', name: 'Other', sourceId: 'b' },
      ],
    } as never
    expect(narrationNodes(slide).map((n) => n.sourceId)).toEqual(['a'])
    expect((withoutNarration(slide) as { nodes: unknown[] }).nodes).toHaveLength(1)
  })
})
