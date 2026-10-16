/**
 * Slide narration: microphone recording stored as an ordinary embedded WAV audio
 * shape on the slide (named NARRATION_NAME). The deck stays a plain PPTX that any
 * PowerPoint can play; GenOffice's own show auto-plays the shape when "Play
 * narrations" is on. No p:timing node is written, so PowerPoint plays it on click.
 */
import type { PictureRenderNode, RenderSlide } from '@genoffice/pptx-render'

export const NARRATION_NAME = 'GenOffice Narration'

/** The slide without its narration icon (the icon is hidden while presenting). */
export function withoutNarration(slide: RenderSlide): RenderSlide {
  const nodes = slide.nodes.filter(
    (n) => !(n.type === 'picture' && (n as PictureRenderNode).name === NARRATION_NAME),
  )
  return nodes.length === slide.nodes.length ? slide : { ...slide, nodes }
}

/** 16-bit mono PCM WAV (RIFF) of Float32 samples in [-1, 1]. */
export function encodeWav(samples: Float32Array, sampleRate: number): Uint8Array {
  const bytes = new Uint8Array(44 + samples.length * 2)
  const view = new DataView(bytes.buffer)
  const text = (at: number, s: string) => {
    for (let i = 0; i < s.length; i++) view.setUint8(at + i, s.charCodeAt(i))
  }
  text(0, 'RIFF')
  view.setUint32(4, 36 + samples.length * 2, true)
  text(8, 'WAVE')
  text(12, 'fmt ')
  view.setUint32(16, 16, true)
  view.setUint16(20, 1, true) // PCM
  view.setUint16(22, 1, true) // mono
  view.setUint32(24, sampleRate, true)
  view.setUint32(28, sampleRate * 2, true)
  view.setUint16(32, 2, true)
  view.setUint16(34, 16, true)
  text(36, 'data')
  view.setUint32(40, samples.length * 2, true)
  for (let i = 0; i < samples.length; i++) {
    const s = Math.max(-1, Math.min(1, samples[i]!))
    view.setInt16(44 + i * 2, s < 0 ? s * 0x8000 : s * 0x7fff, true)
  }
  return bytes
}

/** Average all channels into one. */
export function mixToMono(channels: Float32Array[]): Float32Array {
  if (channels.length === 1) return channels[0]!
  const out = new Float32Array(channels[0]?.length ?? 0)
  for (const channel of channels)
    for (let i = 0; i < out.length; i++) out[i] = (out[i] ?? 0) + (channel[i] ?? 0)
  for (let i = 0; i < out.length; i++) out[i] = (out[i] ?? 0) / channels.length
  return out
}

export function bytesToBase64(bytes: Uint8Array): string {
  let binary = ''
  const chunk = 0x8000
  for (let i = 0; i < bytes.length; i += chunk)
    binary += String.fromCharCode(...bytes.subarray(i, i + chunk))
  return btoa(binary)
}

let active: { rec: MediaRecorder; stream: MediaStream; chunks: Blob[] } | null = null

export function isNarrationRecording(): boolean {
  return active !== null
}

/** Ask for the microphone and start recording; false when it is unavailable or denied. */
export async function startNarrationRecording(): Promise<boolean> {
  if (active) return true
  try {
    const stream = await navigator.mediaDevices.getUserMedia({ audio: true })
    const rec = new MediaRecorder(stream)
    const chunks: Blob[] = []
    rec.ondataavailable = (ev) => {
      if (ev.data.size > 0) chunks.push(ev.data)
    }
    rec.start(500)
    active = { rec, stream, chunks }
    return true
  } catch {
    return false
  }
}

/** Stop and convert to WAV; null when nothing usable was captured. */
export async function stopNarrationRecording(): Promise<Uint8Array | null> {
  const current = active
  active = null
  if (!current) return null
  const done = new Promise<void>((resolve) => {
    current.rec.onstop = () => resolve()
  })
  if (current.rec.state !== 'inactive') current.rec.stop()
  await done
  current.stream.getTracks().forEach((track) => track.stop())
  if (!current.chunks.length) return null
  const ctx = new AudioContext()
  try {
    const decoded = await ctx.decodeAudioData(await new Blob(current.chunks).arrayBuffer())
    const channels = Array.from({ length: decoded.numberOfChannels }, (_, i) =>
      decoded.getChannelData(i),
    )
    if (!decoded.length) return null
    return encodeWav(mixToMono(channels), decoded.sampleRate)
  } catch {
    return null
  } finally {
    void ctx.close()
  }
}
