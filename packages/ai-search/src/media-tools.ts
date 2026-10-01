import { existsSync, readFileSync, statSync } from 'node:fs'
import { basename, extname } from 'node:path'
import {
  activeMediaConfig,
  analyzeMediaWithProvider,
  defaultAiSettings,
  generateImageWithProvider,
  resolveAiSettings,
  type AiSettings,
  type LegacyAiSettings,
  type MediaBlob,
} from '@genoffice/ai-provider'
// deep imports: the package root re-exports Electron-bound modules, and this file also runs in the genoffice CLI
import { readGeneratedImage, storeGeneratedImage } from '@genoffice/electron-utils/generated-images'
import { fetchRemoteImage } from '@genoffice/electron-utils/remote-image'
import { fetchWithSsrfGuard } from '@genoffice/electron-utils/safe-remote-url'

/** 200 MB: enough for a long clip through the Gemini Files API, small enough to hold in memory */
const MAX_MEDIA_BYTES = 200 * 1024 * 1024

export class MediaTooLargeError extends Error {}

const MIME_BY_EXT: Record<string, string> = {
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.jpeg': 'image/jpeg',
  '.gif': 'image/gif',
  '.webp': 'image/webp',
  '.bmp': 'image/bmp',
  '.heic': 'image/heic',
  '.mp4': 'video/mp4',
  '.mov': 'video/quicktime',
  '.webm': 'video/webm',
  '.m4v': 'video/mp4',
  '.mkv': 'video/x-matroska',
  '.mp3': 'audio/mpeg',
  '.wav': 'audio/wav',
  '.m4a': 'audio/mp4',
  '.aac': 'audio/aac',
  '.ogg': 'audio/ogg',
  '.flac': 'audio/flac',
}

export function readAiSettingsFile(path: string): AiSettings {
  let stored: Partial<AiSettings> & LegacyAiSettings = {}
  try {
    if (existsSync(path)) stored = JSON.parse(readFileSync(path, 'utf-8'))
  } catch {
    /* corrupted settings file: defaults */
  }
  return resolveAiSettings(stored, defaultAiSettings())
}

function errorText(err: unknown): string {
  return err instanceof Error ? err.message : String(err)
}

/**
 * Resolves a tool-supplied media reference to bytes: an https URL (SSRF-guarded),
 * a file:// URL from the generated-image store, or a local media file
 * (attachments). Only media extensions are read locally — the model must not be
 * able to ship arbitrary files to a vendor.
 */
export async function loadMediaReference(ref: string): Promise<MediaBlob> {
  if (/^https?:\/\//i.test(ref)) {
    const resp = await (ref.match(/\.(png|jpe?g|gif|webp)(\?|$)/i)
      ? fetchRemoteImage(ref)
      : fetchWithSsrfGuard(ref, { headers: { 'User-Agent': 'Mozilla/5.0' } }))
    if (!resp || !resp.ok) throw new Error(`Could not download ${ref}`)
    const declared = Number(resp.headers.get('content-length') ?? 0)
    if (declared > MAX_MEDIA_BYTES) throw new MediaTooLargeError(`${ref} is too large to analyze`)
    const bytes = new Uint8Array(await resp.arrayBuffer())
    if (bytes.byteLength > MAX_MEDIA_BYTES) {
      throw new MediaTooLargeError(`${ref} is too large to analyze`)
    }
    const rawCt = resp.headers.get('content-type')?.split(';')[0]?.trim().toLowerCase()
    const ct = rawCt && rawCt !== 'application/octet-stream' ? rawCt : undefined
    const name = basename(new URL(ref).pathname) || undefined
    const mime =
      ct && ct !== 'application/octet-stream' ? ct : MIME_BY_EXT[extname(name ?? '').toLowerCase()]
    if (!mime) throw new Error(`Could not tell the media type of ${ref}`)
    return { bytes, mime, ...(name ? { name } : {}) }
  }
  if (ref.startsWith('file:')) {
    const local = readGeneratedImage(ref)
    if (!local) throw new Error(`Not an accessible image: ${ref}`)
    return { bytes: new Uint8Array(local.bytes), mime: local.mime }
  }
  const mime = MIME_BY_EXT[extname(ref).toLowerCase()]
  if (!mime) throw new Error(`Unsupported media file: ${ref} (images, video and audio only)`)
  if (!existsSync(ref)) throw new Error(`File not found: ${ref}`)
  if (statSync(ref).size > MAX_MEDIA_BYTES) {
    throw new MediaTooLargeError(`${ref} is too large to analyze`)
  }
  return { bytes: new Uint8Array(readFileSync(ref)), mime, name: basename(ref) }
}

export type GenerateImageToolOp = {
  prompt: string
  imageSize?: string | undefined
  aspectRatio?: string | undefined
  referenceImageUrls?: string[] | undefined
  model?: string | undefined
  /** Request true alpha transparency when supported by the configured provider. */
  transparentBackground?: boolean
}

export async function generateImageTool(
  settingsPath: string,
  op: GenerateImageToolOp,
): Promise<{ url?: string; error?: string }> {
  const prompt = String(op.prompt ?? '').trim()
  if (!prompt) return { error: 'prompt must not be empty' }
  const settings = readAiSettingsFile(settingsPath)
  const byok = activeMediaConfig(settings, 'image')
  if (!byok)
    return { error: 'Configure an image provider in Settings > AI Media to use this tool.' }
  try {
    const references = await Promise.all((op.referenceImageUrls ?? []).map(loadMediaReference))
    const image = await generateImageWithProvider(byok.provider, byok.config, {
      prompt,
      aspectRatio: op.aspectRatio,
      references,
      transparent: op.transparentBackground === true,
    })
    return { url: storeGeneratedImage(image.bytes, image.mime) }
  } catch (err) {
    return { error: errorText(err) }
  }
}

export async function analyzeMediaTool(
  settingsPath: string,
  op: { mediaUrls: string[]; requirements: string },
): Promise<{ text?: string; error?: string }> {
  const mediaUrls = (op.mediaUrls ?? []).map(String).filter(Boolean)
  const requirements = String(op.requirements ?? '').trim()
  if (!mediaUrls.length) return { error: 'mediaUrls must not be empty' }
  if (!requirements) return { error: 'requirements must not be empty' }
  const settings = readAiSettingsFile(settingsPath)
  const imageByok = activeMediaConfig(settings, 'analysis')
  const videoByok = activeMediaConfig(settings, 'video')
  try {
    const unavailable = () => ({
      error: 'Configure a compatible media provider in Settings > AI Media to use this tool.',
    })
    if (!imageByok && !videoByok) return unavailable()
    // route on the loaded bytes' real MIME, not the URL spelling: images go to the
    // image-analysis provider, anything with video/audio to the video one
    let media: MediaBlob[]
    try {
      media = await Promise.all(mediaUrls.map(loadMediaReference))
    } catch (err) {
      // Scheme, path, and SSRF rejections are reported without a provider fallback.
      return { error: errorText(err) }
    }
    const hasVideo = media.some((m) => !m.mime.startsWith('image/'))
    const byok = hasVideo ? videoByok : imageByok
    if (!byok) return unavailable()
    return {
      text: await analyzeMediaWithProvider(byok.provider, byok.config, { media, requirements }),
    }
  } catch (err) {
    return { error: errorText(err) }
  }
}
