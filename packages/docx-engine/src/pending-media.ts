/**
 * Pictures that arrive in the body XML before they have a package home.
 *
 * A drawing copied from another document still names the SOURCE package's
 * relationship id (`r:embed="rId7"`), which means nothing — or the wrong
 * picture — in the destination. The editor re-points such a drawing at a
 * token and registers the bytes here; the save lands the bytes as a media part
 * + relationship and swaps the token for the real rId (patch.ts).
 */
export const PENDING_IMG_PREFIX = 'aidocs-pending-img:'

export interface PendingMedia {
  base64: string
  mime: 'image/png' | 'image/jpeg' | 'image/gif'
}

const registry = new Map<string, PendingMedia>()

/** short stable key of the bytes (equal pictures share one token and one media part) */
function keyOf(media: PendingMedia): string {
  let h1 = 5381
  let h2 = 52711
  const s = media.base64
  for (let i = 0; i < s.length; i++) {
    const c = s.charCodeAt(i)
    h1 = (h1 * 33) ^ c
    h2 = (h2 * 31 + c) | 0
  }
  return `${(h1 >>> 0).toString(36)}${(h2 >>> 0).toString(36)}${s.length.toString(36)}`
}

/** data:image/(png|jpeg|gif);base64 URL → bytes, null for anything the package cannot store as is */
export function pendingMediaOfDataUrl(dataUrl: string): PendingMedia | null {
  const m = /^data:image\/(png|jpeg|gif);base64,([A-Za-z0-9+/=]+)$/.exec(dataUrl)
  return m ? { base64: m[2]!, mime: `image/${m[1]}` as PendingMedia['mime'] } : null
}

/** register bytes, get the r:embed token for them */
export function registerPendingMedia(media: PendingMedia): string {
  const key = keyOf(media)
  registry.set(key, media)
  return `${PENDING_IMG_PREFIX}${key}`
}

export function lookupPendingMedia(token: string): PendingMedia | undefined {
  return token.startsWith(PENDING_IMG_PREFIX)
    ? registry.get(token.slice(PENDING_IMG_PREFIX.length))
    : undefined
}

/** point every `r:embed` of a drawing fragment at the pending token for `media` */
export function retargetDrawingToPending(drawingXml: string, media: PendingMedia): string {
  const token = registerPendingMedia(media)
  return drawingXml.replace(/\br:embed="[^"]*"/g, `r:embed="${token}"`)
}
