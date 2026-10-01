/// Download inserted images with transient retries and per-hop SSRF checks.

import { readGeneratedImage } from './generated-images'
import { fetchWithSsrfGuard, type FetchWithSsrfGuardOptions } from './safe-remote-url'

const RETRY_DELAYS_MS: readonly number[] = [500, 1500]

export function remoteImageHeaders(_rawUrl: string): Record<string, string> {
  const headers: Record<string, string> = {
    'User-Agent': 'Mozilla/5.0',
    // Only advertise formats the insert pipelines can label correctly: callers
    // map non-png/gif responses to JPEG, so preferring avif/webp would invite
    // content-negotiating CDNs to send bytes that end up mislabeled.
    Accept: 'image/png,image/jpeg,image/gif,image/*;q=0.8,*/*;q=0.5',
  }
  return headers
}

export async function fetchRemoteImage(
  rawUrl: string,
  options: Pick<FetchWithSsrfGuardOptions, 'fetchImpl'> & {
    retryDelaysMs?: readonly number[]
  } = {},
): Promise<Response | null> {
  const { retryDelaysMs = RETRY_DELAYS_MS, ...guardOptions } = options
  // BYOK-generated images live in the local store; only its own files resolve
  if (rawUrl.startsWith('file:')) {
    const local = readGeneratedImage(rawUrl)
    if (!local) return null
    return new Response(new Uint8Array(local.bytes), {
      status: 200,
      headers: { 'content-type': local.mime, 'content-length': String(local.bytes.byteLength) },
    })
  }
  const headers = remoteImageHeaders(rawUrl)
  for (let attempt = 0; ; attempt++) {
    let resp: Response | null = null
    let threw = false
    try {
      resp = await fetchWithSsrfGuard(rawUrl, { ...guardOptions, headers })
    } catch {
      threw = true
    }
    if (resp?.ok) return resp
    if (resp === null && !threw) return null // blocked by the SSRF guard: permanent
    const transient =
      threw ||
      (resp !== null &&
        (resp.status === 403 || resp.status === 408 || resp.status === 429 || resp.status >= 500))
    const delay = retryDelaysMs[attempt]
    if (!transient || delay === undefined) return resp
    await new Promise((r) => setTimeout(r, delay))
  }
}
