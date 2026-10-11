import { aiFetch } from './fetch'

/** how long a local server gets to answer a bare GET before it counts as unreachable */
export const REACH_PROBE_TIMEOUT_MS = 8000

/** `fetch failed` hides the reason: the network error code sits on `cause` (ECONNREFUSED, EHOSTUNREACH, …) */
export function describeNetworkError(error: unknown): string {
  const message = error instanceof Error ? error.message : String(error)
  const cause = (error as { cause?: { code?: unknown; message?: unknown } } | null)?.cause
  const detail =
    typeof cause?.code === 'string'
      ? cause.code
      : typeof cause?.message === 'string'
        ? cause.message
        : ''
  return detail && !message.includes(detail) ? `${message} (${detail})` : message
}

/**
 * Local / LAN servers get a 10-minute first-byte budget (a cold model can really take that long
 * to prefill), which also means a server that cannot be reached at all looks like a request that
 * never ends. A bare GET answers or fails within seconds whatever the model is doing, so ask it
 * first and report an unreachable host as such. Any HTTP answer (even 404) proves the route works.
 */
export async function assertEndpointReachable(
  baseUrl: string,
  signal?: AbortSignal,
  timeoutMs = REACH_PROBE_TIMEOUT_MS,
): Promise<void> {
  const controller = new AbortController()
  const timer = setTimeout(() => controller.abort(), timeoutMs)
  const onParentAbort = () => controller.abort()
  if (signal?.aborted) controller.abort()
  else signal?.addEventListener('abort', onParentAbort, { once: true })
  try {
    await aiFetch(`${baseUrl.replace(/\/+$/, '')}/models`, {
      method: 'GET',
      signal: controller.signal,
    })
  } catch (error) {
    // the caller cancelled: let that surface as the cancel it is
    if (signal?.aborted) throw error
    const reason = controller.signal.aborted
      ? `no answer within ${Math.round(timeoutMs / 1000)}s`
      : describeNetworkError(error)
    throw new Error(
      `Cannot reach ${baseUrl} (${reason}). Check that the server is running and listening on ` +
        'the network (Ollama needs OLLAMA_HOST=0.0.0.0 to accept other machines), and on macOS ' +
        'allow GenOffice under System Settings → Privacy & Security → Local Network.',
    )
  } finally {
    clearTimeout(timer)
    signal?.removeEventListener('abort', onParentAbort)
  }
}
