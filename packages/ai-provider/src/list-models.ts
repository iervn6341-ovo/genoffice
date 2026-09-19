import { aiFetch } from './fetch'
import { parseModelList } from './local'

export type ListModelsResult = { ok: true; models: string[] } | { ok: false; error: string }

const LIST_MODELS_TIMEOUT_MS = 10_000

/**
 * Ask an OpenAI-compatible server which models it serves (`GET {baseUrl}/models`).
 * Never throws: connectivity, HTTP and parse failures come back as `ok: false`
 * so the settings UI can show them inline.
 */
export async function listOpenAiCompatibleModels(
  baseUrl: string,
  apiKey?: string,
  signal?: AbortSignal,
): Promise<ListModelsResult> {
  const root = baseUrl.trim().replace(/\/+$/, '')
  if (!root) return { ok: false, error: 'A custom provider requires a Base URL' }
  let parsed: URL
  try {
    parsed = new URL(root)
  } catch {
    return { ok: false, error: `Invalid Base URL: ${root}` }
  }
  // `new URL('localhost:11434')` parses (scheme "localhost:"), so a forgotten http:// must be caught here
  if (parsed.protocol !== 'http:' && parsed.protocol !== 'https:') {
    return { ok: false, error: `Base URL must start with http:// or https:// (got ${root})` }
  }
  const controller = new AbortController()
  const timer = setTimeout(() => controller.abort(), LIST_MODELS_TIMEOUT_MS)
  const onParentAbort = () => controller.abort()
  if (signal?.aborted) controller.abort()
  else signal?.addEventListener('abort', onParentAbort, { once: true })
  try {
    const response = await aiFetch(`${root}/models`, {
      method: 'GET',
      signal: controller.signal,
      headers: { ...(apiKey?.trim() ? { Authorization: `Bearer ${apiKey.trim()}` } : {}) },
    })
    if (!response.ok) {
      return { ok: false, error: `HTTP ${response.status} from ${root}/models` }
    }
    let body: unknown
    try {
      body = await response.json()
    } catch {
      return { ok: false, error: `${root}/models did not return JSON` }
    }
    const models = parseModelList(body)
    if (models.length === 0) return { ok: false, error: 'The server reported no models' }
    return { ok: true, models }
  } catch (e) {
    if (controller.signal.aborted) {
      return {
        ok: false,
        error: `No response from ${root} within ${LIST_MODELS_TIMEOUT_MS / 1000}s`,
      }
    }
    return { ok: false, error: e instanceof Error ? e.message : String(e) }
  } finally {
    clearTimeout(timer)
    signal?.removeEventListener('abort', onParentAbort)
  }
}
