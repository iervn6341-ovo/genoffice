/**
 * Helpers for self-hosted / on-prem OpenAI-compatible servers (Ollama, LM Studio,
 * llama.cpp, vLLM, an intranet gateway). They ride the `custom` provider; nothing
 * here is a new wire protocol.
 */

export interface LocalEndpointPreset {
  id: string
  label: string
  /** the server's default OpenAI-compatible base URL */
  baseUrl: string
}

/** Default listen addresses of the common local servers, each with its `/v1` API root. */
export const LOCAL_ENDPOINT_PRESETS: readonly LocalEndpointPreset[] = [
  { id: 'ollama', label: 'Ollama', baseUrl: 'http://localhost:11434/v1' },
  { id: 'lmstudio', label: 'LM Studio', baseUrl: 'http://localhost:1234/v1' },
  { id: 'llamacpp', label: 'llama.cpp', baseUrl: 'http://localhost:8080/v1' },
  { id: 'vllm', label: 'vLLM', baseUrl: 'http://localhost:8000/v1' },
]

/** The preset whose base URL matches `baseUrl` (trailing slashes and case ignored), if any. */
export function presetForBaseUrl(baseUrl: string | undefined): LocalEndpointPreset | undefined {
  const normalized = (baseUrl ?? '').trim().replace(/\/+$/, '').toLowerCase()
  if (!normalized) return undefined
  return LOCAL_ENDPOINT_PRESETS.find((p) => p.baseUrl.toLowerCase() === normalized)
}

function isPrivateIPv4(host: string): boolean {
  const m = /^(\d{1,3})\.(\d{1,3})\.(\d{1,3})\.(\d{1,3})$/.exec(host)
  if (!m) return false
  const [a, b] = [Number(m[1]), Number(m[2])]
  const octets = m.slice(1).map(Number)
  if (octets.some((o) => o > 255)) return false
  return (
    a === 127 || // loopback
    a === 10 || // RFC 1918
    (a === 172 && b >= 16 && b <= 31) ||
    (a === 192 && b === 168) ||
    (a === 169 && b === 254) // link-local
  )
}

/**
 * True when the URL points at this machine or the local network: loopback,
 * RFC 1918 / link-local addresses, `localhost`, `*.local` and single-label
 * intranet hostnames. Such servers run on the user's own hardware, where a first
 * token can legitimately take minutes (CPU prefill, a cold model load).
 */
export function isLocalEndpoint(baseUrl: string | undefined): boolean {
  if (!baseUrl) return false
  let host: string
  try {
    host = new URL(baseUrl.trim()).hostname.toLowerCase()
  } catch {
    return false
  }
  // URL keeps the brackets on IPv6 literals
  if (host.startsWith('[') && host.endsWith(']')) host = host.slice(1, -1)
  if (host === 'localhost' || host.endsWith('.localhost') || host.endsWith('.local')) return true
  if (host === '::1' || host.startsWith('fe80:') || /^f[cd][0-9a-f]{2}:/.test(host)) return true
  if (isPrivateIPv4(host)) return true
  // a bare intranet name such as `llm-box` (no dots, not an IP)
  return !host.includes('.') && !host.includes(':') && host.length > 0
}

/**
 * Timeouts for local servers. The cloud defaults (60s to the first byte, 180s
 * idle) fit a hosted API but kill a large local model still prefilling a long
 * document, so a local endpoint gets a generous budget. A truly dead socket is
 * still caught, just later; the user can always cancel the turn.
 */
export const AI_LOCAL_CONNECT_TIMEOUT_MS = 600_000
export const AI_LOCAL_IDLE_TIMEOUT_MS = 600_000

/** Pull the model ids out of an OpenAI-style `/models` body (also tolerant of Ollama-style shapes). */
export function parseModelList(body: unknown): string[] {
  if (typeof body !== 'object' || body === null) return []
  const record = body as { data?: unknown; models?: unknown }
  const rows = Array.isArray(record.data)
    ? record.data
    : Array.isArray(record.models)
      ? record.models
      : []
  const ids: string[] = []
  for (const row of rows) {
    let id: unknown
    if (typeof row === 'string') id = row
    else if (typeof row === 'object' && row !== null) {
      const r = row as { id?: unknown; name?: unknown; model?: unknown }
      id = r.id ?? r.model ?? r.name
    }
    if (typeof id === 'string' && id.trim()) ids.push(id.trim())
  }
  return [...new Set(ids)].sort((a, b) => a.localeCompare(b))
}
