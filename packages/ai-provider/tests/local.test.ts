import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { chatForProvider } from '../src/chat'
import { listOpenAiCompatibleModels } from '../src/list-models'
import {
  AI_LOCAL_CONNECT_TIMEOUT_MS,
  LOCAL_ENDPOINT_PRESETS,
  isLocalEndpoint,
  parseModelList,
  presetForBaseUrl,
} from '../src/local'
import { activeProvider, defaultAiSettings, providerRequiresApiKey } from '../src/providers'
import { streamForProvider } from '../src/stream'
import { AI_CHAT_RESPONSE_TIMEOUT_MS, AI_CONNECT_TIMEOUT_MS, AiTimeoutError } from '../src/watchdog'
import { jsonResponse } from './test-utils'

afterEach(() => {
  vi.useRealTimers()
  vi.unstubAllGlobals()
})

describe('LOCAL_ENDPOINT_PRESETS', () => {
  it('lists each common local server once, on an OpenAI-compatible /v1 root', () => {
    const ids = LOCAL_ENDPOINT_PRESETS.map((p) => p.id)
    expect(new Set(ids).size).toBe(ids.length)
    expect(ids).toEqual(expect.arrayContaining(['ollama', 'lmstudio', 'llamacpp', 'vllm']))
    for (const preset of LOCAL_ENDPOINT_PRESETS) {
      expect(preset.baseUrl).toMatch(/^http:\/\/localhost:\d+\/v1$/)
      expect(isLocalEndpoint(preset.baseUrl)).toBe(true)
    }
    expect(LOCAL_ENDPOINT_PRESETS.find((p) => p.id === 'ollama')?.baseUrl).toBe(
      'http://localhost:11434/v1',
    )
    expect(LOCAL_ENDPOINT_PRESETS.find((p) => p.id === 'lmstudio')?.baseUrl).toBe(
      'http://localhost:1234/v1',
    )
  })

  it('presetForBaseUrl matches ignoring trailing slashes and case, and returns undefined otherwise', () => {
    expect(presetForBaseUrl('http://localhost:11434/v1/')?.id).toBe('ollama')
    expect(presetForBaseUrl(' HTTP://LOCALHOST:1234/v1 ')?.id).toBe('lmstudio')
    expect(presetForBaseUrl('http://localhost:9999/v1')).toBeUndefined()
    expect(presetForBaseUrl('')).toBeUndefined()
    expect(presetForBaseUrl(undefined)).toBeUndefined()
  })
})

describe('isLocalEndpoint', () => {
  it.each([
    'http://localhost:11434/v1',
    'http://LOCALHOST:8000/v1',
    'http://127.0.0.1:8080/v1',
    'http://127.5.6.7/v1',
    'http://[::1]:8000/v1',
    'http://10.0.0.5:8000/v1',
    'http://172.16.0.1/v1',
    'http://172.31.255.255/v1',
    'http://192.168.1.20:1234/v1',
    'http://169.254.10.10/v1',
    'http://[fd12:3456::1]:8000/v1',
    'http://gpu-box.local:8000/v1',
    'http://llm-server:8000/v1',
  ])('treats %s as local', (url) => {
    expect(isLocalEndpoint(url)).toBe(true)
  })

  it.each([
    'https://api.openai.com/v1',
    'https://api.anthropic.com',
    'https://mirror.example.com/v1',
    'http://8.8.8.8/v1',
    'http://172.32.0.1/v1', // just outside 172.16/12
    'http://172.15.0.1/v1',
    'http://192.169.1.1/v1',
    'http://11.0.0.1/v1',
    'http://[2001:db8::1]/v1',
    'https://llm.company.com:8443/v1',
  ])('treats %s as remote', (url) => {
    expect(isLocalEndpoint(url)).toBe(false)
  })

  it('is false for empty, missing or unparsable input instead of throwing', () => {
    expect(isLocalEndpoint(undefined)).toBe(false)
    expect(isLocalEndpoint('')).toBe(false)
    expect(isLocalEndpoint('not a url')).toBe(false)
    expect(isLocalEndpoint('http://10.0.0.999/v1')).toBe(false)
  })
})

describe('parseModelList', () => {
  it('reads the OpenAI /models shape', () => {
    expect(
      parseModelList({ object: 'list', data: [{ id: 'qwen3:8b' }, { id: 'llama3.2:3b' }] }),
    ).toEqual(['llama3.2:3b', 'qwen3:8b'])
  })

  it('tolerates Ollama-style `models` rows keyed by model or name, and bare strings', () => {
    expect(parseModelList({ models: [{ model: 'phi4' }, { name: 'gemma3' }, 'mistral'] })).toEqual([
      'gemma3',
      'mistral',
      'phi4',
    ])
  })

  it('de-duplicates, trims, sorts and drops junk rows', () => {
    expect(
      parseModelList({ data: [{ id: ' b ' }, { id: 'a' }, { id: 'b' }, { id: '' }, {}, null, 7] }),
    ).toEqual(['a', 'b'])
  })

  it('returns an empty list for non-object or unrecognised bodies', () => {
    expect(parseModelList(null)).toEqual([])
    expect(parseModelList('models')).toEqual([])
    expect(parseModelList({ data: 'nope' })).toEqual([])
    expect(parseModelList({})).toEqual([])
  })
})

describe('listOpenAiCompatibleModels', () => {
  it('GETs {baseUrl}/models without an Authorization header when no key is given', async () => {
    const fetchMock = vi
      .fn()
      .mockResolvedValue(jsonResponse({ data: [{ id: 'llama3.2' }, { id: 'qwen3' }] }))
    vi.stubGlobal('fetch', fetchMock)
    const result = await listOpenAiCompatibleModels('http://localhost:11434/v1/')
    expect(result).toEqual({ ok: true, models: ['llama3.2', 'qwen3'] })
    const [url, init] = fetchMock.mock.calls[0] as [string, RequestInit]
    expect(url).toBe('http://localhost:11434/v1/models')
    expect(init.method).toBe('GET')
    const headers = init.headers as Record<string, string>
    expect(Object.keys(headers).map((h) => h.toLowerCase())).not.toContain('authorization')
  })

  it('sends a trimmed bearer key when one is configured', async () => {
    const fetchMock = vi.fn().mockResolvedValue(jsonResponse({ data: [{ id: 'm' }] }))
    vi.stubGlobal('fetch', fetchMock)
    await listOpenAiCompatibleModels('https://gateway.internal/v1', '  sk-secret  ')
    const headers = (fetchMock.mock.calls[0] as [string, RequestInit])[1].headers as Record<
      string,
      string
    >
    expect(headers.Authorization).toBe('Bearer sk-secret')
  })

  it('reports an HTTP error status', async () => {
    vi.stubGlobal('fetch', vi.fn().mockResolvedValue(new Response('nope', { status: 404 })))
    const result = await listOpenAiCompatibleModels('http://localhost:8080/v1')
    expect(result).toEqual({ ok: false, error: 'HTTP 404 from http://localhost:8080/v1/models' })
  })

  it('reports a non-JSON body', async () => {
    vi.stubGlobal('fetch', vi.fn().mockResolvedValue(new Response('<html>hi</html>')))
    const result = await listOpenAiCompatibleModels('http://localhost:8080/v1')
    expect(result.ok).toBe(false)
    if (!result.ok) expect(result.error).toContain('did not return JSON')
  })

  it('reports a server that lists no models', async () => {
    vi.stubGlobal('fetch', vi.fn().mockResolvedValue(jsonResponse({ data: [] })))
    const result = await listOpenAiCompatibleModels('http://localhost:8080/v1')
    expect(result).toEqual({ ok: false, error: 'The server reported no models' })
  })

  it('turns a connection failure into ok:false instead of throwing', async () => {
    vi.stubGlobal(
      'fetch',
      vi.fn().mockRejectedValue(new Error('connect ECONNREFUSED 127.0.0.1:11434')),
    )
    const result = await listOpenAiCompatibleModels('http://localhost:11434/v1')
    expect(result).toEqual({ ok: false, error: 'connect ECONNREFUSED 127.0.0.1:11434' })
  })

  it('rejects an empty or malformed base URL without touching the network', async () => {
    const fetchMock = vi.fn()
    vi.stubGlobal('fetch', fetchMock)
    expect((await listOpenAiCompatibleModels('   ')).ok).toBe(false)
    expect(fetchMock).not.toHaveBeenCalled()
  })

  it('catches a Base URL typed without http:// (which `new URL` would otherwise accept)', async () => {
    const fetchMock = vi.fn()
    vi.stubGlobal('fetch', fetchMock)
    const result = await listOpenAiCompatibleModels('localhost:11434/v1')
    expect(result.ok).toBe(false)
    if (!result.ok) expect(result.error).toContain('http://')
    expect(fetchMock).not.toHaveBeenCalled()
  })

  it('gives up on a server that never answers', async () => {
    vi.useFakeTimers()
    vi.stubGlobal(
      'fetch',
      vi.fn().mockImplementation(
        (_url: string, init: RequestInit) =>
          new Promise((_, reject) => {
            init.signal!.addEventListener('abort', () => reject(new Error('aborted')), {
              once: true,
            })
          }),
      ),
    )
    const pending = listOpenAiCompatibleModels('http://10.0.0.9:8000/v1')
    await vi.advanceTimersByTimeAsync(10_000)
    const result = await pending
    expect(result.ok).toBe(false)
    if (!result.ok) expect(result.error).toContain('No response from')
  })
})

describe('providerRequiresApiKey', () => {
  it('is false for the providers that can run without a key', () => {
    expect(providerRequiresApiKey('custom')).toBe(false)
    expect(providerRequiresApiKey('codex')).toBe(false)
  })

  it('is true for every hosted vendor, and for unknown ids', () => {
    for (const id of [
      'genspark',
      'anthropic',
      'openai',
      'gemini',
      'deepseek',
      'openrouter',
    ] as const) {
      expect(providerRequiresApiKey(id)).toBe(true)
    }
    expect(providerRequiresApiKey('nonsense' as never)).toBe(true)
  })

  it('agrees with activeProvider: a key-less custom setup is selectable, a key-less vendor is not', () => {
    const settings = defaultAiSettings()
    settings.provider = 'custom'
    settings.providers.custom = {
      apiKey: '',
      model: 'qwen3:8b',
      baseUrl: 'http://localhost:11434/v1',
    }
    expect(activeProvider(settings)).toBe('custom')
    expect(providerRequiresApiKey(activeProvider(settings))).toBe(false)

    settings.provider = 'openai'
    settings.providers.openai = { ...settings.providers.openai, apiKey: '' }
    expect(activeProvider(settings)).toBe('genspark')
  })
})

describe('local endpoints get the long watchdog budget', () => {
  /** a fetch that never answers and records the signal it was given */
  function silentFetch() {
    const seen: { signal?: AbortSignal } = {}
    vi.stubGlobal(
      'fetch',
      vi.fn().mockImplementation((_url: string, init: RequestInit) => {
        seen.signal = init.signal as AbortSignal
        return new Promise((_, reject) => {
          seen.signal!.addEventListener('abort', () => reject(new Error('aborted')), { once: true })
        })
      }),
    )
    return seen
  }

  const cb = () => ({
    signal: new AbortController().signal,
    onDelta: () => {},
    onToolCall: () => {},
  })

  const stream = (baseUrl: string) =>
    streamForProvider(
      'custom',
      { apiKey: '', model: 'qwen3:8b', baseUrl },
      'system',
      [{ role: 'user', text: 'hi' }],
      [],
      100,
      cb(),
    )

  beforeEach(() => {
    vi.useFakeTimers()
  })

  it('a hosted custom endpoint still times out after the cloud connect budget', async () => {
    const seen = silentFetch()
    const run = expect(stream('https://gateway.example.com/v1')).rejects.toBeInstanceOf(
      AiTimeoutError,
    )
    await vi.advanceTimersByTimeAsync(AI_CONNECT_TIMEOUT_MS)
    expect(seen.signal!.aborted).toBe(true)
    await run
  })

  it('a localhost stream is still waiting past the cloud budget and only aborts at the local one', async () => {
    const seen = silentFetch()
    const run = expect(stream('http://localhost:11434/v1')).rejects.toBeInstanceOf(AiTimeoutError)
    await vi.advanceTimersByTimeAsync(AI_CONNECT_TIMEOUT_MS + 1_000)
    expect(seen.signal!.aborted).toBe(false)
    await vi.advanceTimersByTimeAsync(AI_LOCAL_CONNECT_TIMEOUT_MS)
    expect(seen.signal!.aborted).toBe(true)
    await run
  })

  it('a LAN server gets the same patience for one-shot chat', async () => {
    const seen = silentFetch()
    const run = chatForProvider(
      'custom',
      { apiKey: '', model: 'qwen3:8b', baseUrl: 'http://192.168.1.20:8000/v1' },
      'system',
      'ping',
    )
    const settled = expect(run).rejects.toBeInstanceOf(AiTimeoutError)
    await vi.advanceTimersByTimeAsync(AI_CHAT_RESPONSE_TIMEOUT_MS + 1_000)
    expect(seen.signal!.aborted).toBe(false)
    await vi.advanceTimersByTimeAsync(AI_LOCAL_CONNECT_TIMEOUT_MS)
    expect(seen.signal!.aborted).toBe(true)
    await settled
  })

  it('does not send an Authorization header for a key-less local server', async () => {
    const fetchMock = vi
      .fn()
      .mockResolvedValue(jsonResponse({ choices: [{ message: { content: 'OK' } }] }))
    vi.stubGlobal('fetch', fetchMock)
    const result = await chatForProvider(
      'custom',
      { apiKey: '', model: 'qwen3:8b', baseUrl: 'http://localhost:11434/v1' },
      'system',
      'ping',
    )
    expect(result).toEqual({ ok: true, content: 'OK' })
    const headers = (fetchMock.mock.calls[0] as [string, RequestInit])[1].headers as Record<
      string,
      string
    >
    expect(Object.keys(headers).map((h) => h.toLowerCase())).not.toContain('authorization')
  })
})
