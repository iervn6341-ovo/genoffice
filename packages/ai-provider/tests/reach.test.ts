import { afterEach, describe, expect, it, vi } from 'vitest'
import { assertEndpointReachable, describeNetworkError } from '../src/reach'

afterEach(() => {
  vi.unstubAllGlobals()
})

describe('describeNetworkError', () => {
  it('surfaces the network error code hidden behind "fetch failed"', () => {
    const e = new TypeError('fetch failed', { cause: { code: 'EHOSTUNREACH' } })
    expect(describeNetworkError(e)).toBe('fetch failed (EHOSTUNREACH)')
  })
})

describe('assertEndpointReachable', () => {
  it('passes on any HTTP answer, even an error status', async () => {
    vi.stubGlobal('fetch', vi.fn().mockResolvedValue(new Response('nope', { status: 404 })))
    await expect(assertEndpointReachable('http://10.0.0.209:11434/v1')).resolves.toBeUndefined()
  })

  it('names the host and the network error when the connection fails', async () => {
    const e = new TypeError('fetch failed', { cause: { code: 'ECONNREFUSED' } })
    vi.stubGlobal('fetch', vi.fn().mockRejectedValue(e))
    await expect(assertEndpointReachable('http://10.0.0.209:11434/v1')).rejects.toThrow(
      /Cannot reach http:\/\/10\.0\.0\.209:11434\/v1 \(fetch failed \(ECONNREFUSED\)\)/,
    )
  })

  it('gives up on a host that never answers instead of waiting out the long local budget', async () => {
    vi.stubGlobal(
      'fetch',
      vi.fn().mockImplementation(
        (_url: string, init: RequestInit) =>
          new Promise((_, reject) => {
            init.signal?.addEventListener('abort', () => reject(new Error('aborted')))
          }),
      ),
    )
    await expect(
      assertEndpointReachable('http://10.0.0.209:11434/v1', undefined, 50),
    ).rejects.toThrow(/no answer within/)
  })

  it('lets a user cancel through as a cancel, not as an unreachable host', async () => {
    const parent = new AbortController()
    vi.stubGlobal(
      'fetch',
      vi.fn().mockImplementation(
        (_url: string, init: RequestInit) =>
          new Promise((_, reject) => {
            init.signal?.addEventListener('abort', () => reject(new Error('aborted')))
          }),
      ),
    )
    const pending = assertEndpointReachable('http://10.0.0.209:11434/v1', parent.signal)
    parent.abort()
    await expect(pending).rejects.toThrow('aborted')
  })
})
