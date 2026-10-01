import { afterEach, describe, expect, it, vi } from 'vitest'
import { mkdtempSync, writeFileSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { defaultAiSettings } from '@genoffice/ai-provider'
import { generateImageTool, analyzeMediaTool } from '../src/media-tools'

vi.mock('@genoffice/electron-utils/generated-images', () => ({
  storeGeneratedImage: vi.fn(() => 'file:///generated/image.png'),
  readGeneratedImage: vi.fn(),
}))

const dirs: string[] = []
afterEach(() => {
  vi.unstubAllGlobals()
  for (const dir of dirs.splice(0)) rmSync(dir, { recursive: true, force: true })
})

describe('media tools use configured providers', () => {
  it('reports missing image configuration without making a request', async () => {
    const fetch = vi.fn()
    vi.stubGlobal('fetch', fetch)
    const result = await generateImageTool('/nonexistent/settings.json', { prompt: 'a flower' })
    expect(result.error).toMatch(/Settings > AI Media/)
    expect(fetch).not.toHaveBeenCalled()
  })

  it('reports missing analysis configuration before downloading user media', async () => {
    const fetch = vi.fn()
    vi.stubGlobal('fetch', fetch)
    const result = await analyzeMediaTool('/nonexistent/settings.json', {
      mediaUrls: ['https://example.com/private.png'],
      requirements: 'Describe it',
    })
    expect(result.error).toMatch(/Settings > AI Media/)
    expect(fetch).not.toHaveBeenCalled()
  })

  it('uses the selected image endpoint and forwards transparency without an account login', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'genoffice-media-test-'))
    dirs.push(dir)
    const settings = defaultAiSettings()
    settings.media!.imageProvider = 'openai'
    settings.media!.providers.openai.apiKey = 'test-key'
    settings.media!.providers.openai.baseUrl = 'https://image.example/v1'
    const path = join(dir, 'ai-settings.json')
    writeFileSync(path, JSON.stringify(settings))
    const fetch = vi.fn(
      async (_url: string, _init?: RequestInit) =>
        new Response(JSON.stringify({ data: [{ b64_json: 'iVBORw0KGgoAAAAA' }] }), {
          status: 200,
          headers: { 'content-type': 'application/json' },
        }),
    )
    vi.stubGlobal('fetch', fetch)
    expect(
      await generateImageTool(path, { prompt: 'a flower', transparentBackground: true }),
    ).toEqual({ url: 'file:///generated/image.png' })
    expect(fetch).toHaveBeenCalledTimes(1)
    expect(fetch.mock.calls[0]?.[0]).toBe('https://image.example/v1/images/generations')
    const init = (fetch.mock.calls as unknown as [string, RequestInit][])[0][1]
    expect(JSON.parse(String(init.body)).background).toBe('transparent')
  })
})
