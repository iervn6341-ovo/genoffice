import { test, expect } from '@playwright/test'
import { createServer } from 'node:http'
import { mkdtemp, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { launchShell, closeAndSaveVideo, waitForPageWithUrl, screenshotPath } from './helpers'

for (const [kind, cardIndex] of [
  ['docs', 0],
  ['sheets', 1],
  ['slides', 2],
  ['markdown', 3],
  ['html', 4],
  ['pdf', 5],
] as const) {
  test(`${kind} Q&A uses the configured endpoint without account login`, async () => {
    const requests: { url: string; model: string; authorization?: string }[] = []
    const server = createServer(async (request, response) => {
      let raw = ''
      for await (const chunk of request) raw += chunk
      const body = JSON.parse(raw)
      requests.push({
        url: request.url ?? '',
        model: body.model,
        authorization: request.headers.authorization,
      })
      response.writeHead(200, { 'Content-Type': 'text/event-stream' })
      response.end(
        'data: ' +
          JSON.stringify({
            choices: [
              { delta: { content: 'Configured provider reply verified.' }, finish_reason: null },
            ],
          }) +
          '\n\ndata: ' +
          JSON.stringify({ choices: [{ delta: {}, finish_reason: 'stop' }] }) +
          '\n\ndata: [DONE]\n\n',
      )
    })
    await new Promise<void>((resolve) => server.listen(0, '127.0.0.1', resolve))
    const address = server.address() as { port: number }
    const dir = await mkdtemp(join(tmpdir(), 'genoffice-provider-test-'))
    await writeFile(
      join(dir, 'ai-settings.json'),
      JSON.stringify({
        provider: 'retired-service',
        providers: {
          custom: {
            apiKey: '',
            model: 'local-test-model',
            baseUrl: `http://127.0.0.1:${address.port}/v1`,
          },
        },
      }),
    )
    const launched = await launchShell({
      userDataDir: dir,
      onboardingSeen: true,
      videoDir: 'provider-independent-ai',
      settings: { defaultSaveDir: dir },
    })
    try {
      const settings = await launched.page.evaluate(() => window.aiOffice.getAiSettings())
      expect(settings.provider).toBe('custom')
      expect(Object.keys(settings.providers)).not.toContain('retired-service')
      await launched.page.getByRole('button', { name: 'Settings', exact: true }).click()
      await expect(launched.page.locator('.set-nav')).not.toContainText(/Account|Cloud Projects/i)
      await expect(launched.page.locator('.set-dialog')).not.toContainText(/genspark/i)
      await launched.page.locator('.set-close').click()
      await launched.page.locator('.quick-card').nth(cardIndex).click()
      const page = await waitForPageWithUrl(launched.app, `://${kind}/`)
      const composer = page.locator('.ai-composer textarea')
      await expect(composer).toBeVisible()
      await composer.fill('Reply with a short greeting.')
      await composer.press('Enter')
      await expect(page.locator('.ai-panel, .copilot')).toContainText(
        'Configured provider reply verified.',
        {
          timeout: 30000,
        },
      )
      expect(requests.length).toBeGreaterThan(0)
      expect(
        requests.every((r) => r.url === '/v1/chat/completions' && r.model === 'local-test-model'),
      ).toBe(true)
      expect(requests.every((r) => !r.authorization)).toBe(true)
      await expect(page.locator('.ai-panel, .copilot')).not.toContainText(/genspark|log in/i)
      await page.screenshot({ path: screenshotPath(`configured-provider-qa-${kind}`) })
    } finally {
      await closeAndSaveVideo(launched, 'provider-independent-ai')
      await new Promise<void>((resolve, reject) => server.close((e) => (e ? reject(e) : resolve())))
    }
  })
}

for (const theme of ['light', 'dark']) {
  for (const [kind, index] of [
    ['docs', 0],
    ['slides', 2],
  ] as const) {
    test(`${kind} File button retains readable text on hover and open (${theme})`, async () => {
      const launched = await launchShell({
        onboardingSeen: true,
        settings: { theme },
        videoDir: 'file-hover',
      })
      try {
        await launched.page.locator('.quick-card').nth(index).click()
        const page = await waitForPageWithUrl(launched.app, `://${kind}/`)
        const button = page.locator('.ribbon-tab-file')
        await expect(button).toBeVisible()
        await page.addStyleTag({ content: '* { transition: none !important; }' })
        const color = () =>
          button.evaluate((el) => ({
            text: getComputedStyle(el).color,
            background: getComputedStyle(el).backgroundColor,
          }))
        const normal = await color()
        await button.hover()
        const hovered = await color()
        expect(hovered.text).toBe(normal.text)
        expect(hovered.text).not.toBe(hovered.background)
        await page.screenshot({ path: screenshotPath(`${kind}-file-hover-${theme}`) })
        await button.click()
        expect((await color()).text).toBe(normal.text)
        await page.keyboard.press('Escape')
      } finally {
        await closeAndSaveVideo(launched, 'file-hover')
      }
    })
  }
}
