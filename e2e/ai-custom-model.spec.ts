import { expect, test } from '@playwright/test'
import { createServer } from 'node:http'
import { once } from 'node:events'
import { closeAndSaveVideo, launchShell, waitForPageWithUrl } from './helpers'

test('OpenAI accepts a custom model and gateway endpoint, including after reopening settings', async () => {
  test.setTimeout(90_000)
  const requests: Array<{ model: string; path: string; authorization: string }> = []
  const server = createServer(async (request, response) => {
    let body = ''
    for await (const chunk of request) body += chunk
    const payload = JSON.parse(body || '{}')
    requests.push({
      model: payload.model,
      path: request.url ?? '',
      authorization: request.headers.authorization ?? '',
    })
    response.setHeader('Content-Type', 'application/json')
    response.end(
      JSON.stringify({
        choices: [
          {
            message: { role: 'assistant', content: payload.stream ? 'Gateway reply' : 'OK' },
            finish_reason: 'stop',
          },
        ],
      }),
    )
  })
  server.listen(0, '127.0.0.1')
  await once(server, 'listening')
  const baseUrl = `http://127.0.0.1:${(server.address() as { port: number }).port}/gateway/v1`
  const launched = await launchShell({ onboardingSeen: true, videoDir: 'ai-custom-model' })
  const home = launched.page
  try {
    await home.locator('.quick-card', { hasText: 'AI Docs' }).click()
    const docs = await waitForPageWithUrl(launched.app, '://docs/')
    await expect(docs.locator('.ai-composer textarea')).toBeVisible()
    await home.locator('.tab-item.tab-home').click()
    await home.getByRole('button', { name: 'Settings', exact: true }).click()
    await home.getByRole('button', { name: 'Provider', exact: true }).click()
    await home.getByRole('option', { name: 'OpenAI', exact: true }).click()
    await home.getByRole('button', { name: 'Model', exact: true }).click()
    await home.getByRole('option', { name: 'Custom', exact: true }).click()
    expect((await home.locator('#set-ai-model').boundingBox())!.width).toBeGreaterThanOrEqual(200)
    await home.locator('#set-ai-model').fill('gateway-only-model')
    await home.locator('#set-ai-key').fill('gateway-test-key')
    await home.locator('#set-ai-base-url').fill(baseUrl)
    await home.getByRole('button', { name: 'Test connection', exact: true }).click()
    await expect(home.locator('.set-pane-footer')).toContainText('Connection OK')
    expect(requests.at(-1)).toEqual({
      model: 'gateway-only-model',
      path: '/gateway/v1/chat/completions',
      authorization: 'Bearer gateway-test-key',
    })
    await home.getByRole('button', { name: 'Save and use', exact: true }).click()
    await expect(home.locator('.set-pane-footer')).toContainText('Saved')
    await home.locator('.set-close').click()
    await home.locator('.tab-item', { hasText: 'Untitled Document' }).click()
    await docs
      .locator('.ai-composer textarea')
      .fill('Answer this question without editing the document.')
    await docs.getByRole('button', { name: 'Send', exact: true }).click()
    await expect(docs.getByText('Gateway reply', { exact: true })).toBeVisible()
    expect(requests.at(-1)?.model).toBe('gateway-only-model')

    await home.locator('.tab-item.tab-home').click()
    await home.getByRole('button', { name: 'Settings', exact: true }).click()
    await expect(home.getByRole('button', { name: 'Provider', exact: true })).toContainText(
      'OpenAI',
    )
    await expect(home.getByRole('button', { name: 'Model', exact: true })).toContainText('Custom')
    await expect(home.locator('#set-ai-model')).toHaveValue('gateway-only-model')
    await expect(home.locator('#set-ai-base-url')).toHaveValue(baseUrl)
    const builtIn = await home.evaluate(
      () =>
        window.aiOffice.getAiProviders().find((provider) => provider.id === 'openai')!.models[0]!,
    )
    await home.getByRole('button', { name: 'Model', exact: true }).click()
    await home.getByRole('option', { name: builtIn, exact: true }).click()
    await expect(home.locator('#set-ai-model')).toHaveCount(0)
    await expect(home.locator('#set-ai-base-url')).toHaveValue(baseUrl)
    // A listed model can also be used through the explicit custom mode.
    await home.getByRole('button', { name: 'Model', exact: true }).click()
    await home.getByRole('option', { name: 'Custom', exact: true }).click()
    await expect(home.locator('#set-ai-model')).toHaveValue(builtIn)
    await home.getByRole('button', { name: 'Save and use', exact: true }).click()
    await expect(home.locator('.set-pane-footer')).toContainText('Saved')
    await home.locator('.set-close').click()
    await home.getByRole('button', { name: 'Settings', exact: true }).click()
    await expect(home.getByRole('button', { name: 'Model', exact: true })).toContainText('Custom')
    await expect(home.locator('#set-ai-model')).toHaveValue(builtIn)
    await home.screenshot({ path: '/private/tmp/genoffice-ai-custom-model.png' })
  } finally {
    await closeAndSaveVideo(launched, 'ai-custom-model')
    server.close()
  }
})
