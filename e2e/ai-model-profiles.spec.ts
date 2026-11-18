import { expect, test, type Page } from '@playwright/test'
import { createServer } from 'node:http'
import { once } from 'node:events'
import { readFile } from 'node:fs/promises'
import { join } from 'node:path'
import { closeAndSaveVideo, launchShell, waitForPageWithUrl } from './helpers'

test('saved profiles immediately update an open chat and persist across restart', async () => {
  test.setTimeout(120_000)
  const requests: Array<{ model: string; authorization: string; url: string }> = []
  const server = createServer(async (request, response) => {
    response.setHeader('Content-Type', 'application/json')
    if (request.url?.endsWith('/models')) {
      response.end(JSON.stringify({ data: [{ id: 'model-a' }, { id: 'model-b' }] }))
      return
    }
    let body = ''
    for await (const chunk of request) body += chunk
    const payload = JSON.parse(body || '{}')
    requests.push({
      model: payload.model,
      authorization: request.headers.authorization ?? '',
      url: request.url ?? '',
    })
    response.end(
      JSON.stringify({
        choices: [
          {
            message: {
              role: 'assistant',
              content: payload.stream ? `Reply from ${payload.model}` : 'OK',
            },
            finish_reason: 'stop',
          },
        ],
      }),
    )
  })
  server.listen(0, '127.0.0.1')
  await once(server, 'listening')
  const port = (server.address() as { port: number }).port
  const base = `http://127.0.0.1:${port}`
  const launched = await launchShell({ onboardingSeen: true, videoDir: 'ai-model-profiles' })
  const openSettings = async (page: Page) => {
    await page.locator('.tab-item.tab-home').click()
    await page.getByRole('button', { name: 'Settings', exact: true }).click()
  }
  try {
    await launched.page.evaluate(async (url) => {
      const settings = await window.aiOffice.getAiSettings()
      settings.provider = 'custom'
      settings.providers.custom = { apiKey: 'test-key-a', model: 'model-a', baseUrl: `${url}/a/v1` }
      await window.aiOffice.setAiSettings(settings)
    }, base)
    await launched.page.locator('.quick-card', { hasText: 'AI Docs' }).click()
    const docs = await waitForPageWithUrl(launched.app, '://docs/')
    const composer = docs.locator('.ai-composer textarea')
    await composer.fill('Preserve this chat draft')

    await openSettings(launched.page)
    await launched.page.getByRole('button', { name: 'New profile', exact: true }).click()
    await launched.page.locator('#set-ai-profile-name').fill('Work server')
    await launched.page.locator('#set-ai-model').fill('model-b')
    await launched.page.locator('#set-ai-key').fill('test-key-b')
    await launched.page.locator('#set-ai-base-url').fill(`${base}/b/v1`)
    await launched.page.getByRole('button', { name: 'Test connection', exact: true }).click()
    await expect(launched.page.locator('.set-pane-footer')).toContainText('Connection OK')
    await launched.page.getByRole('button', { name: 'Save and use', exact: true }).click()
    await expect(launched.page.locator('.set-ai-profile-row.selected')).toContainText('Active')
    await launched.page.locator('.set-close').click()
    await launched.page.locator('.tab-item', { hasText: 'Untitled Document' }).click()
    await expect(composer).toHaveValue('Preserve this chat draft')
    const picker = docs.getByRole('combobox', { name: 'Model profiles', exact: true })
    await expect(picker.locator('option:checked')).toHaveText('Work server')
    await composer.fill('Which model is responding?')
    await docs.getByRole('button', { name: 'Send', exact: true }).click()
    await expect(docs.locator('.ai-msg')).toContainText([
      'Which model is responding?',
      'Reply from model-b',
    ])
    expect(requests.at(-1)).toEqual({
      model: 'model-b',
      authorization: 'Bearer test-key-b',
      url: '/b/v1/chat/completions',
    })

    await picker.selectOption('default')
    await composer.fill('Use the original profile')
    await docs.getByRole('button', { name: 'Send', exact: true }).click()
    await expect(docs.locator('.ai-msg').last()).toContainText('Reply from model-a')
    expect(requests.at(-1)).toEqual({
      model: 'model-a',
      authorization: 'Bearer test-key-a',
      url: '/a/v1/chat/completions',
    })
    const stored = JSON.parse(
      await readFile(join(launched.userDataDir, 'ai-settings.json'), 'utf8'),
    )
    expect(stored.modelProfiles).toHaveLength(2)
    expect(stored.activeModelProfileId).toBe('default')
    await docs.screenshot({ path: '/private/tmp/genoffice-ai-profiles-chat.png' })
  } finally {
    await closeAndSaveVideo(launched, 'ai-model-profiles')
    server.close()
  }
  const restarted = await launchShell({
    userDataDir: launched.userDataDir,
    videoDir: 'ai-model-profiles-restart',
  })
  try {
    await openSettings(restarted.page)
    await expect(restarted.page.locator('.set-ai-profile-row')).toHaveCount(2)
    await expect(restarted.page.locator('.set-ai-profile-row.selected')).toContainText('Default')
    await restarted.page.screenshot({ path: '/private/tmp/genoffice-ai-profiles-settings.png' })
    await restarted.page.locator('.set-ai-profile-row', { hasText: 'Work server' }).click()
    await restarted.page.locator('#set-ai-profile-name').fill('Renamed server')
    await restarted.page.getByRole('button', { name: 'Save and use', exact: true }).click()
    await expect(restarted.page.locator('.set-ai-profile-row.selected')).toContainText(
      'Renamed server',
    )
    restarted.page.once('dialog', (dialog) => dialog.accept())
    await restarted.page.getByRole('button', { name: 'Delete profile', exact: true }).click()
    await expect(restarted.page.locator('.set-ai-profile-row')).toHaveCount(1)
    await expect(restarted.page.locator('.set-ai-profile-row.selected')).toContainText('Default')
  } finally {
    await closeAndSaveVideo(restarted, 'ai-model-profiles-restart')
  }
})

test('all open editor chats use a newly selected profile on their next question', async () => {
  test.setTimeout(150_000)
  const requests: Array<{ model: string; messages: Array<{ role: string; content: unknown }> }> = []
  const server = createServer(async (request, response) => {
    let body = ''
    for await (const chunk of request) body += chunk
    const payload = JSON.parse(body || '{}')
    requests.push(payload)
    response.setHeader('Content-Type', 'application/json')
    response.end(
      JSON.stringify({
        choices: [
          { message: { role: 'assistant', content: 'Profile B response' }, finish_reason: 'stop' },
        ],
      }),
    )
  })
  server.listen(0, '127.0.0.1')
  await once(server, 'listening')
  const url = `http://127.0.0.1:${(server.address() as { port: number }).port}/v1`
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: 'ai-model-profiles-editors',
  })
  try {
    await launched.page.evaluate(async (baseUrl) => {
      const settings = await window.aiOffice.getAiSettings()
      settings.provider = 'custom'
      settings.providers.custom = { apiKey: 'test-a', model: 'model-a', baseUrl }
      await window.aiOffice.setAiSettings(settings)
    }, url)
    const editors: Array<{ name: string; page: Page; tab: number }> = []
    for (const [index, name] of ['Docs', 'Sheets', 'Slides', 'Markdown', 'HTML', 'PDF'].entries()) {
      await launched.page.locator('.tab-item.tab-home').click()
      await launched.page.locator('.quick-card', { hasText: `AI ${name}` }).click()
      const page = await waitForPageWithUrl(launched.app, `://${name.toLowerCase()}/`, 30_000)
      await expect(page.getByRole('combobox', { name: 'Model profiles', exact: true })).toBeVisible(
        { timeout: 30_000 },
      )
      editors.push({ name, page, tab: index + 1 })
    }
    await launched.page.evaluate(async () => {
      const settings = await window.aiModelSettings!.get()
      settings.providers.custom = {
        ...settings.providers.custom,
        model: 'model-b',
        apiKey: 'test-b',
      }
      await window.aiModelSettings!.saveProfile(settings, 'b', 'Profile B')
    })
    for (const { name, page, tab } of editors) {
      await launched.page.locator('.tab-item').nth(tab).click()
      await expect(page.getByRole('combobox', { name: 'Model profiles', exact: true })).toHaveValue(
        'b',
      )
      const instruction = `Just answer this question for ${name}. Do not edit the document.`
      await page.locator('.ai-composer textarea, [data-slides-ai-input]').fill(instruction)
      await page.getByRole('button', { name: 'Send', exact: true }).click()
      await expect
        .poll(
          () =>
            requests.find((r) =>
              r.messages?.some(
                (m) => m.role === 'user' && JSON.stringify(m.content).includes(instruction),
              ),
            )?.model,
          { timeout: 30_000 },
        )
        .toBe('model-b')
      await expect(page.getByText('Profile B response', { exact: true }).first()).toBeVisible()
    }
  } finally {
    await closeAndSaveVideo(launched, 'ai-model-profiles-editors')
    server.close()
  }
})
