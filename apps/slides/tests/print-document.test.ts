import { readFile, stat } from 'node:fs/promises'
import { dirname } from 'node:path'
import { describe, expect, it } from 'vitest'
import { loadPrintHtmlFile, printErrorMessage } from '../src/main/print-document'

describe('large native print document', () => {
  it('loads multi-megabyte image HTML from a private file and removes it after printing', async () => {
    const html = `<html><img src="data:image/png;base64,${'A'.repeat(3_000_000)}"></html>`
    let loadedPath = ''
    const cleanup = await loadPrintHtmlFile(
      {
        async loadFile(path) {
          loadedPath = path
          expect(await readFile(path, 'utf8')).toBe(html)
        },
      },
      html,
    )
    expect(loadedPath).toMatch(/genoffice-slides-print-.*\/print\.html$/)
    expect((await stat(loadedPath)).mode & 0o777).toBe(0o600)
    await cleanup()
    await expect(stat(dirname(loadedPath))).rejects.toThrow()
  })

  it('cleans its temporary file if Chromium rejects the page', async () => {
    let loadedPath = ''
    await expect(
      loadPrintHtmlFile(
        {
          async loadFile(path) {
            loadedPath = path
            throw new Error('page failed')
          },
        },
        '<html>test</html>',
      ),
    ).rejects.toThrow('page failed')
    await expect(stat(dirname(loadedPath))).rejects.toThrow()
  })

  it('hides embedded image bytes and bounds native error text', () => {
    const error = `ERR_INVALID_URL loading 'data:text/html;base64,${'A'.repeat(3_000_000)}'`
    const message = printErrorMessage(error)
    expect(message).toContain('ERR_INVALID_URL')
    expect(message).toContain('[embedded document]')
    expect(message).not.toContain('AAAA')
    expect(message.length).toBeLessThanOrEqual(512)
  })
})
