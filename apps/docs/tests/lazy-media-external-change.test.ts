import { afterEach, describe, expect, it, vi } from 'vitest'
import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import JSZip from 'jszip'
import { lazyMediaUrl } from '@genoffice/docx-engine/lazy-media'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'

vi.mock('electron', () => ({ protocol: { handle: vi.fn() } }))
import {
  forgetLazyMediaOwner,
  materializeLazyDocx,
  openLazyDocx,
  pointLazyMediaAt,
  readLazyMedia,
} from '../src/main/lazy-media'

let root: string | undefined
afterEach(async () => {
  forgetLazyMediaOwner(9101)
  if (root) await rm(root, { recursive: true, force: true })
})

describe('lazy document source preservation', () => {
  it('keeps images from the opened/saved version after an external rewrite and in Save a Copy', async () => {
    root = await mkdtemp(join(tmpdir(), 'docx-lazy-external-'))
    const path = join(root, 'source.docx')
    const zip = await JSZip.loadAsync(
      await buildDocx({ bodyXml: '<w:p><w:r><w:t>Opened text</w:t></w:r></w:p>' }),
    )
    const image = Buffer.alloc(16 * 1024 * 1024, 0x41)
    zip.file('word/media/image1.png', image)
    const original = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
    await writeFile(path, original)
    const lazy = await openLazyDocx(path, 9101)
    expect(lazy).not.toBeNull()
    pointLazyMediaAt(new Set([lazy!.hash]), path, 9101, original)
    zip.file('word/media/image1.png', Buffer.alloc(image.length, 0x42))
    await writeFile(path, await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }))
    const displayed = await readLazyMedia(lazyMediaUrl(lazy!.hash, 'word/media/image1.png'))
    expect(displayed?.body.equals(image)).toBe(true)
    const copy = await materializeLazyDocx(lazy!.bytes)
    const reopened = await JSZip.loadAsync(copy.bytes)
    expect((await reopened.file('word/media/image1.png')!.async('nodebuffer')).equals(image)).toBe(
      true,
    )
  })
})
