import { afterEach, describe, expect, it, vi } from 'vitest'
import { mkdtemp, readFile, rename, rm, utimes, writeFile } from 'node:fs/promises'
import { join } from 'node:path'
import { tmpdir } from 'node:os'
import {
  ExternalFileChangedError,
  ExternalFileTracker,
  fileContentHash,
} from '../src/external-file-tracker'

const roots: string[] = []
afterEach(async () => {
  await Promise.all(roots.splice(0).map((path) => rm(path, { recursive: true, force: true })))
})
async function setup() {
  const root = await mkdtemp(join(tmpdir(), 'external-file-tracker-'))
  roots.push(root)
  const path = join(root, 'document.txt')
  await writeFile(path, 'Opened version')
  const tracker = new ExternalFileTracker()
  const notice = vi.fn(async () => {})
  await tracker.remember(1, path, fileContentHash('Opened version'), notice)
  await tracker.check(1)
  return { root, path, tracker, notice }
}

describe('external file persistence boundary', () => {
  it('notifies on atomic replacement, once per external version, and keeps blocking after dismissal', async () => {
    const { root, path, tracker, notice } = await setup()
    const temp = join(root, 'replacement')
    await writeFile(temp, 'External version')
    await rename(temp, path)
    expect(await tracker.check(1)).toBe(true)
    expect(notice).toHaveBeenCalledTimes(1)
    await tracker.check(1, true)
    expect(notice).toHaveBeenCalledTimes(1)
    const writer = vi.fn(() => writeFile(path, 'Local edit'))
    await expect(tracker.write(1, path, writer)).rejects.toBeInstanceOf(ExternalFileChangedError)
    expect(writer).not.toHaveBeenCalled()
    expect(await readFile(path, 'utf8')).toBe('External version')
    await writeFile(path, 'Another external version')
    await tracker.check(1)
    expect(notice).toHaveBeenCalledTimes(2)
  })

  it('allows Save a Copy without granting permission to overwrite the external original', async () => {
    const { root, path, tracker } = await setup()
    await writeFile(path, 'External version')
    const copy = join(root, 'copy.txt')
    await tracker.write(1, copy, () => writeFile(copy, 'Local edit'))
    expect(await readFile(copy, 'utf8')).toBe('Local edit')
    expect(tracker.path(1)).toBe(path)
    await expect(tracker.assertWritable(1, path)).rejects.toBeInstanceOf(ExternalFileChangedError)
  })

  it('advances after a successful own save without notifying, and detects the next external edit', async () => {
    const { path, tracker, notice } = await setup()
    await tracker.write(1, path, async () => {
      await writeFile(path, 'Saved local edit')
      await tracker.check(1, true)
    })
    await tracker.check(1)
    expect(notice).not.toHaveBeenCalled()
    await tracker.assertWritable(1, path)
    await writeFile(path, 'External after save')
    await tracker.check(1)
    expect(notice).toHaveBeenCalledTimes(1)
  })

  it('ignores touch but catches changes made during opening and before saving', async () => {
    const { path, tracker, notice } = await setup()
    await utimes(path, new Date(), new Date(Date.now() + 5000))
    await tracker.check(1)
    expect(notice).not.toHaveBeenCalled()
    await writeFile(path, 'External version')
    await tracker.remember(1, path, fileContentHash('Opened version'), notice)
    await tracker.check(1)
    expect(notice).toHaveBeenCalledTimes(1)
    await expect(tracker.assertWritable(1, path)).rejects.toBeInstanceOf(ExternalFileChangedError)
  })

  it('does not acknowledge a failed save or a source deletion, and stops after teardown', async () => {
    const { path, tracker, notice } = await setup()
    await expect(
      tracker.write(1, path, async () => {
        throw new Error('Write failed')
      }),
    ).rejects.toThrow('Write failed')
    await rm(path)
    await tracker.check(1)
    expect(notice).toHaveBeenCalledWith(true)
    await expect(tracker.assertWritable(1, path)).rejects.toBeInstanceOf(ExternalFileChangedError)
    tracker.forget(1)
    await writeFile(path, 'Recreated externally')
    await tracker.check(1)
    expect(notice).toHaveBeenCalledTimes(1)
    expect(tracker.ids()).toEqual([])
  })
})
