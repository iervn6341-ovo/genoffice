/**
 * Saving promotes a same-directory temp file over the target with rename.
 * Windows AV/indexer/cloud-sync locks make that rename fail transiently with
 * EPERM/EACCES/EBUSY (user report: "EPERM: operation not permitted, rename
 * .tmp.xlsx → …"), so the promotion retries, falls back to an in-place copy,
 * and surfaces a stable, localizable message when the target stays locked.
 */
import {
  chmod,
  copyFile,
  mkdir,
  mkdtemp,
  readFile,
  rename,
  rm,
  stat,
  truncate,
  writeFile,
} from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { basename, join } from 'node:path'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { promoteFileAtomically } from '@genoffice/xlsx-gateway/gateway/xlsx-package-io'

vi.mock('node:fs/promises', async (importOriginal) => {
  const actual = await importOriginal<typeof import('node:fs/promises')>()
  return { ...actual, copyFile: vi.fn(actual.copyFile), rename: vi.fn(actual.rename) }
})
const copyFileMock = vi.mocked(copyFile)

const scratches: string[] = []
const actualCopyFile = copyFileMock.getMockImplementation()!
const renameMock = vi.mocked(rename)
const actualRename = renameMock.getMockImplementation()!
const busy = () => Object.assign(new Error('EBUSY: resource busy'), { code: 'EBUSY' })

async function scratchDir(): Promise<string> {
  const dir = await mkdtemp(join(tmpdir(), 'genoffice-promote-test-'))
  scratches.push(dir)
  return dir
}

afterEach(async () => {
  renameMock.mockReset().mockImplementation(actualRename)
  copyFileMock.mockReset()
  copyFileMock.mockImplementation(actualCopyFile)
  for (const dir of scratches.splice(0)) {
    await chmod(dir, 0o755).catch(() => {})
    for (const sub of ['locked']) await chmod(join(dir, sub), 0o755).catch(() => {})
    await rm(dir, { recursive: true, force: true })
  }
})

describe('promoteFileAtomically', () => {
  it('replaces the target and removes the temp file', async () => {
    const dir = await scratchDir()
    const temporary = join(dir, '.new.tmp.xlsx')
    const target = join(dir, 'book.xlsx')
    await writeFile(temporary, 'new-bytes')
    await writeFile(target, 'old-bytes')
    await promoteFileAtomically(temporary, target)
    expect(await readFile(target, 'utf8')).toBe('new-bytes')
    await expect(stat(temporary)).rejects.toMatchObject({ code: 'ENOENT' })
  })

  it('falls back to an in-place copy when only the rename is blocked', async () => {
    const dir = await scratchDir()
    const locked = join(dir, 'locked')
    await mkdir(locked)
    const temporary = join(locked, '.new.tmp.xlsx')
    const target = join(locked, 'book.xlsx')
    await writeFile(temporary, 'new-bytes')
    await writeFile(target, 'old-bytes')
    // Inject a transient-style lock on every OS while keeping the copy real.
    renameMock.mockRejectedValue(busy())
    await promoteFileAtomically(temporary, target)
    await chmod(locked, 0o755)
    expect(await readFile(target, 'utf8')).toBe('new-bytes')
  }, 15_000)

  it('reports a persistently locked target with a stable localizable message', async () => {
    const dir = await scratchDir()
    const locked = join(dir, 'locked')
    await mkdir(locked)
    const temporary = join(locked, '.new.tmp.xlsx')
    const target = join(locked, 'book.xlsx')
    await writeFile(temporary, 'new-bytes')
    await writeFile(target, 'old-bytes')
    // rename AND in-place copy both refused — the Excel-holds-the-file case
    copyFileMock.mockImplementation(async (src, dest, mode) => {
      if (String(dest) === target) throw busy()
      return actualCopyFile(src, dest, mode)
    })
    renameMock.mockRejectedValue(busy())
    await expect(promoteFileAtomically(temporary, target)).rejects.toThrow(
      'The save target is locked by another program',
    )
    await chmod(locked, 0o755)
    await chmod(target, 0o644)
    // the finished bytes survive the failure for the caller's cleanup/retry
    expect(await readFile(temporary, 'utf8')).toBe('new-bytes')
    expect(await readFile(target, 'utf8')).toBe('old-bytes')
  }, 15_000)

  it('restores the target when the in-place copy dies after truncating it', async () => {
    const dir = await scratchDir()
    const locked = join(dir, 'locked')
    await mkdir(locked)
    const temporary = join(locked, '.new.tmp.xlsx')
    const target = join(locked, 'book.xlsx')
    await writeFile(temporary, 'new-bytes')
    await writeFile(target, 'old-bytes')
    renameMock.mockRejectedValue(busy())
    // the backup copy runs for real; the copy over the target truncates it
    // and then fails the way a lock acquired mid-write does
    copyFileMock.mockImplementation(async (src, dest, mode) => {
      if (String(src) !== temporary) return actualCopyFile(src, dest, mode)
      await truncate(String(dest))
      throw Object.assign(new Error('EBUSY: resource busy'), { code: 'EBUSY' })
    })
    await expect(promoteFileAtomically(temporary, target)).rejects.toThrow(
      'The save target is locked by another program',
    )
    await chmod(locked, 0o755)
    expect(await readFile(target, 'utf8')).toBe('old-bytes')
    expect(await readFile(temporary, 'utf8')).toBe('new-bytes')
  }, 15_000)

  it('names the surviving backup when the target cannot be restored either', async () => {
    const dir = await scratchDir()
    const locked = join(dir, 'locked')
    await mkdir(locked)
    const temporary = join(locked, '.new.tmp.xlsx')
    const target = join(locked, 'book.xlsx')
    await writeFile(temporary, 'new-bytes')
    await writeFile(target, 'old-bytes')
    renameMock.mockRejectedValue(busy())
    copyFileMock.mockImplementation(async (src, dest, mode) => {
      if (String(src) === temporary) {
        await truncate(String(dest))
        throw busy()
      }
      if (String(dest).startsWith(locked)) throw busy()
      return actualCopyFile(src, dest, mode)
    })
    const failure = await promoteFileAtomically(temporary, target).catch((error: Error) => error)
    await chmod(locked, 0o755)
    expect(failure?.message).toContain('preserved at: ')
    const survivor = failure!.message.split('preserved at: ')[1] ?? ''
    // the injected lock also refuses the recovered copy, so the tmp backup stays
    expect(survivor.startsWith(tmpdir())).toBe(true)
    expect(basename(survivor)).toMatch(/^book\.recovered-[0-9a-f-]+\.xlsx$/)
    expect(await readFile(survivor, 'utf8')).toBe('old-bytes')
    await rm(survivor, { force: true })
  }, 15_000)

  it('propagates non-retryable errors untouched', async () => {
    const dir = await scratchDir()
    const temporary = join(dir, '.missing.tmp.xlsx')
    await expect(promoteFileAtomically(temporary, join(dir, 'book.xlsx'))).rejects.toMatchObject({
      code: 'ENOENT',
    })
  })
})
