import { createHash } from 'node:crypto'
import { createReadStream } from 'node:fs'
import { realpath, stat } from 'node:fs/promises'
import { basename, dirname, join, resolve } from 'node:path'

export const fileContentHash = (bytes: string | Uint8Array): string =>
  createHash('sha256').update(bytes).digest('hex')

export async function diskContentHash(path: string): Promise<string> {
  const hash = createHash('sha256')
  for await (const chunk of createReadStream(path)) hash.update(chunk)
  return hash.digest('hex')
}

async function identity(path: string): Promise<string> {
  return realpath(path).catch(async () =>
    join(await realpath(dirname(path)).catch(() => resolve(dirname(path))), basename(path)),
  )
}

interface TrackedFile {
  path: string
  identity: string
  hash: string
  stamp: string
  notified: string | null
  checking: boolean
  writing: boolean
  changed: (missing: boolean) => Promise<void>
}

async function diskStamp(path: string): Promise<string> {
  try {
    const s = await stat(path)
    return `${s.dev}:${s.ino}:${s.size}:${s.mtimeMs}:${s.ctimeMs}`
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code === 'ENOENT') return 'missing'
    throw error
  }
}

export class ExternalFileChangedError extends Error {
  constructor() {
    super('The file changed outside GenOffice. Keep your edits, save a copy, or reload the file.')
    this.name = 'ExternalFileChangedError'
  }
}

/** One persistence baseline per editor. Dismissing a notice never authorizes overwriting. */
export class ExternalFileTracker {
  private files = new Map<number, TrackedFile>()
  private generations = new Map<number, number>()

  async remember(
    id: number,
    path: string,
    hash: string,
    changed: TrackedFile['changed'],
  ): Promise<void> {
    const generation = (this.generations.get(id) ?? 0) + 1
    this.generations.set(id, generation)
    const canonical = await identity(path)
    if (this.generations.get(id) !== generation) return
    this.files.set(id, {
      path,
      identity: canonical,
      hash,
      stamp: '',
      notified: null,
      checking: false,
      writing: false,
      changed,
    })
    // Monitoring is triggered by the owner's foreground focus gate. Registering
    // a background/missing file must not start a check on its own.
  }

  forget(id: number): void {
    this.generations.set(id, (this.generations.get(id) ?? 0) + 1)
    this.files.delete(id)
  }

  ids(): number[] {
    return [...this.files.keys()]
  }
  path(id: number): string | undefined {
    return this.files.get(id)?.path
  }
  hash(id: number): string | undefined {
    return this.files.get(id)?.hash
  }

  async check(id: number, force = false, repeatNotice = false): Promise<boolean> {
    const file = this.files.get(id)
    if (!file || file.writing || file.checking) return false
    file.checking = true
    try {
      const stamp = await diskStamp(file.path)
      if (!force && stamp === file.stamp) return false
      const hash = stamp === 'missing' ? 'missing' : await diskContentHash(file.path)
      if (this.files.get(id) !== file || file.writing) return false
      file.stamp = stamp
      if (hash === file.hash) {
        file.notified = null
        return false
      }
      if (repeatNotice || hash !== file.notified) {
        file.notified = hash
        await file.changed(hash === 'missing')
      }
      return true
    } finally {
      file.checking = false
    }
  }

  /** Always hash before a write, even if a writer preserved the file's timestamps. */
  async assertWritable(id: number, path: string): Promise<void> {
    const file = this.files.get(id)
    if (!file || file.identity !== (await identity(path))) return
    let hash: string
    try {
      hash = await diskContentHash(path)
    } catch {
      hash = 'missing'
    }
    if (hash !== file.hash) throw new ExternalFileChangedError()
  }

  /** Suspend self-notices while writing, then advance only to our successfully written bytes. */
  async write<T>(id: number, path: string, writer: () => Promise<T>): Promise<T> {
    const file = this.files.get(id)
    if (file?.writing) throw new Error('A save is already in progress.')
    if (file) file.writing = true
    try {
      await this.assertWritable(id, path)
      const result = await writer()
      if (file && this.files.get(id) === file && file.identity === (await identity(path))) {
        const hash = await diskContentHash(path)
        await this.remember(id, path, hash, file.changed)
      }
      return result
    } finally {
      if (file) file.writing = false
    }
  }
}
