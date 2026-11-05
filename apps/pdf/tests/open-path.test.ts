import { mkdtempSync, readdirSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { describe, expect, it, vi } from 'vitest'
import { PDFDocument } from 'pdf-lib'
import { dialog } from 'electron'

/**
 * pdf-main open-path bookkeeping: the renderer calls consume-pending on every
 * mount — including remounts caused by View > Reload (Cmd+R). The path must
 * survive re-consumption or a reload strands the tab on "No file to open".
 */

type IpcHandler = (event: { sender: { id: number } }, ...args: unknown[]) => unknown
const handlers = new Map<string, IpcHandler>()

interface FakeWebContents {
  id: number
  once: ReturnType<typeof vi.fn>
  on: ReturnType<typeof vi.fn>
  setWindowOpenHandler: ReturnType<typeof vi.fn>
  loadURL: ReturnType<typeof vi.fn>
  loadFile: ReturnType<typeof vi.fn>
  send: ReturnType<typeof vi.fn>
  isDestroyed: ReturnType<typeof vi.fn>
  listeners: Map<string, () => void>
}

let nextWcId = 1
let lastWebContents: FakeWebContents

function makeFakeWebContents(): FakeWebContents {
  const listeners = new Map<string, () => void>()
  const wc: FakeWebContents = {
    id: nextWcId++,
    listeners,
    once: vi.fn((event: string, handler: () => void) => {
      listeners.set(event, handler)
    }),
    on: vi.fn((event: string, handler: () => void) => {
      listeners.set(event, handler)
    }),
    setWindowOpenHandler: vi.fn(),
    loadURL: vi.fn(),
    loadFile: vi.fn(),
    send: vi.fn(),
    isDestroyed: vi.fn(() => false),
  }
  lastWebContents = wc
  return wc
}

vi.mock('electron', () => ({
  app: { on: vi.fn(), whenReady: vi.fn(() => new Promise(() => {})) },
  dialog: { showSaveDialog: vi.fn(), showMessageBox: vi.fn() },
  shell: {},
  BrowserWindow: class {
    static fromWebContents() {
      return null
    }
  },
  WebContentsView: class {
    webContents = makeFakeWebContents()
  },
  ipcMain: {
    handle: vi.fn((channel: string, handler: IpcHandler) => {
      handlers.set(channel, handler)
    }),
    on: vi.fn((channel: string, handler: IpcHandler) => {
      handlers.set(channel, handler)
    }),
    removeHandler: vi.fn(),
  },
}))

import { PDF_CHANNELS } from '../src/shared/ipc'
import {
  configurePdfRuntime,
  createPdfView,
  flushPdfSave,
  pdfIsDirty,
  requestPdfClose,
  setPdfMaterializedHook,
} from '../src/main/pdf-main'
import type { CreateDocumentRequest, CreateDocumentResult } from '../src/shared/ipc'

function makePdfFile(): string {
  const dir = mkdtempSync(join(tmpdir(), 'pdf-open-path-'))
  const path = join(dir, 'doc.pdf')
  writeFileSync(path, '%PDF-1.4\n%%EOF\n')
  return path
}

const consume = (wcId: number) =>
  handlers.get(PDF_CHANNELS.consumePending)?.({ sender: { id: wcId } })

describe('pdf open-path lifecycle', () => {
  it('consume-pending returns the path again after a reload (second consume)', () => {
    const path = makePdfFile()
    createPdfView(path)
    const wcId = lastWebContents.id

    expect(consume(wcId)).toBe(path)
    // View > Reload remounts the renderer, which consumes again
    expect(consume(wcId)).toBe(path)
  })

  it('drops the path when the view is destroyed', () => {
    const path = makePdfFile()
    createPdfView(path)
    const wc = lastWebContents

    expect(consume(wc.id)).toBe(path)
    wc.listeners.get('destroyed')?.()
    expect(consume(wc.id)).toBeNull()
  })

  it('clears dirty on did-start-loading so a reload does not keep the close-save prompt', () => {
    const path = makePdfFile()
    createPdfView(path)
    const wc = lastWebContents
    handlers.get(PDF_CHANNELS.dirtyChanged)?.({ sender: { id: wc.id } }, true)
    expect(pdfIsDirty(wc.id)).toBe(true)
    wc.listeners.get('did-start-loading')?.()
    expect(pdfIsDirty(wc.id)).toBe(false)
  })
})

describe('memory-only new PDF', () => {
  it('closes an untouched blank without a save prompt', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'pdf-blank-close-'))
    createPdfView(undefined, { blank: true, saveDir: dir })
    const wc = lastWebContents
    vi.mocked(dialog.showMessageBox).mockClear()
    expect(pdfIsDirty(wc.id)).toBe(false)
    expect(await requestPdfClose(wc as never)).toBe(true)
    expect(dialog.showMessageBox).not.toHaveBeenCalled()
    expect(readdirSync(dir)).toEqual([])
  })

  it('keeps rewritten memory content dirty after renderer reload or clean reports', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'pdf-edited-close-'))
    createPdfView(undefined, { blank: true, saveDir: dir })
    const wc = lastWebContents
    const path = consume(wc.id)
    expect(handlers.get(PDF_CHANNELS.sourceModified)!({ sender: wc })).toBe(false)
    expect(
      await handlers.get(PDF_CHANNELS.insertBlankPage)!(
        { sender: wc },
        {
          path,
          afterPageIndex: 0,
        },
      ),
    ).toMatchObject({ ok: true })
    wc.listeners.get('did-start-loading')?.()
    handlers.get(PDF_CHANNELS.dirtyChanged)!({ sender: wc }, false)
    expect(pdfIsDirty(wc.id)).toBe(true)
    expect(handlers.get(PDF_CHANNELS.sourceModified)!({ sender: wc })).toBe(true)
    vi.mocked(dialog.showMessageBox).mockResolvedValueOnce({ response: 2, checkboxChecked: false })
    expect(await requestPdfClose(wc as never)).toBe(false)
    expect(readdirSync(dir)).toEqual([])
  })

  it('routes explicit Save through the renderer even before any edits exist', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'pdf-blank-save-'))
    createPdfView(undefined, { blank: true, saveDir: dir })
    const wc = lastWebContents
    const saving = flushPdfSave(wc as never)
    expect(wc.send).toHaveBeenCalledWith(PDF_CHANNELS.closeSaveRequest)
    handlers.get(PDF_CHANNELS.closeSaveResult)!({ sender: wc }, false)
    expect(await saving).toBe(false)
    expect(readdirSync(dir)).toEqual([])
  })

  it('does not create a file until the user confirms Save in its originating folder', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'pdf-origin-'))
    createPdfView(undefined, { blank: true, saveDir: dir, untitledName: '未命名 PDF' })
    const wc = lastWebContents
    const virtualPath = consume(wc.id)
    expect(virtualPath).toMatch(/^genoffice-memory-pdf:\/\//)
    expect(readdirSync(dir)).toEqual([])

    const read = handlers.get(PDF_CHANNELS.readFile)!
    const bytes = (await read({ sender: wc }, virtualPath)) as ArrayBuffer
    expect((await PDFDocument.load(new Uint8Array(bytes))).getPageCount()).toBe(1)
    expect(readdirSync(dir)).toEqual([])

    const inserted = (await handlers.get(PDF_CHANNELS.insertBlankPage)!(
      { sender: wc },
      {
        path: virtualPath,
        afterPageIndex: 0,
      },
    )) as { ok: boolean }
    expect(inserted.ok).toBe(true)
    const editedBytes = (await read({ sender: wc }, virtualPath)) as ArrayBuffer
    expect((await PDFDocument.load(new Uint8Array(editedBytes))).getPageCount()).toBe(2)
    expect(readdirSync(dir)).toEqual([])
    expect(handlers.get(PDF_CHANNELS.isUntitled)!({ sender: wc }, virtualPath)).toBe(false)
    const memorySave = (await handlers.get(PDF_CHANNELS.save)!(
      { sender: wc },
      {
        path: virtualPath,
        markups: [],
        drawings: [],
        formValues: [],
        stamps: [],
      },
    )) as { ok: boolean }
    expect(memorySave.ok).toBe(true)
    expect(readdirSync(dir)).toEqual([])

    vi.mocked(dialog.showSaveDialog).mockResolvedValueOnce({ canceled: true, filePath: '' })
    const initialSave = handlers.get(PDF_CHANNELS.requestInitialSave)!
    expect(await initialSave({ sender: wc })).toBe(false)
    expect(vi.mocked(dialog.showSaveDialog).mock.calls.at(-1)?.[0]).toMatchObject({
      defaultPath: join(dir, '未命名 PDF.pdf'),
    })
    expect(readdirSync(dir)).toEqual([])
    expect(consume(wc.id)).toBe(virtualPath)

    const target = join(dir, 'Saved.pdf')
    const materialized = vi.fn()
    setPdfMaterializedHook(materialized)
    vi.mocked(dialog.showSaveDialog).mockResolvedValueOnce({ canceled: false, filePath: target })
    wc.send.mockImplementation((channel: string, requestedPath?: string) => {
      if (channel !== PDF_CHANNELS.saveAsRequest) return
      void (async () => {
        const result = (await handlers.get(PDF_CHANNELS.save)!(
          { sender: wc },
          {
            path: virtualPath,
            targetPath: requestedPath,
            markups: [],
            drawings: [],
            formValues: [],
            stamps: [],
          },
        )) as { ok: boolean }
        handlers.get(PDF_CHANNELS.saveAsResult)!({ sender: wc }, result.ok)
      })()
    })
    expect(await initialSave({ sender: wc })).toBe(true)
    expect(readdirSync(dir)).toEqual(['Saved.pdf'])
    expect(consume(wc.id)).toBe(target)
    expect(materialized).toHaveBeenCalledWith(wc, target)
    const savedBytes = (await read({ sender: wc }, target)) as ArrayBuffer
    expect((await PDFDocument.load(new Uint8Array(savedBytes))).getPageCount()).toBe(2)
  })
})

describe('pdf create-document IPC', () => {
  const request: CreateDocumentRequest = {
    type: 'pdf',
    title: 'Summary',
    content: '<h1>Summary</h1><p>Body</p>',
  }

  it('allows a registered pathless PDF view and forwards a normalized request', async () => {
    const createDocument = vi.fn(async (): Promise<CreateDocumentResult> => ({
      ok: true,
      path: '/tmp/Summary.pdf',
    }))
    configurePdfRuntime({ preloadPath: '', createDocument })
    createPdfView()
    const result = await handlers.get(PDF_CHANNELS.createDocument)?.(
      { sender: { id: lastWebContents.id } },
      request,
    )

    expect(result).toEqual({ ok: true, path: '/tmp/Summary.pdf' })
    expect(createDocument).toHaveBeenCalledWith(request)
  })

  it('rejects unregistered or destroyed senders', async () => {
    const createDocument = vi.fn(async (): Promise<CreateDocumentResult> => ({ ok: true }))
    configurePdfRuntime({ preloadPath: '', createDocument })
    createPdfView(makePdfFile())
    const handler = handlers.get(PDF_CHANNELS.createDocument)

    await expect(handler?.({ sender: { id: 999_999 } }, request)).resolves.toMatchObject({
      ok: false,
      error: expect.stringContaining('not a registered PDF view'),
    })

    const wc = lastWebContents
    wc.listeners.get('destroyed')?.()
    await expect(handler?.({ sender: { id: wc.id } }, request)).resolves.toMatchObject({
      ok: false,
      error: expect.stringContaining('not a registered PDF view'),
    })
    expect(createDocument).not.toHaveBeenCalled()
  })

  it('rejects malformed and oversized requests before calling the host', async () => {
    const createDocument = vi.fn(async (): Promise<CreateDocumentResult> => ({ ok: true }))
    configurePdfRuntime({ preloadPath: '', createDocument })
    createPdfView(makePdfFile())
    const handler = handlers.get(PDF_CHANNELS.createDocument)
    const event = { sender: { id: lastWebContents.id } }

    for (const bad of [
      { ...request, type: 'xlsx' },
      { ...request, title: ' ' },
      { ...request, content: ' ' },
      { ...request, content: 'x'.repeat(2_000_001) },
    ]) {
      await expect(handler?.(event, bad)).resolves.toMatchObject({
        ok: false,
        error: expect.stringContaining('invalid create-document request'),
      })
    }
    expect(createDocument).not.toHaveBeenCalled()
  })
})
