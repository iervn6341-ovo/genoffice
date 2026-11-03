/**
 * System clipboard renditions of a copy.
 *
 * Electron's `clipboard` API starts every write by clearing the pasteboard, so one copy can
 * not carry a custom type (PowerPoint's `Art::GVML ClipFormat`, our own marker) together with
 * text / HTML / an image. On macOS the whole set is therefore written in one pasteboard
 * transaction through the system's JavaScript-for-Automation bridge (`osascript -l
 * JavaScript` → NSPasteboard): nothing to build or bundle, and it is the same pasteboard the
 * other apps read. Elsewhere (or when the bridge fails) the copy degrades to the marker only,
 * which keeps in-app paste working.
 */
import { clipboard } from 'electron'
import { execFile } from 'node:child_process'
import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { GVML_FORMAT, TABLE_CLIP_FORMAT } from '@genoffice/pptx-engine'

export const MARKER_ELEMENTS = 'io.genoffice.slides.elements'
export const MARKER_SLIDE = 'io.genoffice.slides.slide'

export interface SystemClipboardPayload {
  /** Our own marker type: while present, the last copy came from this app (internal paste wins) */
  marker: string
  gvml?: Uint8Array
  table?: Uint8Array
  png?: Uint8Array
  text?: string
  html?: string
  /** Further pasteboard types by name (e.g. PowerPoint's slide package) */
  extra?: Record<string, Uint8Array>
}

const PB_TYPES = {
  text: 'public.utf8-plain-text',
  html: 'public.html',
  png: 'public.png',
} as const

const WRITE_SCRIPT = `
ObjC.import('AppKit');
function run(argv) {
  var spec = JSON.parse(argv[0]);
  var pb = $.NSPasteboard.generalPasteboard;
  pb.clearContents;
  var n = 0;
  for (var t in spec) {
    var d = $.NSData.dataWithContentsOfFile($(spec[t]));
    if (d.isNil()) continue;
    if (pb.setDataForType(d, $(t))) n++;
  }
  return String(n);
}`

let chain: Promise<unknown> = Promise.resolve()

function runOsascript(spec: Record<string, string>): Promise<number> {
  return new Promise((resolve, reject) => {
    execFile(
      '/usr/bin/osascript',
      ['-l', 'JavaScript', '-e', WRITE_SCRIPT, JSON.stringify(spec)],
      { timeout: 8000 },
      (err, stdout) => (err ? reject(err) : resolve(parseInt(String(stdout).trim(), 10) || 0)),
    )
  })
}

/** Legacy path: only the marker survives (writeBuffer clears everything else) */
function writeMarkerOnly(marker: string): void {
  clipboard.writeBuffer(marker, Buffer.from('1'))
}

/**
 * Put every rendition of a copy on the system clipboard at once. Resolves true when the
 * full set landed; on failure the marker alone is written so in-app paste still works.
 */
export function writeSystemClipboard(payload: SystemClipboardPayload): Promise<boolean> {
  const run = async (): Promise<boolean> => {
    if (process.platform !== 'darwin') {
      writeMarkerOnly(payload.marker)
      return false
    }
    let dir: string | null = null
    try {
      dir = await mkdtemp(join(tmpdir(), 'genoffice-cb-'))
      const spec: Record<string, string> = {}
      const put = async (type: string, data: Uint8Array | string | undefined) => {
        if (data === undefined || (typeof data !== 'string' && data.byteLength === 0)) return
        const file = join(dir!, `f${Object.keys(spec).length}`)
        await writeFile(file, data)
        spec[type] = file
      }
      await put(payload.marker, '1')
      await put(GVML_FORMAT, payload.gvml)
      await put(TABLE_CLIP_FORMAT, payload.table)
      await put(PB_TYPES.png, payload.png)
      await put(PB_TYPES.html, payload.html)
      await put(PB_TYPES.text, payload.text)
      for (const [type, data] of Object.entries(payload.extra ?? {})) await put(type, data)
      const written = await runOsascript(spec)
      if (written < Object.keys(spec).length) {
        // a partial write still carries the marker first; make sure it is there
        if (!clipboard.readBuffer(payload.marker).length) writeMarkerOnly(payload.marker)
        return false
      }
      return true
    } catch {
      writeMarkerOnly(payload.marker)
      return false
    } finally {
      if (dir) void rm(dir, { recursive: true, force: true })
    }
  }
  const next = chain.then(run, run)
  chain = next.catch(() => undefined)
  return next
}

/** Raw bytes of a custom pasteboard type (empty buffer when absent) */
export function readCustomFormat(format: string): Buffer {
  try {
    return clipboard.readBuffer(format)
  } catch {
    return Buffer.alloc(0)
  }
}
