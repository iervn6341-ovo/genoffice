/**
 * Shared launcher for Electron E2E tests.
 *
 * Each test boots the built shell (`apps/shell/out`) against a scratch
 * userData dir (via GENOFFICE_USER_DATA) so runs never touch real settings
 * and never collide with a running install's single-instance lock.
 * Build first: `npm run build:all`.
 */
import { _electron as electron, type ElectronApplication, type Page } from '@playwright/test'
import { mkdtemp, writeFile } from 'node:fs/promises'
import { existsSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'
import { createRequire } from 'node:module'

export const SHELL_DIR = resolve(__dirname, '../apps/shell')
export const ARTIFACTS_DIR = resolve(__dirname, 'artifacts')

const SHELL_MAIN = join(SHELL_DIR, 'out/main/index.js')

/** Largest window the specs open: fits a 13" laptop screen without covering the whole desktop */
export const TEST_WINDOW = { width: 1280, height: 800 }

interface LaunchOptions {
  /** reuse a previous scratch dir to simulate a second launch */
  userDataDir?: string
  /** UI language override (GENOFFICE_LANG); defaults to English for stable assertions */
  lang?: string
  /** pre-seed app-settings.json with onboardingSeen=true to start at the home screen */
  onboardingSeen?: boolean
  /** extra app-settings.json keys (e.g. defaultSaveDir) written before launch */
  settings?: Record<string, unknown>
  /** subdir of e2e/artifacts to store this launch's video in */
  videoDir: string
  /** absolute document path passed as argv, opened in an editor tab on launch */
  openFile?: string
  /** extra environment variables for the launched app */
  env?: Record<string, string>
  /** extra Chromium switches (e.g. a fake camera: --use-fake-device-for-media-stream) */
  chromiumArgs?: string[]
  /**
   * Let the app take the foreground. By default test windows open in the background
   * (GENOFFICE_E2E_BACKGROUND: no Dock icon, never activated, so a run does not steal the
   * keyboard from whoever is using the machine); input still arrives through CDP. Only
   * specs that assert real OS focus (a full-screen show's keyboard, native contenteditable
   * focus hand-offs) need this. GENOFFICE_E2E_FOREGROUND=1 forces it for a whole run.
   */
  foreground?: boolean
}

export interface LaunchedApp {
  app: ElectronApplication
  page: Page
  userDataDir: string
}

export async function launchShell(options: LaunchOptions): Promise<LaunchedApp> {
  if (!existsSync(SHELL_MAIN)) {
    throw new Error(`Missing build output at ${SHELL_MAIN} — run \`npm run build:all\` first`)
  }
  const userDataDir = options.userDataDir ?? (await mkdtemp(join(tmpdir(), 'genoffice-e2e-')))
  if (options.onboardingSeen || options.settings) {
    await writeFile(
      join(userDataDir, 'app-settings.json'),
      JSON.stringify({
        ...(options.onboardingSeen ? { onboardingSeen: true } : {}),
        ...options.settings,
      }),
    )
  }
  const require = createRequire(join(SHELL_DIR, 'package.json'))
  const executablePath = require('electron') as unknown as string
  // ELECTRON_RUN_AS_NODE (set by VS Code/CI hosts) would boot Electron as
  // plain Node with no windows — strip it so the app always starts as an app
  const { ELECTRON_RUN_AS_NODE: _electronRunAsNode, ...hostEnv } = process.env
  // Linux CI runners restrict unprivileged user namespaces (no usable SUID
  // sandbox) and run under xvfb without GPU — without these the window opens
  // but the renderer never loads. The suite drives trusted local builds only.
  // Switches go before the app path so Chromium is guaranteed to consume them
  // and they never leak into the argv the app parses for documents to open.
  const args: string[] = []
  if (process.platform === 'linux') args.push('--no-sandbox', '--disable-gpu')
  if (options.chromiumArgs) args.push(...options.chromiumArgs)
  args.push(SHELL_DIR)
  if (options.openFile) args.push(options.openFile)
  const app = await electron.launch({
    executablePath,
    args,
    env: {
      ...hostEnv,
      GENOFFICE_USER_DATA: userDataDir,
      GENOFFICE_NO_SPARE_VIEW: '1',
      GENOFFICE_LANG: options.lang ?? 'en',
      ...(options.foreground || process.env.GENOFFICE_E2E_FOREGROUND === '1'
        ? {}
        : { GENOFFICE_E2E_BACKGROUND: '1' }),
      ...(options.env ?? {}),
      ...(process.platform === 'linux' ? { ELECTRON_DISABLE_SANDBOX: '1' } : {}),
    },
    // Playwright's Electron screencast wedges the page CDP session on Linux
    // (page.url() stays empty, no lifecycle events, evaluate hangs) — record
    // only where it works
    // GENOFFICE_E2E_NO_VIDEO=1 skips recording (the macOS screencast can stall the launch)
    recordVideo:
      process.platform === 'linux' || process.env.GENOFFICE_E2E_NO_VIDEO
        ? undefined
        : {
            dir: join(ARTIFACTS_DIR, options.videoDir),
            size: { width: 1280, height: 800 },
          },
  })
  const page = await app.firstWindow()
  await waitForDocumentReady(app, page)
  // keep test windows small (the default fills most of a laptop screen) and pinned to the
  // top-left of the primary screen's work area, so every run opens in the same place
  await app.evaluate(({ BrowserWindow, screen }, max) => {
    const win = BrowserWindow.getAllWindows()[0]
    if (!win) return
    const [w, h] = win.getContentSize() as [number, number]
    if (w > max.width || h > max.height) {
      win.setContentSize(Math.min(w, max.width), Math.min(h, max.height))
    }
    const { x, y } = screen.getPrimaryDisplay().workArea
    win.setPosition(x, y)
  }, TEST_WINDOW)
  return { app, page, userDataDir }
}

/**
 * Playwright can attach to the Electron window mid-navigation and miss the
 * load lifecycle events entirely (Linux timing) — waitForLoadState then hangs
 * on a page that is actually loaded. Polling through evaluate uses the live
 * CDP session instead of the missed events.
 */
async function waitForDocumentReady(
  app: ElectronApplication,
  page: Page,
  timeoutMs = 30_000,
): Promise<void> {
  const deadline = Date.now() + timeoutMs
  while (Date.now() < deadline) {
    // the pre-navigation about:blank document also reports readyState complete
    const ready = await page
      .evaluate(() =>
        document.readyState !== 'loading' && window.location.href !== 'about:blank'
          ? document.readyState
          : null,
      )
      .catch(() => null)
    if (ready) return
    await new Promise((r) => setTimeout(r, 100))
  }
  // process list tells renderer-spawn failures apart from slow loads
  const diag = await app
    .evaluate(({ app: electronApp, BrowserWindow }) => ({
      processes: electronApp.getAppMetrics().map((m) => m.type),
      contents: BrowserWindow.getAllWindows().map((w) => w.webContents.getURL()),
    }))
    .catch((e) => String(e))
  throw new Error(`Shell window never loaded (url: ${page.url()}, diag: ${JSON.stringify(diag)})`)
}

/**
 * Close the app and return the recorded video path for the given page.
 *
 * Open editor tabs trigger a native Save/Don't Save/Cancel dialog on close,
 * which would block app.close() forever — stub the dialog to answer
 * "Don't Save" (button index 1) so shutdown stays unattended. If close still
 * hangs, force-kill the process after 20s and wait for close to finish.
 */
export async function closeAndSaveVideo(
  launched: LaunchedApp,
  name: string,
): Promise<string | undefined> {
  const video = launched.page.video()
  await launched.app
    .evaluate(({ dialog }) => {
      dialog.showMessageBox = (async () => ({
        response: 1,
        checkboxChecked: false,
      })) as typeof dialog.showMessageBox
    })
    .catch(() => {})
  const child = launched.app.process()
  const killTimer = setTimeout(() => {
    // SIGTERM can enter Electron's graceful quit path and leave it alive.
    // child.killed only means a signal was sent, not that the process exited.
    if (child.exitCode === null && child.signalCode === null) child.kill('SIGKILL')
  }, 20_000)
  try {
    await launched.app.close()
  } finally {
    clearTimeout(killTimer)
  }
  if (!video) return undefined
  const target = join(ARTIFACTS_DIR, 'videos', `${name}.webm`)
  try {
    await video.saveAs(target)
    return target
  } catch {
    return undefined
  }
}

export function screenshotPath(name: string): string {
  return join(ARTIFACTS_DIR, 'screenshots', `${name}.png`)
}

/**
 * Wait for a page whose URL contains `urlPart` (e.g. an editor WebContentsView).
 * Checks windows that already exist before listening, so it never races the
 * view being created between launch and the first waitForEvent call.
 */
export async function waitForPageWithUrl(
  app: ElectronApplication,
  urlPart: string,
  timeoutMs = 30_000,
): Promise<Page> {
  const deadline = Date.now() + timeoutMs
  for (;;) {
    for (const candidate of app.windows()) {
      if (candidate.url().includes(urlPart)) return candidate
      // page.url() stays empty when attach raced navigation; ask the document
      const href = await candidate.evaluate(() => window.location.href).catch(() => '')
      if (href.includes(urlPart)) return candidate
    }
    const remaining = deadline - Date.now()
    if (remaining <= 0) throw new Error(`No window with URL containing "${urlPart}"`)
    await app.waitForEvent('window', { timeout: Math.min(remaining, 1_000) }).catch(() => {})
  }
}

/**
 * Lay the editor out `cssWidth` CSS px wide while keeping the real window within TEST_WINDOW.
 * Up to TEST_WINDOW.width the window itself is sized; wider layouts (a fully expanded ribbon)
 * zoom the editor view out instead, so the window never outgrows the screen. Playwright's mouse
 * and getBoundingClientRect both work in CSS px, so specs stay coordinate-agnostic.
 * Call after the editor page exists (`urlPart` finds its WebContents, e.g. '://slides/').
 */
export async function setEditorLayoutWidth(
  app: ElectronApplication,
  urlPart: string,
  cssWidth: number,
): Promise<void> {
  const width = Math.min(cssWidth, TEST_WINDOW.width)
  await app.evaluate(
    ({ BrowserWindow, screen, webContents }, { width, height, cssWidth, urlPart }) => {
      const win = BrowserWindow.getAllWindows()[0]!
      win.setContentSize(width, height)
      const { x, y } = screen.getPrimaryDisplay().workArea
      win.setPosition(x, y)
      const wc = webContents.getAllWebContents().find((w) => w.getURL().includes(urlPart))
      wc?.setZoomFactor(width / cssWidth)
    },
    { width, height: TEST_WINDOW.height, cssWidth, urlPart },
  )
  // let the view resize and the ribbon fit settle
  await new Promise((r) => setTimeout(r, 400))
}
