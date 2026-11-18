import { BrowserWindow, nativeTheme, screen } from 'electron'
import type { WebContents } from 'electron'

/** A read-only thumbnail. It never takes capture, keyboard focus or the drop. */
export class TabDragPreview {
  private win: BrowserWindow | null = null
  private loaded = false
  private visible = false
  private title = ''
  private label = ''
  private thumbnail = ''
  private painted = ''
  private owner: WebContents | null = null
  tabId: string | null = null

  belongsTo(owner: WebContents): boolean {
    return this.owner === owner
  }

  show(
    owner: WebContents,
    editor: WebContents,
    id: string,
    title: string,
    label: string,
    x: number,
    y: number,
  ): void {
    if (this.owner !== owner || this.tabId !== id) {
      this.end()
      this.owner = owner
      this.tabId = id
      owner.once('destroyed', this.onOwnerDestroyed)
      const win = new BrowserWindow({
        width: 284,
        height: 218,
        show: false,
        frame: false,
        transparent: true,
        focusable: false,
        skipTaskbar: true,
        alwaysOnTop: true,
        resizable: false,
        movable: false,
        minimizable: false,
        maximizable: false,
        title: 'GenOffice Tab Preview',
        webPreferences: { sandbox: true, contextIsolation: true, nodeIntegration: false },
      })
      this.win = win
      win.setIgnoreMouseEvents(true)
      const dark = nativeTheme.shouldUseDarkColors
      const html = `<!doctype html><html><head><meta charset="utf-8"><meta http-equiv="Content-Security-Policy" content="default-src 'none'; img-src data:; style-src 'unsafe-inline'"><style>
        :root { --surface: ${dark ? '#292929' : '#ffffff'}; --text: ${dark ? '#eeeeee' : '#242424'}; --border: ${dark ? '#555555' : '#d9d9d9'}; --accent: ${dark ? '#81b4ff' : '#0f7fff'}; }
        * { box-sizing: border-box } body { margin: 0; padding: 4px; font: 12px system-ui; color: var(--text); user-select: none }
        .card { height: 210px; padding: 10px; border: 2px solid var(--accent); border-radius: 12px; background: var(--surface); overflow: hidden }
        #title { font-weight: 600; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; margin-bottom: 8px }
        img { display: block; width: 100%; height: 132px; object-fit: contain; border: 1px solid var(--border); border-radius: 4px }
        #label { margin-top: 8px; font-size: 11px; line-height: 16px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis }
      </style></head><body><div class="card"><div id="title"></div><img alt=""><div id="label"></div></div></body></html>`
      void win
        .loadURL(`data:text/html;charset=utf-8,${encodeURIComponent(html)}`)
        .then(() => {
          if (this.win !== win || win.isDestroyed()) return
          this.loaded = true
          this.paint()
          if (this.visible) win.showInactive()
        })
        .catch(() => {
          if (this.win === win) this.end()
        })
      void editor
        .capturePage()
        .then((image) => {
          if (this.win !== win || win.isDestroyed() || image.isEmpty()) return
          this.thumbnail = image.resize({ width: 264 }).toDataURL()
          this.paint()
        })
        .catch(() => {})
    }
    this.title = title
    this.label = label
    this.visible = true
    const area = screen.getDisplayNearestPoint({ x: Math.round(x), y: Math.round(y) }).workArea
    this.win?.setPosition(
      Math.round(Math.max(area.x, Math.min(x + 18, area.x + area.width - 284))),
      Math.round(Math.max(area.y, Math.min(y + 24, area.y + area.height - 218))),
    )
    this.paint()
    if (this.loaded && !this.win?.isVisible()) this.win?.showInactive()
  }

  hide(): void {
    this.visible = false
    if (this.win && !this.win.isDestroyed()) this.win.hide()
  }

  private paint(): void {
    if (!this.loaded || !this.win || this.win.isDestroyed()) return
    const payload = JSON.stringify({ title: this.title, label: this.label, image: this.thumbnail })
    if (payload === this.painted) return
    this.painted = payload
    void this.win.webContents
      .executeJavaScript(
        `(() => { const p=${payload}; document.getElementById('title').textContent=p.title; document.getElementById('label').textContent=p.label; if(p.image) document.querySelector('img').src=p.image; })()`,
      )
      .catch(() => {})
  }

  private onOwnerDestroyed = (): void => this.end()

  end(): void {
    this.owner?.removeListener('destroyed', this.onOwnerDestroyed)
    const win = this.win
    this.win = null
    this.owner = null
    this.tabId = null
    this.loaded = false
    this.visible = false
    this.thumbnail = ''
    this.painted = ''
    if (win && !win.isDestroyed()) win.destroy()
  }
}
