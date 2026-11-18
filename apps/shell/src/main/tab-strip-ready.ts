/** Renderer acknowledgement after the active tab's title bar has been laid out. */
export class TabStripReady {
  private activeId: string | null = null
  private closed = false
  private pending = new Set<{ id: string; finish: (ready: boolean) => void }>()

  acknowledge(id: string): void {
    if (this.closed) return
    this.activeId = id
    for (const p of this.pending) if (p.id === id) p.finish(true)
  }

  wait(id: string, timeoutMs = 8000): Promise<boolean> {
    if (this.closed) return Promise.resolve(false)
    if (this.activeId === id) return Promise.resolve(true)
    return new Promise((resolve) => {
      const entry = {
        id,
        finish: (ready: boolean) => {
          clearTimeout(timer)
          this.pending.delete(entry)
          resolve(ready)
        },
      }
      const timer = setTimeout(() => entry.finish(false), timeoutMs)
      this.pending.add(entry)
    })
  }

  dispose(): void {
    this.closed = true
    for (const p of this.pending) p.finish(false)
  }
}
