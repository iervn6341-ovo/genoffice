/**
 * Slide Show → Monitor (PowerPoint's Set Up group): which screen shows the slides when presenter
 * view opens. Stored by display label (ids change across reconnects); null = Automatic.
 * A per-machine preference, like "Use Presenter View".
 */
const KEY = 'genoffice.slides.showMonitor'

export function readShowMonitor(): string | null {
  try {
    return localStorage.getItem(KEY) || null
  } catch {
    return null
  }
}

export function writeShowMonitor(label: string | null): void {
  try {
    if (label) localStorage.setItem(KEY, label)
    else localStorage.removeItem(KEY)
  } catch {
    // storage unavailable: the choice lasts for this session
  }
}
