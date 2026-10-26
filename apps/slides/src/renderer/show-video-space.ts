/**
 * Space-bar ownership during the show. Space turns the page, except while the pointer rests on a
 * video (or the video's control bar has keyboard focus): then it plays / pauses that video, as
 * in PowerPoint. Each video player registers a claimer; the capture-phase show keys ask them
 * before treating Space as "next".
 */
const claimers = new Set<(e: KeyboardEvent) => boolean>()

/** Register a handler that may claim Space (returns true when it consumed the key). Returns the unregister function. */
export function registerShowVideoSpaceClaimer(fn: (e: KeyboardEvent) => boolean): () => void {
  claimers.add(fn)
  return () => {
    claimers.delete(fn)
  }
}

/** True when a video took the Space key; the caller must then not turn the page. */
export function claimShowSpace(e: KeyboardEvent): boolean {
  if (e.key !== ' ' || e.metaKey || e.ctrlKey || e.altKey) return false
  for (const c of claimers) if (c(e)) return true
  return false
}
