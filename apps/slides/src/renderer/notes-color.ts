/** Black/white notes use the UI's readable foreground; authored colors stay in the document. */
export function isNotesDefaultColor(color: string | undefined): boolean {
  if (!color) return true
  const value = color.toLowerCase().replace(/\s/g, '')
  return /^(#000(?:000)?|#fff(?:fff)?|black|white|rgb\(0,0,0\)|rgb\(255,255,255\)|rgba\(0,0,0,1\)|rgba\(255,255,255,1\))$/.test(
    value,
  )
}

/** Paint-only markers: inline color remains untouched for edit extraction and PPTX save. */
export function syncNotesDisplayColors(root: HTMLElement): void {
  const visit = (parent: HTMLElement, inherited: string | undefined): void => {
    for (const el of parent.children) {
      if (!(el instanceof HTMLElement)) continue
      const color = el.style.color || inherited
      el.toggleAttribute('data-notes-default-color', isNotesDefaultColor(color))
      visit(el, color)
    }
  }
  visit(root, root.style.color || undefined)
}
