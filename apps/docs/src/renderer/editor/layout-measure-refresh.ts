import type { Editor } from '@tiptap/core'

/** Refresh a display-only measurement after document settings change. */
export function refreshLayoutMeasurement(
  editor: Editor,
  storageName: 'cjkPunctShrink' | 'tabStops',
): void {
  const storage = editor.storage[storageName]
  storage.measureVersion = (storage.measureVersion ?? 0) + 1
  editor.view.dispatch(
    editor.state.tr.setMeta('layoutMeasurementRefresh', storageName).setMeta('addToHistory', false),
  )
}
