import { useEffect, useState } from 'react'
import type { Lang } from '@genoffice/i18n'
import { PRINTER_SETTINGS_LABELS } from './strings-printer-settings'

export type PrinterDuplex = 'off' | 'long' | 'short'

export interface PrinterSettings {
  /** System printer name; '' = no printer known (the system dialog is used instead) */
  printer: string
  /** Raw field text so the box can be cleared while typing */
  copies: string
  duplex: PrinterDuplex
  collate: boolean
}

export interface PrinterListItem {
  name: string
  displayName: string
  isDefault: boolean
}

export const DEFAULT_PRINTER_SETTINGS: PrinterSettings = {
  printer: '',
  copies: '1',
  duplex: 'off',
  collate: true,
}

/** Normalized copies for the main process (1…999) */
export function printerCopies(settings: PrinterSettings): number {
  return Math.min(Math.max(Math.floor(Number(settings.copies)) || 1, 1), 999)
}

/** The IPC payload for a print job; empty when no printer is selected (system dialog) */
export function printerJob(settings: PrinterSettings): {
  printer?: string
  copies?: number
  duplex?: PrinterDuplex
  collate?: boolean
} {
  if (!settings.printer) return {}
  return {
    printer: settings.printer,
    copies: printerCopies(settings),
    duplex: settings.duplex,
    collate: settings.collate,
  }
}

/**
 * Printer / copies / two-sided / collated, the printer block of Office's print
 * pane. Loads the printer list once and preselects the system default.
 */
export function PrinterSettingsFields({
  lang,
  value,
  onChange,
  listPrinters,
}: {
  lang: Lang
  value: PrinterSettings
  onChange: (next: PrinterSettings) => void
  listPrinters: () => Promise<PrinterListItem[]>
}) {
  const labels = PRINTER_SETTINGS_LABELS[lang]
  const [printers, setPrinters] = useState<PrinterListItem[]>([])

  useEffect(() => {
    let alive = true
    void listPrinters().then((list) => {
      if (!alive) return
      setPrinters(list)
      const pick = list.find((p) => p.isDefault) ?? list[0]
      if (pick) onChange({ ...value, printer: pick.name })
    })
    return () => {
      alive = false
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [])

  const direct = value.printer !== ''
  return (
    <div className="printer-settings">
      <select
        className="ps-select"
        aria-label={labels.printer}
        value={value.printer}
        onChange={(e) => onChange({ ...value, printer: e.target.value })}
      >
        {printers.length === 0 && <option value="">{labels.systemDefault}</option>}
        {printers.map((p) => (
          <option key={p.name} value={p.name}>
            {p.displayName}
          </option>
        ))}
      </select>
      <label className="ps-field">
        {labels.copies}
        <input
          type="number"
          min={1}
          max={999}
          className="ps-num"
          value={value.copies}
          disabled={!direct}
          onChange={(e) => onChange({ ...value, copies: e.target.value })}
        />
      </label>
      <label className="ps-field">
        {labels.twoSided}
        <select
          className="ps-select"
          value={value.duplex}
          disabled={!direct}
          onChange={(e) => onChange({ ...value, duplex: e.target.value as PrinterDuplex })}
        >
          <option value="off">{labels.off}</option>
          <option value="long">{labels.longEdge}</option>
          <option value="short">{labels.shortEdge}</option>
        </select>
      </label>
      <label className="ps-check">
        <input
          type="checkbox"
          checked={value.collate}
          disabled={!direct}
          onChange={(e) => onChange({ ...value, collate: e.target.checked })}
        />
        {labels.collated}
      </label>
    </div>
  )
}
