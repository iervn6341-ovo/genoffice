import { describe, expect, it } from 'vitest'
import {
  applyNotesHandoutsHeaderFooter,
  createBlankPptx,
  listMasterParts,
  openPptx,
  parseMasterPart,
  savePptx,
} from '@genoffice/pptx-engine'
import { runTxn } from '../src/ops/executor'
import '../src/ops/index'

describe('native master formatting ops', () => {
  it('keeps a notes-master field ID and type while changing its font, alignment, geometry and background', async () => {
    const opened = await openPptx(await createBlankPptx())
    applyNotesHandoutsHeaderFooter(opened, {
      headerEnabled: false,
      header: '',
      footerEnabled: false,
      footer: '',
      dateEnabled: true,
      date: '10/13/2026',
      dateAuto: true,
      dateFormat: 'datetime8',
      pageNum: true,
    })
    const path = listMasterParts(opened.archive).find((p) => p.kind === 'notesMaster')!.partPath
    const before = opened.archive.readText(path)!
    const fieldId = /<a:fld id="([^"]+)" type="datetime8"/.exec(before)![1]
    const master = parseMasterPart(opened.archive, path)!
    const date = master.elements.find((e) => e.placeholder === 'dt')!
    const result = runTxn(opened, {
      ops: [
        {
          op: 'setFont',
          target: { part: path, el: date.id },
          font: { fontFamily: 'Arial', fontSizePt: 18, bold: true, color: '#224466' },
        },
        {
          op: 'setParagraphFormat',
          target: { part: path, el: date.id },
          format: { align: 'right', lineSpacingPct: 120 },
        },
        {
          op: 'setTransform',
          target: { part: path, el: date.id },
          box: { x: 700000, y: 8000000, cx: 1800000, cy: 400000 },
          rotDeg: 0,
        },
        { op: 'setBackground', target: { part: path }, kind: 'solid', color: '#EEEEDD' },
      ],
      parts: new Map([[path, master]]),
    })
    expect(result.applied).toBe(true)
    const after = opened.archive.readText(path)!
    expect(after).toContain(`<a:fld id="${fieldId}" type="datetime8"`)
    expect(after).toContain('sz="1800"')
    expect(after).toContain('algn="r"')
    expect(after).toContain('<a:off x="700000" y="8000000"/>')
    expect(after).toContain('EEEEDD')
    const reopened = await openPptx(await savePptx(opened))
    expect(reopened.archive.readText(path)).toContain(`<a:fld id="${fieldId}" type="datetime8"`)
  })
})
