import { describe, expect, it } from 'vitest'
import { createBlankPptx, getShowSetup, openPptx } from '@genoffice/pptx-engine'
import { runTxn } from '../src/index'

describe('setShowSetup operation', () => {
  it('validates and atomically applies native show properties', async () => {
    const opened = await openPptx(await createBlankPptx())
    const settings = {
      mode: 'browse',
      loop: true,
      showNarration: false,
      showAnimation: false,
      useTimings: false,
      slides: { kind: 'range', start: 1, end: 1 },
    }
    const dryRun = runTxn(opened, { ops: [{ op: 'setShowSetup', settings }], dryRun: true })
    expect(dryRun.failures).toBeUndefined()
    expect(getShowSetup(opened).mode).toBe('speaker')
    expect(runTxn(opened, { ops: [{ op: 'setShowSetup', settings }] }).applied).toBe(true)
    expect(getShowSetup(opened)).toMatchObject(settings)
    const before = opened.archive.readText('ppt/presProps.xml')
    expect(
      runTxn(opened, {
        ops: [
          {
            op: 'setShowSetup',
            settings: { ...settings, slides: { kind: 'range', start: 2, end: 3 } },
          },
        ],
      }).applied,
    ).toBe(false)
    expect(opened.archive.readText('ppt/presProps.xml')).toBe(before)
  })
})
