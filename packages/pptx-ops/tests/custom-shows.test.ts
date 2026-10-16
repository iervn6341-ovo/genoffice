import { describe, expect, it } from 'vitest'
import { createBlankPptx, getCustomShows, openPptx, savePptx } from '@genoffice/pptx-engine'
import { runTxn } from '../src/ops/executor'
import '../src/ops/index'

describe('setCustomShows operation', () => {
  it('validates in dry run, writes native presentation metadata, and survives save/reopen', async () => {
    const opened = await openPptx(await createBlankPptx())
    const shows = [{ id: '0', name: 'One slide', slideIndices: [0] }]
    const before = opened.archive.readText('ppt/presentation.xml')
    const dryRun = runTxn(opened, { ops: [{ op: 'setCustomShows', shows }], dryRun: true })
    expect(dryRun.failures).toBeUndefined()
    expect(opened.archive.readText('ppt/presentation.xml')).toBe(before)
    expect(runTxn(opened, { ops: [{ op: 'setCustomShows', shows }] }).applied).toBe(true)
    expect(getCustomShows(opened)).toEqual(shows)
    expect(getCustomShows(await openPptx(await savePptx(opened)))).toEqual(shows)
  })

  it('rejects malformed payloads atomically, preserving the existing native list', async () => {
    const opened = await openPptx(await createBlankPptx())
    const valid = [{ id: '9', name: 'Keep', slideIndices: [0] }]
    expect(runTxn(opened, { ops: [{ op: 'setCustomShows', shows: valid }] }).applied).toBe(true)
    const before = opened.archive.readText('ppt/presentation.xml')
    expect(
      runTxn(opened, {
        ops: [{ op: 'setCustomShows', shows: [{ id: 'bad', name: 'Wrong', slideIndices: [0] }] }],
        dryRun: true,
      }).failures?.[0]?.error,
    ).toContain('invalid show id')
    const result = runTxn(opened, {
      ops: [{ op: 'setCustomShows', shows: [{ id: 'bad', name: 'Wrong', slideIndices: [0] }] }],
    })
    expect(result.applied).toBe(false)
    expect(result.failures?.[0]?.error).toContain('invalid show id')
    expect(opened.archive.readText('ppt/presentation.xml')).toBe(before)
  })
})
