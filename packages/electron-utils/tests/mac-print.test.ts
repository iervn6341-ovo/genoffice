import { mkdtempSync, mkdirSync, utimesSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'

import { outcomeForExit, planHelper } from '../src/mac-print'

describe('outcomeForExit', () => {
  it('maps the helper exit codes', () => {
    expect(outcomeForExit(0, '')).toEqual({ status: 'printed' })
    expect(outcomeForExit(1, '')).toEqual({ status: 'canceled' })
    expect(outcomeForExit(2, '')).toMatchObject({
      status: 'error',
      message: expect.stringContaining('PDF'),
    })
    expect(outcomeForExit(3, '')).toMatchObject({ status: 'error' })
    expect(outcomeForExit(null, 'boom\n')).toEqual({ status: 'error', message: 'boom' })
  })
})

describe('planHelper', () => {
  const dir = (): string => mkdtempSync(join(tmpdir(), 'mac-print-'))

  it('prefers the binary shipped in the app resources', () => {
    const resources = dir()
    mkdirSync(join(resources, 'print'))
    writeFileSync(join(resources, 'print', 'print-pdf'), '')
    expect(planHelper({ resourcesPath: resources, devDir: dir() })).toEqual({
      binary: join(resources, 'print', 'print-pdf'),
    })
  })

  it('compiles from source when there is no binary, or the source is newer', () => {
    const dev = dir()
    writeFileSync(join(dev, 'print-pdf.swift'), '')
    expect(planHelper({ devDir: dev })).toEqual({
      binary: join(dev, 'print-pdf'),
      compileFrom: join(dev, 'print-pdf.swift'),
    })
    writeFileSync(join(dev, 'print-pdf'), '')
    const past = new Date(Date.now() - 60_000)
    utimesSync(join(dev, 'print-pdf.swift'), past, past)
    expect(planHelper({ devDir: dev })).toEqual({ binary: join(dev, 'print-pdf') })
    const future = new Date(Date.now() + 60_000)
    utimesSync(join(dev, 'print-pdf.swift'), future, future)
    expect(planHelper({ devDir: dev })).toMatchObject({ compileFrom: join(dev, 'print-pdf.swift') })
  })

  it('returns null when nothing is available', () => {
    expect(planHelper({})).toBeNull()
    expect(planHelper({ devDir: dir() })).toBeNull()
  })
})
