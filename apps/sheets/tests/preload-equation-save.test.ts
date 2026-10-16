import { describe, expect, it, vi } from 'vitest'
import { latexToOmml } from '@genoffice/office-math'

vi.mock('electron', () => ({
  contextBridge: { exposeInMainWorld: vi.fn() },
  ipcRenderer: { invoke: vi.fn(), on: vi.fn(), removeListener: vi.fn() },
  webUtils: { getPathForFile: vi.fn() },
}))
vi.mock('@genoffice/electron-utils/drop-open', () => ({ installDropOpenBridge: vi.fn() }))

import { parseSaveRequest } from '../src/preload/index'
import { workbookSaveRequestSchema } from '../src/shared/desktop-api'

const equation = { latex: 'x+1', omml: latexToOmml('x+1') }

function equationSaveRequest() {
  return workbookSaveRequestSchema.parse({
    sessionId: '5d4f6f7a-1c2b-4e3d-9a8f-0b1c2d3e4f5a',
    mode: 'save',
    edits: [],
    structuralOps: [],
    chartEdits: [],
    visualEdits: [{ drawingPath: 'xl/drawings/drawing1.xml', drawingIndex: 0, equation }],
    visualAdditions: [],
    tableAdditions: [],
    pivotAdditions: [],
    sheetOps: [],
    sheetOrder: [],
    filterStates: [],
    hyperlinkEdits: [],
    cfStates: [],
    dvStates: [],
    pageSetupStates: [],
    noteStates: [],
    formulaValues: [],
    pivotCacheRefreshPaths: [],
    pivotRefreshUpdates: [],
    sheetProtections: [],
    sparklineAdditions: [],
    definedNamesState: null,
  })
}

describe('preload save validation for native equation edits', () => {
  it('passes an equation-only drawing edit without an anchor to the main process', () => {
    const request = equationSaveRequest()
    expect(parseSaveRequest(request).visualEdits).toEqual([
      { drawingPath: 'xl/drawings/drawing1.xml', drawingIndex: 0, equation },
    ])
  })

  it('rejects malformed equation data at the preload boundary', () => {
    const request = equationSaveRequest()
    const malformed = {
      ...request,
      visualEdits: [{ ...request.visualEdits[0], equation: { latex: '', omml: equation.omml } }],
    }
    expect(() => parseSaveRequest(malformed as never)).toThrow(/Invalid workbook visual edit/)
  })
})
