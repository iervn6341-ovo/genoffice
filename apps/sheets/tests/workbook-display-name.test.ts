import { describe, expect, it } from 'vitest'

import { workbookDisplayName } from '../src/main/workbook-display-name'

describe('workbookDisplayName', () => {
  it('uses the user workbook instead of the UUID snapshot name', () => {
    expect(workbookDisplayName('/Users/me/Documents/Untitled Spreadsheet.xlsx')).toBe(
      'Untitled Spreadsheet.xlsx',
    )
  })

  it('uses the original name for recovery and CSV imports', () => {
    expect(
      workbookDisplayName('/tmp/recovery/uuid.xlsx', {
        restoreTarget: '/Users/me/Documents/Budget 2026.xlsx',
      }),
    ).toBe('Budget 2026.xlsx')
    expect(
      workbookDisplayName('/tmp/genoffice-imports/uuid/Customers.xlsx', {
        csvSourcePath: '/Users/me/Documents/Customers.csv',
      }),
    ).toBe('Customers.csv')
  })

  it('uses the suggested XLSX name for a converted legacy workbook', () => {
    expect(
      workbookDisplayName('/tmp/genoffice-imports/uuid/Accounts.xlsx', {
        suggestSaveAs: '/Users/me/Documents/Accounts.xlsx',
      }),
    ).toBe('Accounts.xlsx')
  })
})
