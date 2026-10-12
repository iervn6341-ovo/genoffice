import { mkdir, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { expect, test, type Page } from '@playwright/test'
import {
  closeAndSaveVideo,
  launchShell,
  setEditorLayoutWidth,
  waitForPageWithUrl,
  type LaunchedApp,
} from './helpers'

/**
 * A sweep of the Sheets ribbon: every tab, every dropdown, every option once, with a 4 × 3 range
 * selected (values, formulas, styles). For each choice it records
 *   - uncaught page errors and console errors raised by it,
 *   - whether it opened a dialog (closed again with Esc),
 *   - whether one ⌘Z brings the model back (cell values, formulas, styles, merges, row heights and
 *     column widths over A1:H14) — an edit that cannot be undone is a finding unless it changes nothing.
 * The findings are written to AUDIT_OUT (default: a folder in the temp directory) and the spec fails on
 * a page error or on a choice that changes the model and cannot be undone.
 *
 *   AUDIT_TABS="Home,Data"   only these tabs (default: all seven)
 *   AUDIT_SKIP="Pivot|Open"  option / button labels (regex, case-insensitive) that are not pressed
 *   AUDIT_OUT=<dir>          where all.summary.txt and all.findings.json go
 */
const ALL_TABS = ['Home', 'Insert', 'Page Layout', 'Formulas', 'Data', 'Review', 'View']
const TABS = (process.env.AUDIT_TABS?.split(',').map((tab) => tab.trim()) ?? ALL_TABS).filter(
  (tab) => ALL_TABS.includes(tab),
)
const SKIP = new RegExp(
  process.env.AUDIT_SKIP ??
    '^(Save|Save As|Open|Print|Share|Export|Close|New|Quick Print|Pivot|PivotTable|Screenshot|Pictures?|Online|From (Text|Web)|Get Data|Refresh All|Protect Workbook|Unprotect|Password|Sign|Macros?|Record|Sync|Help|Copilot|AI)',
  'i',
)
/// only the controls whose label matches (a regex), for following up one finding
const CONTROLS = process.env.AUDIT_CONTROLS ? new RegExp(process.env.AUDIT_CONTROLS, 'i') : null
const OUT = process.env.AUDIT_OUT ?? join(tmpdir(), 'genoffice-sheets-ribbon-audit')
/// of a gallery with many cards, this many are pressed (the first ones and the last)
const GALLERY_SAMPLE = 3

interface Finding {
  readonly tab: string
  readonly control: string
  readonly option: string | null
  readonly kind:
    | 'page-error'
    | 'console-error'
    | 'not-undoable'
    | 'undo-incomplete'
    | 'dialog'
    | 'skipped'
    | 'failed'
  readonly detail: string
}

async function openSheets(): Promise<{ launched: LaunchedApp; p: Page }> {
  const launched = await launchShell({
    onboardingSeen: true,
    lang: 'en',
    videoDir: 'sheets-ribbon-audit',
    env: { GENOFFICE_DEBUG_HOOKS: '1' },
  })
  await launched.page.locator('.quick-card').nth(1).click()
  const p = await waitForPageWithUrl(launched.app, '://sheets/', 30_000)
  await p.waitForSelector('[data-ribbon-body]', { timeout: 20_000 })
  await p.waitForTimeout(1500)
  await setEditorLayoutWidth(launched.app, '://sheets/', 2400)
  return { launched, p }
}
const kill = (l: LaunchedApp) => closeAndSaveVideo(l, 'cleanup')
const api = '(window).__genofficeDebug.univerAPI.getActiveWorkbook().getActiveSheet()'
const tab = (p: Page, name: string) =>
  p
    .locator('nav.ribbon-tabs > button:not(.qa-btn):not(.ribbon-tab-file)', {
      hasText: new RegExp(`^${name}$`),
    })
    .first()
    .click()

/// the model the audit compares before / after: A1:H14 values, formulas and styles, merges, sizes
async function snapshot(p: Page): Promise<string> {
  return p.evaluate(`(() => {
    const sheet = ${api}
    const range = sheet.getRange('A1:H14')
    const formulas = range.getFormulas()
    // a formula cell's result is recalculated a moment after the edit: its formula is what is compared
    const values = range.getRawValues().map((line, r) => line.map((v, c) => (formulas[r][c] ? null : v)))
    const styles = []
    for (let r = 0; r < 14; r += 1) for (let c = 0; c < 8; c += 1) {
      styles.push(JSON.stringify(sheet.getRange(r, c, 1, 1).getCellStyleData() ?? null))
    }
    return JSON.stringify({
      values,
      formulas,
      styles,
      merges: sheet.getMergedRanges().map((m) => [m.getRow(), m.getColumn(), m.getHeight(), m.getWidth()]),
      rows: Array.from({ length: 14 }, (_, r) => sheet.getRowHeight(r)),
      columns: Array.from({ length: 8 }, (_, c) => sheet.getColumnWidth(c)),
      sheets: ${api.replace('.getActiveSheet()', '')}.getSheets().map((s) => s.getSheetName()),
    })
  })()`) as Promise<string>
}

const SEED = [
  ['Item', 'Qty', 'Price'],
  ['Apple', 3, 1.5],
  ['Pear', 5, 2.25],
  ['Plum', 8, 0.75],
]

async function seed(p: Page): Promise<void> {
  await p.evaluate(`${api}.getRange('A1:C4').setValues(${JSON.stringify(SEED)})`)
  await p.evaluate(`${api}.getRange('D2').setFormula('=B2*C2')`)
  await p.evaluate(`${api}.getRange('A1:C1').setFontWeight('bold')`)
  await p.evaluate(`${api}.getRange('C2:C4').setNumberFormat('0.00')`)
  await p.evaluate(`${api}.getRange('A1:C4').activate()`)
}

/// closes what a choice opened: Esc (what a user presses), then the dialog's own Cancel / Close button.
/// `stuck` names a dialog that Esc did not close.
async function closeDialogs(p: Page): Promise<{ opened: boolean; stuck: string | null }> {
  let opened = false
  let stuck: string | null = null
  for (let i = 0; i < 4; i += 1) {
    const dialog = p.locator('[role="dialog"], .dialog-backdrop').first()
    if (!(await dialog.isVisible().catch(() => false))) break
    opened = true
    await p.keyboard.press('Escape')
    await p.waitForTimeout(300)
    if (!(await dialog.isVisible().catch(() => false))) break
    stuck ??=
      (await p
        .locator('[role="dialog"] header, .dialog-backdrop header')
        .first()
        .innerText()
        .catch(() => 'a dialog')) || 'a dialog'
    const cancel = p
      .locator('[role="dialog"] button, .dialog-backdrop button')
      .filter({ hasText: /^(Cancel|Close|Done)$/i })
      .first()
    if (await cancel.isVisible().catch(() => false)) await cancel.click().catch(() => undefined)
    await p.waitForTimeout(300)
  }
  return { opened, stuck }
}

test(`the ribbon sweep: ${TABS.join(', ')}`, async () => {
  test.setTimeout(55 * 60 * 1000)
  let { launched, p } = await openSheets()
  const findings: Finding[] = []
  let pressed = 0
  let logged: string[] = []
  let currentTab = TABS[0] ?? 'Home'
  const attach = (): void => {
    p.on('pageerror', (error) => logged.push(`page-error: ${error.message}`))
    p.on('console', (message) => {
      if (message.type() === 'error') logged.push(`console-error: ${message.text().slice(0, 300)}`)
    })
  }
  attach()
  let baseline = ''
  /// a fresh app on the same tab: after a choice that cannot be undone, or one that wedges the window
  const restart = async (): Promise<void> => {
    await kill(launched).catch(() => undefined)
    ;({ launched, p } = await openSheets())
    attach()
    await seed(p)
    baseline = await snapshot(p)
    await tab(p, currentTab)
    await p.waitForTimeout(400)
  }
  /// a window that no longer answers is restarted
  const healthy = async (): Promise<boolean> =>
    Promise.race([
      p.evaluate('1').then(() => true),
      new Promise<boolean>((resolve) => setTimeout(() => resolve(false), 5000)),
    ]).catch(() => false)
  try {
    await seed(p)
    baseline = await snapshot(p)

    const press = async (
      tabName: string,
      control: string,
      option: string | null,
      run: () => Promise<void>,
    ) => {
      const label = option === null ? control : `${control} ▸ ${option}`
      if (SKIP.test(option ?? control)) {
        findings.push({
          tab: tabName,
          control,
          option,
          kind: 'skipped',
          detail: 'on the skip list',
        })
        return
      }
      logged = []
      try {
        await run()
        await p.waitForTimeout(500)
        pressed += 1
        const closed = await closeDialogs(p)
        if (closed.opened) {
          findings.push({
            tab: tabName,
            control,
            option,
            kind: 'dialog',
            detail: closed.stuck
              ? `opens a dialog that Esc does not close: ${closed.stuck}`
              : 'opens a dialog',
          })
        }
        for (const entry of logged) {
          findings.push({
            tab: tabName,
            control,
            option,
            kind: entry.startsWith('page-error') ? 'page-error' : 'console-error',
            detail: entry,
          })
        }
        const after = await snapshot(p)
        if (after !== baseline) {
          // one ⌘Z, then up to two more: the model must come back
          let steps = 0
          let current = after
          while (current !== baseline && steps < 3) {
            const undo = p.locator('button[aria-label^="Undo"]').first()
            if (await undo.isDisabled().catch(() => true)) break
            await undo.click()
            await p.waitForTimeout(450)
            steps += 1
            current = await snapshot(p)
          }
          if (current !== baseline) {
            findings.push({
              tab: tabName,
              control,
              option,
              kind: 'not-undoable',
              detail: `${label}: the model differs after ${steps} undo step(s)`,
            })
            await restart()
          } else if (steps > 1) {
            findings.push({
              tab: tabName,
              control,
              option,
              kind: 'undo-incomplete',
              detail: `${label}: ${steps} undo steps were needed`,
            })
          }
        }
      } catch (error) {
        findings.push({
          tab: tabName,
          control,
          option,
          kind: 'failed',
          detail: `${label}: ${error instanceof Error ? error.message.split('\n').slice(0, 8).join(' | ') : String(error)}`,
        })
        if (findings.filter((finding) => finding.kind === 'failed').length <= 3) {
          await mkdir(OUT, { recursive: true })
          await p
            .screenshot({ path: join(OUT, `failed-${findings.length}.png`) })
            .catch(() => undefined)
        }
        await closeDialogs(p)
        await p.keyboard.press('Escape').catch(() => undefined)
        if (!(await healthy())) await restart()
      }
      await p.evaluate(`${api}.getRange('A1:C4').activate()`).catch(() => undefined)
    }

    for (const [tabIndex, tabName] of TABS.entries()) {
      currentTab = tabName
      // a tab starts on a fresh window too: the last button of the one before can leave a popup open (a note
      // editor, a floating panel) over this tab's first dropdown
      if (tabIndex > 0 && !CONTROLS) await restart()
      await tab(p, tabName)
      await p.waitForTimeout(400)
      const triggerSelector =
        '[data-ribbon-body] .cover-select, [data-ribbon-body] [aria-haspopup="menu"], [data-ribbon-body] [aria-haspopup="listbox"]'
      const triggerCount = await p.locator(triggerSelector).count()
      for (let index = 0; index < triggerCount; index += 1) {
        // every dropdown starts on a fresh window: one choice that wedges it does not take the rest down
        if (index > 0 && !CONTROLS) await restart()
        const trigger = p.locator(triggerSelector).nth(index)
        const control =
          (await trigger.getAttribute('aria-label').catch(() => null)) ??
          (await trigger.innerText().catch(() => `#${index}`)).trim()
        if (CONTROLS && !CONTROLS.test(control)) continue
        if (SKIP.test(control)) {
          findings.push({
            tab: tabName,
            control,
            option: null,
            kind: 'skipped',
            detail: 'on the skip list',
          })
          continue
        }
        // the options of this dropdown
        let options: string[] = []
        try {
          await trigger.click({ timeout: 4000 })
          await p.waitForTimeout(250)
          options = await p
            .locator('.menu-select-drop button, .menu-select-drop [role="option"]')
            .evaluateAll((items) =>
              items
                .filter((item) => (item as HTMLElement).offsetParent !== null)
                .map(
                  (item) =>
                    item.getAttribute('data-style') ??
                    item.getAttribute('data-theme-id') ??
                    (item.textContent ?? '').trim(),
                ),
            )
          await p.keyboard.press('Escape')
          await p.waitForTimeout(150)
        } catch {
          await closeDialogs(p)
          continue
        }
        const picks = [...new Set(options.filter((option) => option !== ''))]
        const chosen =
          picks.length > 12 ? [...picks.slice(0, GALLERY_SAMPLE), picks[picks.length - 1]!] : picks
        for (const option of chosen) {
          await press(tabName, control, option, async () => {
            await p.locator(triggerSelector).nth(index).click({ timeout: 4000 })
            await p.waitForTimeout(200)
            const byAttr = p.locator(
              `.menu-select-drop [data-style="${option}"], .menu-select-drop [data-theme-id="${option}"]`,
            )
            if (await byAttr.count()) await byAttr.first().click({ timeout: 4000 })
            else {
              await p
                .locator('.menu-select-drop button, .menu-select-drop [role="option"]')
                .filter({
                  hasText: new RegExp(`^\\s*${option.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}\\s*$`),
                })
                .first()
                .click({ timeout: 4000 })
            }
          })
        }
      }
      // plain buttons
      const buttonSelector =
        '[data-ribbon-body] button:not(.cover-select):not([aria-haspopup]):not(.rb-fold-btn)'
      const buttonCount = await p.locator(buttonSelector).count()
      for (let index = 0; index < buttonCount; index += 1) {
        if (index % 6 === 5 && !CONTROLS) await restart()
        const button = p.locator(buttonSelector).nth(index)
        const control =
          (await button.getAttribute('aria-label').catch(() => null)) ??
          (await button.getAttribute('data-tip').catch(() => null)) ??
          (await button.innerText().catch(() => `#${index}`)).trim()
        if (CONTROLS && !CONTROLS.test(control)) continue
        if (
          !(await button.isVisible().catch(() => false)) ||
          (await button.isDisabled().catch(() => true))
        )
          continue
        await press(tabName, control, null, async () => {
          await p.locator(buttonSelector).nth(index).click({ timeout: 4000 })
        })
      }
    }
  } finally {
    await mkdir(OUT, { recursive: true })
    await writeFile(join(OUT, 'all.findings.json'), JSON.stringify(findings, null, 2))
    const by = (kind: Finding['kind']) => findings.filter((finding) => finding.kind === kind)
    const lines = [
      `tabs: ${TABS.join(', ')} — ${pressed} choices pressed`,
      ...(
        [
          'page-error',
          'console-error',
          'not-undoable',
          'undo-incomplete',
          'failed',
          'dialog',
          'skipped',
        ] as const
      ).map((kind) => `${kind}: ${by(kind).length}`),
      '',
      ...findings
        .filter((finding) => !['skipped', 'dialog'].includes(finding.kind))
        .map(
          (finding) =>
            `[${finding.kind}] ${finding.tab} › ${finding.control}${finding.option ? ` ▸ ${finding.option}` : ''} — ${finding.detail}`,
        ),
    ]
    await writeFile(join(OUT, 'all.summary.txt'), lines.join('\n'))
    await kill(launched)
  }
  const errors = findings.filter((finding) => finding.kind === 'page-error')
  expect(
    errors,
    errors.map((error) => `${error.tab} › ${error.control}: ${error.detail}`).join('\n'),
  ).toEqual([])
})
