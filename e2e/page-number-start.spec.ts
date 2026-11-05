import { expect, test, type Page } from '@playwright/test'
import { mkdir, readFile, writeFile } from 'node:fs/promises'
import { join, resolve } from 'node:path'
import JSZip from 'jszip'
import { buildDocx } from '../packages/docx-engine/tests/helpers/build-docx'
import { closeAndSaveVideo, launchShell, waitForPageWithUrl } from './helpers'

test.skip(!process.env.GENOFFICE_PAGE_START, 'Opt-in page-five starts at one verification')
const root = resolve('docs/qa/page-number-start-2026-11-04')
async function docFormat(page: Page) {
  await page.locator('.ribbon-tabs .ribbon-tab', { hasText: /^Insert$/ }).click()
  await page.locator('.ribbon-body button[data-tip="Page Number"]').click()
  await page.locator('[data-rb-panel] button', { hasText: /Format/ }).click()
  return page.locator('.modal', { has: page.getByRole('heading', { name: 'Page Number Format' }) })
}
async function hf(page: Page) {
  await page.locator('.ribbon-tabs .ribbon-tab', { hasText: /^Insert$/ }).click()
  await page.locator('.ribbon-body button[data-tip^="Header & footer"]').click()
  return page.locator('.modal', { has: page.getByRole('heading', { name: 'Header and Footer' }) })
}

test('Docs: page five starts at one with a section break, save and reopen', async () => {
  test.setTimeout(120_000)
  await mkdir(root, { recursive: true })
  const file = join(root, 'docs-fifth-page-one.docx')
  const p = (i: number, breakBefore = false) =>
    `<w:p><w:pPr>${breakBefore ? '<w:pageBreakBefore/>' : ''}</w:pPr><w:r><w:t>Physical page ${i}</w:t></w:r></w:p>`
  await writeFile(
    file,
    await buildDocx({
      bodyXml: [p(1), p(2, true), p(3, true), p(4, true), p(5, true), p(6, true)].join(''),
      sectPrExtra: '<w:footerReference w:type="default" r:id="rId20"/><w:pgNumType w:start="9"/>',
      extraRels:
        '<Relationship Id="rId20" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>',
      extraParts: [
        {
          path: 'word/footer1.xml',
          contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml',
          xml: '<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:fldSimple w:instr=" PAGE "><w:r><w:t>9</w:t></w:r></w:fldSimple></w:p></w:ftr>',
        },
      ],
    }),
  )
  let l = await launchShell({
    onboardingSeen: true,
    openFile: file,
    videoDir: 'page-start-docs',
    settings: { autoSave: false },
  })
  try {
    let page = await waitForPageWithUrl(l.app, '://docs/')
    await page.locator('.doc-page.ProseMirror p').filter({ hasText: 'Physical page 5' }).click()
    const dialog = await docFormat(page)
    await dialog.getByTestId('pgnum-physical').check()
    await dialog.getByTestId('pgnum-from-page').fill('5')
    await dialog.locator('input[type="number"]').last().fill('1')
    await dialog.getByRole('button', { name: 'OK', exact: true }).click()
    await expect
      .poll(() => page.evaluate(() => (window as any).__aidocs.editor.state.doc.childCount))
      .toBe(7)
    await page.evaluate(() => (window as any).__aidocs.editor.commands.undo())
    await expect
      .poll(() => page.evaluate(() => (window as any).__aidocs.editor.state.doc.childCount))
      .toBe(6)
    await page.evaluate(() => (window as any).__aidocs.editor.commands.redo())
    await expect
      .poll(() => page.evaluate(() => (window as any).__aidocs.editor.state.doc.childCount))
      .toBe(7)
    expect(await page.evaluate(() => (window as any).__aidocs.save())).toBe(true)
    const xml = await (
      await JSZip.loadAsync(await readFile(file))
    )
      .file('word/document.xml')!
      .async('string')
    expect(xml).toContain('w:start="1"')
    await closeAndSaveVideo(l, 'page-start-docs-saved')
    l = await launchShell({
      onboardingSeen: true,
      openFile: file,
      videoDir: 'page-start-docs-reopened',
    })
    page = await waitForPageWithUrl(l.app, '://docs/')
    await page.waitForFunction(() =>
      (window as any).__aidocs?.editor?.state.doc.textContent.includes('Physical page 6'),
    )
    await page.evaluate(() => (window as any).__openPagePreview())
    await expect(page.locator('.pv-page')).toHaveCount(6)
    const footers = await page
      .locator('.pv-page')
      .evaluateAll((nodes) =>
        nodes.map((p) => p.querySelector('.page-hf-footer')?.textContent?.trim() ?? ''),
      )
    expect(footers).toEqual(['', '', '', '', '1', '2'])
    await page
      .locator('.pv-page')
      .nth(4)
      .screenshot({ path: join(root, 'docs-fifth-page-reopened.png') })
    await writeFile(
      join(root, 'docs-result.json'),
      JSON.stringify(
        {
          physicalPages: 6,
          footerValues: footers,
          nativeField: 'PAGE',
          sectionRestart: 1,
          saveReopen: true,
        },
        null,
        2,
      ),
    )
  } finally {
    await closeAndSaveVideo(l, 'page-start-docs')
  }
})

test('Slides: first four unnumbered, fifth is one, sixth is two after save/reopen', async () => {
  test.setTimeout(180_000)
  await mkdir(root, { recursive: true })
  const file = join(root, 'slides-fifth-page-one.pptx')
  let l = await launchShell({
    onboardingSeen: true,
    videoDir: 'page-start-slides',
    settings: { autoSave: false },
  })
  try {
    await l.page.locator('.quick-card', { hasText: 'AI Slides' }).click()
    let page = await waitForPageWithUrl(l.app, '://slides/')
    await expect(page.locator('.slide-list .thumb')).toHaveCount(1)
    await page.evaluate(async () => {
      const a = (window as any).slidesApi
      for (let i = 0; i < 6; i++) {
        if (i > 0) await a.addBlankSlide({ sourceIndex: i - 1, fitWidthPx: 1280 })
        await a.addElement({
          slideIndex: i,
          kind: 'textbox',
          xPx: 80,
          yPx: 90,
          wPx: 1050,
          hPx: 100,
          fitWidthPx: 1280,
          text: `Physical slide ${i + 1}`,
        })
      }
    })
    await l.app.evaluate(({ dialog }, f) => {
      dialog.showSaveDialog = async () => ({ canceled: false, filePath: f })
    }, file)
    expect(
      (await page.evaluate(() => (window as any).slidesApi.saveAs('slides-fifth-page-one.pptx')))
        .ok,
    ).toBe(true)
    await closeAndSaveVideo(l, 'page-start-slides-created')
    l = await launchShell({
      onboardingSeen: true,
      openFile: file,
      videoDir: 'page-start-slides-settings',
    })
    page = await waitForPageWithUrl(l.app, '://slides/')
    await expect(page.locator('.slide-list .thumb')).toHaveCount(6)
    for (const index of [4]) {
      await page.locator('.slide-list .thumb').nth(index).click()
      const d = await hf(page)
      await d.getByTestId('hf-slide-number').check()
      await d.getByTestId('hf-number-from-slide').fill('5')
      await d.getByTestId('hf-slide-start-at').fill('1')
      await d.getByTestId('hf-apply-current').click()
    }
    expect(
      (await page.evaluate(() => (window as any).slidesApi.getHeaderFooter(4))).numbering
        .startSlide,
    ).toBe(5)
    await page.evaluate(() => (window as any).slidesApi.undo())
    expect(
      (await page.evaluate(() => (window as any).slidesApi.getHeaderFooter(4))).numbering,
    ).toBeNull()
    await page.evaluate(() => (window as any).slidesApi.redo())
    expect(
      (await page.evaluate(() => (window as any).slidesApi.getHeaderFooter(4))).numbering
        .startSlide,
    ).toBe(5)
    expect((await page.evaluate(() => (window as any).slidesApi.save())).ok).toBe(true)
    await closeAndSaveVideo(l, 'page-start-slides-saved')
    l = await launchShell({
      onboardingSeen: true,
      openFile: file,
      videoDir: 'page-start-slides-reopened',
    })
    page = await waitForPageWithUrl(l.app, '://slides/')
    await expect(page.locator('.slide-list .thumb')).toHaveCount(6)
    const values = await page.evaluate(async () => {
      const ss = await (window as any).slidesApi.getRenderSlides()
      return ss.map((s: any) =>
        s.nodes
          .filter((n: any) => n.placeholder === 'sldNum')
          .map((n: any) =>
            n.text?.lines?.map((p: any) => p.runs.map((r: any) => r.text).join('')).join(''),
          )
          .join(''),
      )
    })
    expect(values).toEqual(['', '', '', '', '1', '2'])
    await page.locator('.slide-list .thumb').nth(4).click()
    await page
      .locator('.stage-rel')
      .screenshot({ path: join(root, 'slides-fifth-page-reopened.png') })
    const d = await hf(page)
    await expect(d.getByTestId('hf-slide-start-at')).toHaveValue('1')
    await expect(d.getByTestId('hf-number-from-slide')).toHaveValue('5')
    await page.screenshot({ path: join(root, 'slides-start-at-dialog.png') })
    const pkg = await JSZip.loadAsync(await readFile(file))
    expect(await pkg.file('ppt/presentation.xml')!.async('string')).not.toContain(
      'firstSlideNum="-3"',
    )
    for (const i of [5, 6])
      expect(await pkg.file(`ppt/slides/slide${i}.xml`)!.async('string')).toContain(
        'type="genofficeSlideNumber"',
      )
    await writeFile(
      join(root, 'slides-result.json'),
      JSON.stringify(
        {
          physicalSlides: 6,
          displayedNumbers: values,
          firstSlideNum: 1,
          startSlide: 5,
          nativeField: 'genofficeSlideNumber',
          saveReopen: true,
          uiSteps: 'Enable slide number, start on slide 5, number 1, hide previous slide numbers.',
        },
        null,
        2,
      ),
    )
  } finally {
    await closeAndSaveVideo(l, 'page-start-slides')
  }
})

test('Docs: start on a measured page inside a long paragraph, keep text and fields after save', async () => {
  test.setTimeout(180_000)
  const file = join(root, 'docs-natural-page-start-explicit-font.docx')
  const text = Array.from({ length: 2000 }, (_, i) => `word${i}`).join(' ')
  await writeFile(
    file,
    await buildDocx({
      bodyXml: `<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="exact"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/></w:rPr><w:t>${text}</w:t></w:r></w:p>`,
    }),
  )
  let l = await launchShell({
    videoDir: 'page-number-start-extra',
    onboardingSeen: true,
    openFile: file,
    settings: { autoSave: false },
  })
  try {
    let page = await waitForPageWithUrl(l.app, '://docs/')
    await page.waitForFunction(() =>
      (window as any).__aidocs?.editor?.state.doc.textContent.endsWith('word1999'),
    )
    await page.evaluate(() => (window as any).__openPagePreview())
    await expect.poll(() => page.locator('.pv-page').count()).toBeGreaterThan(2)
    const count = await page.locator('.pv-page').count()
    await page.locator('.pv-toolbar button').last().click()
    const dialog = await docFormat(page)
    await dialog.getByTestId('pgnum-physical').check()
    await dialog.getByTestId('pgnum-from-page').fill('2')
    await dialog.locator('input[type="number"]').last().fill('3')
    await dialog.getByRole('button', { name: 'OK', exact: true }).click()
    await expect(dialog).toHaveCount(0)
    const joined = await page.evaluate(() => {
      const parts: string[] = []
      ;(window as any).__aidocs.editor.state.doc.forEach((n: any) => {
        if (n.type.name === 'docParagraph') parts.push(n.textContent)
      })
      return parts.join('')
    })
    expect(joined).toBe(text)
    expect(await page.evaluate(() => (window as any).__aidocs.save())).toBe(true)
    await closeAndSaveVideo(l, 'natural-page-number-saved')
    l = await launchShell({
      videoDir: 'page-number-start-extra',
      onboardingSeen: true,
      openFile: file,
      settings: { autoSave: false },
    })
    page = await waitForPageWithUrl(l.app, '://docs/')
    await page.waitForFunction(() =>
      (window as any).__aidocs?.editor?.state.doc.textContent.endsWith('word1999'),
    )
    await page.evaluate(() => (window as any).__openPagePreview())
    await expect(page.locator('.pv-page')).toHaveCount(count)
    expect(
      await page
        .locator('.pv-page')
        .evaluateAll((ns) =>
          ns.map((n) => n.querySelector('.page-hf-footer')?.textContent?.trim() ?? ''),
        ),
    ).toEqual(Array.from({ length: count }, (_, i) => (i === 0 ? '' : String(i + 2))))
    await writeFile(
      join(root, 'natural-docs-result.json'),
      JSON.stringify({ count, textPreserved: true, page2: 3, nativeSection: true }),
    )
  } finally {
    await closeAndSaveVideo(l, 'natural-page-number')
  }
})

test('Office save roundtrip: native fields remain editable in both GenOffice editors', async () => {
  test.setTimeout(180_000)
  for (const kind of ['docs', 'slides'] as const) {
    const ext = kind === 'docs' ? 'docx' : 'pptx'
    const source = join(root, `${kind}-numbering-office-roundtrip.${ext}`)
    const file = join(root, `${kind}-office-roundtrip-reedit.${ext}`)
    await writeFile(file, await readFile(source))
    let l = await launchShell({
      videoDir: 'page-number-start-extra',
      onboardingSeen: true,
      openFile: file,
      settings: { autoSave: false },
    })
    try {
      let page = await waitForPageWithUrl(l.app, `://${kind}/`)
      if (kind === 'docs') {
        await page.waitForFunction(() =>
          (window as any).__aidocs?.editor?.state.doc.textContent.includes('Physical page 6'),
        )
        await page
          .locator('.missing-fonts-confirm-btn')
          .waitFor({ state: 'visible', timeout: 5000 })
          .then(() => page.locator('.missing-fonts-confirm-btn').click())
          .catch(() => {})
        await page.locator('.doc-page.ProseMirror p').filter({ hasText: 'Physical page 5' }).click()
        const d = await docFormat(page)
        await d.getByText('Start at', { exact: true }).click()
        await d.locator('input[type="number"]').last().fill('7')
        await d.getByRole('button', { name: 'OK', exact: true }).click()
        await page.evaluate(() => (window as any).__aidocs.editor.commands.undo())
        await page.evaluate(() => (window as any).__aidocs.editor.commands.redo())
        expect(await page.evaluate(() => (window as any).__aidocs.save())).toBe(true)
      } else {
        await expect(page.locator('.slide-list .thumb')).toHaveCount(6)
        const state = await page.evaluate(() => (window as any).slidesApi.getHeaderFooter(4))
        expect(state.numbering).toEqual({
          startSlide: 5,
          startNumber: 1,
          hideBefore: true,
          enabled: true,
        })
        await page.locator('.slide-list .thumb').nth(4).click()
        const d = await hf(page)
        await d.getByTestId('hf-slide-start-at').fill('7')
        await d.getByTestId('hf-apply-all').click()
        await page.evaluate(() => (window as any).slidesApi.undo())
        await page.evaluate(() => (window as any).slidesApi.redo())
        expect((await page.evaluate(() => (window as any).slidesApi.save())).ok).toBe(true)
      }
      await closeAndSaveVideo(l, `${kind}-office-reedit-saved`)
      l = await launchShell({
        videoDir: 'page-number-start-extra',
        onboardingSeen: true,
        openFile: file,
        settings: { autoSave: false },
      })
      page = await waitForPageWithUrl(l.app, `://${kind}/`)
      if (kind === 'docs') {
        await page.waitForFunction(() =>
          (window as any).__aidocs?.editor?.state.doc.textContent.includes('Physical page 6'),
        )
        await page.evaluate(() => (window as any).__openPagePreview())
        await expect(page.locator('.pv-page')).toHaveCount(6)
        expect(
          await page
            .locator('.pv-page')
            .evaluateAll((ns) =>
              ns.map((n) => n.querySelector('.page-hf-footer')?.textContent?.trim() ?? ''),
            ),
        ).toEqual(['', '', '', '', '7', '8'])
        const zip = await JSZip.loadAsync(await readFile(file))
        expect(
          await Promise.all(
            Object.keys(zip.files)
              .filter((n) => /^word\/footer\d+\.xml$/.test(n))
              .map((n) => zip.file(n)!.async('string')),
          ).then((xs) => xs.join('')),
        ).toMatch(/PAGE/)
      } else {
        await expect(page.locator('.slide-list .thumb')).toHaveCount(6)
        expect(
          (await page.evaluate(() => (window as any).slidesApi.getHeaderFooter(4))).numbering
            .startNumber,
        ).toBe(7)
        const values = await page.evaluate(async () =>
          (await (window as any).slidesApi.getRenderSlides()).map((s: any) =>
            s.nodes
              .filter((n: any) => n.placeholder === 'sldNum')
              .map((n: any) =>
                n.text?.lines.map((p: any) => p.runs.map((r: any) => r.text).join('')).join(''),
              )
              .join(''),
          ),
        )
        expect(values).toEqual(['', '', '', '', '7', '8'])
      }
    } finally {
      await closeAndSaveVideo(l, `${kind}-office-reedit`)
    }
  }
})
