import { test, expect, type Page } from '@playwright/test'
import { mkdir, readFile, writeFile, copyFile } from 'node:fs/promises'
import { join, resolve } from 'node:path'
import JSZip from 'jszip'
import { PNG } from 'pngjs'
import { closeAndSaveVideo, launchShell, waitForPageWithUrl } from './helpers'

test.skip(!process.env.GENOFFICE_EDIT_POSITION, 'Opt-in text edit positioning audit')
const root = resolve('docs/qa/text-edit-position-2026-11-04')
const cases = [
  { name: 'small-centered-number', align: 'ctr', w: 26 },
  { name: 'small-right-label', align: 'r', w: 26 },
  { name: 'small-left-label', align: 'l', w: 26 },
  { name: 'small-nowrap-centered', align: 'ctr', w: 26, wrap: 'none' },
  { name: 'small-nowrap-right', align: 'r', w: 26, wrap: 'none' },
  { name: 'nowrap-overflow-centered', align: 'ctr', w: 18, wrap: 'none', text: 'ABCDE' },
  { name: 'rotated-small-label', align: 'ctr', w: 26, rotation: 25 },
  { name: 'normal-title', align: 'ctr', w: 250, text: 'A normal title', size: 24 },
  {
    name: 'wrapped-body',
    align: 'l',
    w: 180,
    text: 'Body text and mixed 中文文字 that wraps over several lines.',
    h: 100,
  },
  { name: 'group-small-label', align: 'ctr', w: 26, group: true },
  { name: 'small-table-cell', align: 'ctr', w: 26, table: true },
  { name: 'native-page-number', align: 'ctr', w: 26, field: 'slidenum' },
  { name: 'offset-page-number', align: 'ctr', w: 26, field: 'genofficeSlideNumber' },
]
const emu = (px: number) => Math.round(px * 9525)
function shape(c: (typeof cases)[number], id = 20, x = 500, y = 220) {
  const r = `<a:rPr sz="${(c.size ?? 16) * 100}"><a:solidFill><a:srgbClr val="7F2399"/></a:solidFill><a:latin typeface="Arial"/></a:rPr><a:t>${c.text ?? '2'}</a:t>`
  const para = `<a:p><a:pPr algn="${c.align}"/>${c.field ? `<a:fld id="{93DF4881-9D55-4E14-8B2E-632D56874A31}" type="${c.field}">${r}</a:fld>` : `<a:r>${r}</a:r>`}</a:p>`
  const body = `<a:bodyPr wrap="${c.wrap ?? 'square'}" anchor="ctr" lIns="34250" rIns="34250" tIns="34250" bIns="34250"><a:noAutofit/></a:bodyPr><a:lstStyle/>${para}`
  const xfrm = `<a:xfrm rot="${(c.rotation ?? 0) * 60000}"><a:off x="${emu(x)}" y="${emu(y)}"/><a:ext cx="${emu(c.w)}" cy="${emu(c.h ?? 36)}"/></a:xfrm>`
  if (c.table)
    return `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="20" name="${c.name}"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="${emu(x)}" y="${emu(y)}"/><a:ext cx="${emu(c.w)}" cy="${emu(36)}"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr/><a:tblGrid><a:gridCol w="${emu(c.w)}"/></a:tblGrid><a:tr h="${emu(36)}"><a:tc><a:txBody>${body}</a:txBody><a:tcPr anchor="ctr" marL="34250" marR="34250" marT="34250" marB="34250"/></a:tc></a:tr></a:tbl></a:graphicData></a:graphic></p:graphicFrame>`
  const sp = `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${c.name}"/><p:cNvSpPr txBox="1"/><p:nvPr>${c.field ? '<p:ph type="sldNum" idx="12"/>' : ''}</p:nvPr></p:nvSpPr><p:spPr>${xfrm}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></p:spPr><p:txBody>${body}</p:txBody></p:sp>`
  if (!c.group) return sp
  const child = shape({ ...c, group: false }, 21, 0, 0)
  return `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="22" name="Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="${emu(x)}" y="${emu(y)}"/><a:ext cx="${emu(100)}" cy="${emu(100)}"/><a:chOff x="0" y="0"/><a:chExt cx="${emu(100)}" cy="${emu(100)}"/></a:xfrm></p:grpSpPr>${child}</p:grpSp>`
}
async function fixture(path: string) {
  const z = await JSZip.loadAsync(
    await readFile(resolve('docs/qa/page-number-start-2026-11-04/slides-fifth-page-one.pptx')),
  )
  const base = await z.file('ppt/slides/slide1.xml')!.async('string')
  const start = base.indexOf('<p:spTree>'),
    end = base.indexOf('</p:spTree>', start)
  const head = base
    .slice(start, end)
    .match(/^<p:spTree>[\s\S]*?<p:grpSpPr(?:\s[^>]*)?(?:\/>|>[\s\S]*?<\/p:grpSpPr>)/)![0]
  const rel = await z.file('ppt/slides/_rels/slide1.xml.rels')!.async('string')
  let pres = await z.file('ppt/presentation.xml')!.async('string')
  pres = pres
    .replace(
      /<p:sldIdLst>[\s\S]*?<\/p:sldIdLst>/,
      `<p:sldIdLst>${cases.map((_, i) => `<p:sldId id="${256 + i}" r:id="rIdAudit${i}"/>`).join('')}</p:sldIdLst>`,
    )
    .replace(/<p:extLst>[\s\S]*?<\/p:extLst>/, '')
  z.file('ppt/presentation.xml', pres)
  let rels = await z.file('ppt/_rels/presentation.xml.rels')!.async('string')
  rels = rels
    .replace(/<Relationship\b[^>]*Type="[^"]*\/slide"[^>]*\/>/g, '')
    .replace(
      '</Relationships>',
      cases
        .map(
          (_, i) =>
            `<Relationship Id="rIdAudit${i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${i + 1}.xml"/>`,
        )
        .join('') + '</Relationships>',
    )
  z.file('ppt/_rels/presentation.xml.rels', rels)
  let ct = await z.file('[Content_Types].xml')!.async('string')
  for (const [i, c] of cases.entries()) {
    z.file(`ppt/slides/slide${i + 1}.xml`, base.slice(0, start) + head + shape(c) + base.slice(end))
    z.file(`ppt/slides/_rels/slide${i + 1}.xml.rels`, rel)
    if (!ct.includes(`PartName="/ppt/slides/slide${i + 1}.xml"`))
      ct = ct.replace(
        '</Types>',
        `<Override PartName="/ppt/slides/slide${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>`,
      )
  }
  z.file('[Content_Types].xml', ct)
  await writeFile(path, await z.generateAsync({ type: 'nodebuffer' }))
}
async function ink(page: Page) {
  const stage = (await page.locator('.stage-rel').boundingBox())!
  const p = PNG.sync.read(await page.locator('.stage-rel').screenshot())
  let x0 = p.width,
    y0 = p.height,
    x1 = -1,
    y1 = -1
  for (let y = 0; y < p.height; y++)
    for (let x = 0; x < p.width; x++) {
      const k = (y * p.width + x) * 4,
        r = p.data[k]!,
        g = p.data[k + 1]!,
        b = p.data[k + 2]!
      if (r > 60 && b > 80 && r > g * 1.5 && b > g * 1.7 && b > r * 1.06) {
        x0 = Math.min(x0, x)
        y0 = Math.min(y0, y)
        x1 = Math.max(x1, x)
        y1 = Math.max(y1, y)
      }
    }
  expect(x1).toBeGreaterThanOrEqual(0)
  return {
    x: (x0 * stage.width) / p.width,
    y: (y0 * stage.height) / p.height,
    w: ((x1 - x0 + 1) * stage.width) / p.width,
    h: ((y1 - y0 + 1) * stage.height) / p.height,
  }
}
async function targetPoint(page: Page, index: number) {
  const slide = await page.evaluate(
    async (i) => (await (window as any).slidesApi.getRenderSlides())[i],
    index,
  )
  const stage = (await page.locator('.stage-rel').boundingBox())!
  const node = slide.nodes.find((n: any) => !n.decoration)
  const x = node.box.x + (node.type === 'group' ? 13 : node.box.w / 2),
    y = node.box.y + (node.type === 'group' ? 18 : node.box.h / 2)
  return {
    x: stage.x + (x * stage.width) / slide.widthPx,
    y: stage.y + (y * stage.height) / slide.heightPx,
    node,
  }
}
test('audit edit entry across small shapes, tables, groups and native fields', async () => {
  test.setTimeout(240_000)
  await mkdir(root, { recursive: true })
  const file = join(root, 'edit-position-cases.pptx')
  await fixture(file)
  const l = await launchShell({
    onboardingSeen: true,
    openFile: file,
    videoDir: 'edit-position-audit',
    settings: { autoSave: false },
  })
  const results: any[] = []
  try {
    const page = await waitForPageWithUrl(l.app, '://slides/')
    await expect(page.locator('.slide-list .thumb')).toHaveCount(cases.length)
    for (const [i, c] of cases.entries()) {
      await page.locator('.slide-list .thumb').nth(i).click()
      await page.waitForFunction(
        () => !(window as any).document.querySelector('.slide-text-editor'),
      )
      const pt = await targetPoint(page, i)
      const before = await ink(page)
      if (c.group) {
        await page.mouse.dblclick(pt.x, pt.y)
        await page.mouse.click(pt.x, pt.y)
        await page.keyboard.press('F2')
      } else if (c.table) await page.mouse.dblclick(pt.x, pt.y)
      else {
        await page.mouse.click(pt.x, pt.y)
        await page.keyboard.press('F2')
      }
      await expect(page.locator('.slide-text-editor')).toBeVisible()
      const after = await ink(page)
      const layout = await page.locator('.slide-text-editor').evaluate((el) => ({
        width: el.getBoundingClientRect().width,
        frame: el.parentElement!.getBoundingClientRect().width,
        html: el.innerHTML,
      }))
      results.push({
        name: c.name,
        before,
        after,
        dx: after.x - before.x,
        dy: after.y - before.y,
        editorWidth: layout.width,
        frameWidth: layout.frame,
      })
      if (Math.abs(after.x - before.x) > 1.5 || Math.abs(after.y - before.y) > 1.5)
        await page.locator('.stage-rel').screenshot({ path: join(root, c.name + '-editing.png') })
      await page.keyboard.press('Escape')
      await expect(page.locator('.slide-text-editor')).toHaveCount(0)
      const exit = await ink(page)
      expect(Math.abs(exit.x - before.x)).toBeLessThanOrEqual(1.5)
      expect(Math.abs(exit.y - before.y)).toBeLessThanOrEqual(1.5)
    }
    await writeFile(
      join(root, process.env.GENOFFICE_EDIT_BASELINE ? 'baseline.json' : 'audit-results.json'),
      JSON.stringify(results, null, 2),
    )
    console.log(
      JSON.stringify(
        results.map(({ name, dx, dy, editorWidth, frameWidth }) => ({
          name,
          dx,
          dy,
          editorWidth,
          frameWidth,
        })),
        null,
        2,
      ),
    )
    expect(results.filter((r) => Math.abs(r.dx) > 1.5 || Math.abs(r.dy) > 1.5)).toEqual([])
  } finally {
    await closeAndSaveVideo(l, 'edit-position-audit')
  }
})

test('page number font color survives leaving editing, Undo/Redo, save and reopen', async () => {
  test.setTimeout(180_000)
  await mkdir(root, { recursive: true })
  const file = join(root, 'page-number-color.pptx')
  await fixture(file)
  let l = await launchShell({
    onboardingSeen: true,
    openFile: file,
    videoDir: 'edit-field-color',
    settings: { autoSave: false },
  })
  const results: any[] = []
  try {
    let page = await waitForPageWithUrl(l.app, '://slides/')
    for (const i of [11, 12]) {
      let previousColor = '#7F2399'
      await page.locator('.slide-list .thumb').nth(i).click()
      for (const color of ['#EE602F', '#FFFFFF', '#7F2399']) {
        const pt = await targetPoint(page, i)
        await page.mouse.click(pt.x, pt.y)
        await page.keyboard.press('F2')
        await expect(page.locator('.slide-text-editor')).toBeVisible()
        await page.locator('.slide-text-editor').evaluate((el, color) => {
          const r = document.createRange()
          r.selectNodeContents(el)
          const s = window.getSelection()!
          s.removeAllRanges()
          s.addRange(r)
          document.execCommand('styleWithCSS', false, 'true')
          document.execCommand('foreColor', false, color)
        }, color)
        const html = await page.locator('.slide-text-editor').innerHTML()
        const stage = (await page.locator('.stage-rel').boundingBox())!
        await page.mouse.click(stage.x + 20, stage.y + 20)
        await expect(page.locator('.slide-text-editor')).toHaveCount(0)
        const savedColor = await page.evaluate(
          async (i) =>
            (await (window as any).slidesApi.getRenderSlides())[i].nodes.find(
              (n: any) => !n.decoration,
            ).text.lines[0].runs[0].color,
          i,
        )
        results.push({ index: i, color, savedColor, html })
        expect(savedColor.toUpperCase()).toBe(color)
        const readColor = () =>
          page.evaluate(
            async (i) =>
              (await (window as any).slidesApi.getRenderSlides())[i].nodes.find(
                (n: any) => !n.decoration,
              ).text.lines[0].runs[0].color,
            i,
          )
        await page.keyboard.press('ControlOrMeta+z')
        await expect.poll(readColor).toBe(previousColor)
        await page.keyboard.press('ControlOrMeta+Shift+z')
        await expect.poll(readColor).toBe(color)
        previousColor = color
      }
    }
    expect((await page.evaluate(() => (window as any).slidesApi.save())).ok).toBe(true)
    const zip = await JSZip.loadAsync(await readFile(file))
    for (const [n, type] of [
      [12, 'slidenum'],
      [13, 'genofficeSlideNumber'],
    ] as const) {
      const xml = await zip.file(`ppt/slides/slide${n}.xml`)!.async('string')
      expect(xml).toContain(`type="${type}"`)
      expect(xml).toContain('val="7F2399"')
    }
    await closeAndSaveVideo(l, 'field-color-saved')
    l = await launchShell({
      onboardingSeen: true,
      openFile: file,
      videoDir: 'edit-field-color-reopened',
      settings: { autoSave: false },
    })
    page = await waitForPageWithUrl(l.app, '://slides/')
    await expect(page.locator('.slide-list .thumb')).toHaveCount(cases.length)
    for (const i of [11, 12]) {
      const c = await page.evaluate(
        async (i) =>
          (await (window as any).slidesApi.getRenderSlides())[i].nodes.find(
            (n: any) => !n.decoration,
          ).text.lines[0].runs[0].color,
        i,
      )
      expect(c.toUpperCase()).toBe('#7F2399')
    }
  } finally {
    await writeFile(join(root, 'field-color-results.json'), JSON.stringify(results, null, 2))
    await closeAndSaveVideo(l, 'field-color')
  }
})

test('page-number palette color applies to the entire field when the caret is collapsed', async () => {
  test.setTimeout(90_000)
  const file = join(root, 'page-number-caret-color.pptx')
  await fixture(file)
  const l = await launchShell({
    onboardingSeen: true,
    openFile: file,
    videoDir: 'page-number-caret-color',
    settings: { autoSave: false },
  })
  try {
    const page = await waitForPageWithUrl(l.app, '://slides/')
    await page.locator('.slide-list .thumb').nth(12).click()
    const pt = await targetPoint(page, 12)
    await page.mouse.click(pt.x, pt.y)
    await page.keyboard.press('F2')
    await expect(page.locator('.slide-text-editor')).toBeVisible()
    expect(await page.evaluate(() => window.getSelection()?.isCollapsed)).toBe(true)
    await page.locator('.ribbon-body button[data-tip="Font Color"]').click()
    await page.locator('.rb-color-pop .gcp-swatch[title="White"]').click()
    const stage = (await page.locator('.stage-rel').boundingBox())!
    await page.mouse.click(stage.x + 20, stage.y + 20)
    await expect(page.locator('.slide-text-editor')).toHaveCount(0)
    const color = await page.evaluate(
      async () =>
        (await (window as any).slidesApi.getRenderSlides())[12].nodes.find(
          (n: any) => !n.decoration,
        ).text.lines[0].runs[0].color,
    )
    expect(color.toUpperCase()).toBe('#FFFFFF')
  } finally {
    await closeAndSaveVideo(l, 'page-number-caret-color')
  }
})

test('reported Rockwell page badge retains position, font color and native field', async () => {
  test.skip(!process.env.GENOFFICE_EDIT_REPRO_PATH, 'Requires the local reported deck')
  test.setTimeout(150_000)
  await mkdir(root, { recursive: true })
  const file = join(root, 'reported-badge-copy.pptx')
  await copyFile(process.env.GENOFFICE_EDIT_REPRO_PATH!, file)
  // Copy the reported layout's actual Rockwell badge onto slide 2 without recreating
  // it with generic header/footer defaults. The source file is never edited.
  const input = await JSZip.loadAsync(await readFile(file))
  const layout = await input.file('ppt/slideLayouts/slideLayout13.xml')!.async('string')
  const badge = (layout.match(/<p:sp\b[\s\S]*?<\/p:sp>/g) ?? [])
    .find((sp) => /<p:ph\b[^>]*type="sldNum"/.test(sp))!
    .replace(/(<p:cNvPr[^>]*\bid=")[^"]+"/, '$19001"')
  const original = await input.file('ppt/slides/slide2.xml')!.async('string')
  input.file(
    'ppt/slides/slide2.xml',
    original
      .replace(/<p:sp\b[\s\S]*?<\/p:sp>/g, (sp) => (/<p:ph\b[^>]*type="sldNum"/.test(sp) ? '' : sp))
      .replace('</p:spTree>', badge + '</p:spTree>'),
  )
  let pres = await input.file('ppt/presentation.xml')!.async('string')
  pres = pres.replace(/\bfirstSlideNum="[^"]*"/, 'firstSlideNum="1"')
  // Isolate the standard native-field case from any in-progress offset settings
  // in the user's source deck; those are independently covered by the fixture suite.
  pres = pres.replace(
    /<p:ext\b[^>]*uri="\{581A1EBB-85C7-4B55-A5FD-43AA26531141\}"[^>]*>[\s\S]*?<\/p:ext>/,
    '',
  )
  input.file('ppt/presentation.xml', pres)
  await writeFile(file, await input.generateAsync({ type: 'nodebuffer' }))
  let l = await launchShell({
    onboardingSeen: true,
    openFile: file,
    videoDir: 'reported-badge',
    settings: { autoSave: false },
  })
  try {
    let page = await waitForPageWithUrl(l.app, '://slides/')
    await expect(page.locator('.slide-list .thumb')).toHaveCount(25)
    await page.locator('.slide-list .thumb').nth(1).click()
    const sl = await page.evaluate(
      async () => (await (window as any).slidesApi.getRenderSlides())[1],
    )
    const n = sl.nodes.find((n: any) => !n.decoration && n.placeholder === 'sldNum'),
      r = (await page.locator('.stage-rel').boundingBox())!
    expect(n).toBeTruthy()
    const k = r.width / sl.widthPx,
      point = { x: r.x + (n.box.x + n.box.w / 2) * k, y: r.y + (n.box.y + n.box.h / 2) * k }
    const clip = {
      x: Math.max(0, point.x - 45),
      y: Math.max(0, point.y - 45),
      width: 90,
      height: 90,
    }
    await page.screenshot({ path: join(root, 'reported-badge-before.png'), clip })
    await page.mouse.click(point.x, point.y)
    await page.keyboard.press('F2')
    await expect(page.locator('.slide-text-editor')).toBeVisible()
    const dimensions = await page.locator('.slide-text-editor').evaluate((el) => ({
      editor: el.getBoundingClientRect().width,
      frame: el.parentElement!.getBoundingClientRect().width,
    }))
    expect(Math.abs(dimensions.editor - dimensions.frame)).toBeLessThan(1)
    await page.screenshot({ path: join(root, 'reported-badge-editing.png'), clip })
    await page.locator('.ribbon-body button[data-tip="Font Color"]').click()
    await page.locator('.rb-color-pop .gcp-theme-base .gcp-swatch[title="Blue"]').click()
    await page.mouse.click(r.x + 20, r.y + 20)
    await expect(page.locator('.slide-text-editor')).toHaveCount(0)
    const readColor = () =>
      page.evaluate(
        async () =>
          (await (window as any).slidesApi.getRenderSlides())[1].nodes.find(
            (n: any) => !n.decoration && n.placeholder === 'sldNum',
          ).text.lines[0].runs[0].color,
      )
    await expect.poll(readColor).toBe('#156082')
    const readBox = () =>
      page.evaluate(
        async () =>
          (await (window as any).slidesApi.getRenderSlides())[1].nodes.find(
            (node: any) => !node.decoration && node.placeholder === 'sldNum',
          ).box,
      )
    expect(await readBox()).toEqual(n.box)
    await page.keyboard.press('ControlOrMeta+z')
    await expect.poll(readColor).toBe(n.text.lines[0].runs[0].color)
    await page.keyboard.press('ControlOrMeta+Shift+z')
    await expect.poll(readColor).toBe('#156082')
    // The selected-element font API is the other ribbon formatting path.
    await page.evaluate(
      async (id) =>
        (window as any).slidesApi.setElementFont({
          slideIndex: 1,
          sourceIds: [id],
          color: '#FFFFFF',
        }),
      n.sourceId,
    )
    await expect.poll(readColor).toBe('#FFFFFF')
    expect(await readBox()).toEqual(n.box)
    await page.evaluate(() => (window as any).slidesApi.undo())
    await expect.poll(readColor).toBe('#156082')
    expect((await page.evaluate(() => (window as any).slidesApi.save())).ok).toBe(true)
    await closeAndSaveVideo(l, 'reported-badge-saved')
    l = await launchShell({
      onboardingSeen: true,
      openFile: file,
      videoDir: 'reported-badge-reopen',
      settings: { autoSave: false },
    })
    page = await waitForPageWithUrl(l.app, '://slides/')
    await page.locator('.slide-list .thumb').nth(1).click()
    await expect.poll(readColor).toBe('#156082')
    const zip = await JSZip.loadAsync(await readFile(file))
    expect(await zip.file('ppt/slides/slide2.xml')!.async('string')).toContain('type="slidenum"')
    await writeFile(
      join(root, 'reported-badge-result.json'),
      JSON.stringify(
        {
          font: n.text.lines[0].runs[0].fontFamily,
          box: n.box,
          dimensions,
          color: '#156082',
          nativeField: true,
          colorOnlyPreservesBox: true,
          undoRedo: true,
          saveReopen: true,
        },
        null,
        2,
      ),
    )
  } finally {
    await closeAndSaveVideo(l, 'reported-badge')
  }
})
