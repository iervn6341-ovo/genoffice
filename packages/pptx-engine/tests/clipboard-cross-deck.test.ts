import { describe, it, expect } from 'vitest'
import { readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { dirname, join } from 'node:path'
import JSZip from 'jszip'
import {
  addChart,
  addElement,
  addPicture,
  appendRawElements,
  buildNativeClipboard,
  buildSlidesClipboardPackage,
  openSlidesClipboardPackage,
  copyElementData,
  copySlide,
  elementPlainText,
  elementSpid,
  getSlideAnimations,
  getSlideNotes,
  setSlideAnimations,
  pasteSlide,
  setSlideNotes,
  ensureTableStylePart,
  htmlTableToItem,
  htmlTextToItem,
  tableToHtml,
  openPptx,
  parseGvml,
  parseTableClip,
  pasteElements,
  savePptx,
  type OpenedPptx,
} from '../src/index'
import type { TableElement, TextElement } from '../src/types'

const here = dirname(fileURLToPath(import.meta.url))
const fx = (name: string) => readFileSync(join(here, 'fixtures', name))
const clip = (name: string) =>
  new Uint8Array(readFileSync(join(here, 'fixtures', 'ppt-clip', name)))
const ZERO = { dx: 0, dy: 0 }
const EMU = 12700

const PNG = Uint8Array.from(
  Buffer.from(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==',
    'base64',
  ),
)

const STYLE = '{11111111-2222-3333-4444-555555555555}'
const STYLE_DEF = (accent: string) =>
  `<a:tblStyle styleId="${STYLE}" styleName="Test"><a:wholeTbl><a:tcTxStyle><a:fontRef idx="minor"/><a:schemeClr val="dk1"/></a:tcTxStyle><a:tcStyle><a:tcBdr/><a:fill><a:solidFill><a:schemeClr val="${accent}"/></a:solidFill></a:fill></a:tcStyle></a:wholeTbl></a:tblStyle>`

const cell = (t: string) =>
  `<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" sz="1200"/><a:t>${t}</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>`
const tableXml = (id: number, styleId: string) =>
  `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${id}" name="T${id}"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="${40 * EMU}" y="${200 * EMU}"/><a:ext cx="${200 * EMU}" cy="${60 * EMU}"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr firstRow="1" bandRow="1"><a:tableStyleId>${styleId}</a:tableStyleId></a:tblPr><a:tblGrid><a:gridCol w="${100 * EMU}"/><a:gridCol w="${100 * EMU}"/></a:tblGrid><a:tr h="${30 * EMU}">${cell('a')}${cell('b')}</a:tr><a:tr h="${30 * EMU}">${cell('c')}${cell('d')}</a:tr></a:tbl></a:graphicData></a:graphic></p:graphicFrame>`

const cNvPrIds = (o: OpenedPptx, slideIndex: number) =>
  o.deck.slides[slideIndex]!.elements.flatMap((e) => [
    ...e.anchor.originalXml.matchAll(/<p:cNvPr\s[^>]*\bid="(\d+)"/g),
  ]).map((m) => m[1]!)

async function sourceDeck(): Promise<OpenedPptx> {
  return openPptx(fx('01_standard_business.pptx'))
}

describe('cross-deck paste: table styles', () => {
  it('a custom table style travels with the table (target tableStyles.xml lacks it)', async () => {
    const src = await sourceDeck()
    ensureTableStylePart(src, STYLE, STYLE_DEF('accent2'))
    appendRawElements(src, 0, [tableXml(900, STYLE)])
    const el = src.deck.slides[0]!.elements.find((e) => e.type === 'table') as TableElement
    const item = copyElementData(src, src.deck.slides[0]!, el)
    expect(item.tableStyles?.[0]?.id).toBe(STYLE)
    expect(item.tableStyles?.[0]?.xml).toContain('styleName="Test"')

    const dst = await sourceDeck()
    // the destination already has a populated tableStyles part without this id
    ensureTableStylePart(
      dst,
      '{AAAAAAAA-0000-0000-0000-000000000000}',
      STYLE_DEF('accent1').replace(STYLE, '{AAAAAAAA-0000-0000-0000-000000000000}'),
    )
    const r = pasteElements(dst, 0, [item], ZERO)!
    const pasted = r.slide.elements.at(-1) as TableElement
    expect(pasted.type).toBe('table')
    expect(pasted.styleId).toBe(STYLE)
    expect(dst.archive.readText('ppt/tableStyles.xml')).toContain(`styleId="${STYLE}"`)
    // the cells resolve a fill from the landed style (accent2 of the destination theme)
    expect(pasted.rows[0]![0]!.fill).toBeDefined()

    const reopened = await openPptx(await savePptx(dst))
    const again = reopened.deck.slides[0]!.elements.at(-1) as TableElement
    expect(again.rows[0]![0]!.fill).toEqual(pasted.rows[0]![0]!.fill)
  })

  it('a style id that exists with a different definition is re-minted, not overwritten', async () => {
    const src = await sourceDeck()
    ensureTableStylePart(src, STYLE, STYLE_DEF('accent2'))
    appendRawElements(src, 0, [tableXml(900, STYLE)])
    const item = copyElementData(
      src,
      src.deck.slides[0]!,
      src.deck.slides[0]!.elements.find((e) => e.type === 'table')!,
    )
    const dst = await sourceDeck()
    ensureTableStylePart(dst, STYLE, STYLE_DEF('accent5'))
    const r = pasteElements(dst, 0, [item], ZERO)!
    const pasted = r.slide.elements.at(-1) as TableElement
    expect(pasted.styleId).not.toBe(STYLE)
    const ts = dst.archive.readText('ppt/tableStyles.xml')!
    expect(ts).toContain(`styleId="${STYLE}"`)
    expect(ts).toContain(`styleId="${pasted.styleId}"`)
    expect(ts).toContain('accent2')
    expect(ts).toContain('accent5')
  })

  it('a built-in style is regenerated in a target whose tableStyles part is populated', async () => {
    const builtin = '{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}'
    const src = await sourceDeck()
    appendRawElements(src, 0, [tableXml(900, builtin)])
    const item = copyElementData(
      src,
      src.deck.slides[0]!,
      src.deck.slides[0]!.elements.find((e) => e.type === 'table')!,
    )
    expect(item.tableStyles).toEqual([{ id: builtin }])
    const dst = await sourceDeck()
    ensureTableStylePart(dst, STYLE, STYLE_DEF('accent1'))
    const r = pasteElements(dst, 0, [item], ZERO)!
    const pasted = r.slide.elements.at(-1) as TableElement
    expect(dst.archive.readText('ppt/tableStyles.xml')).toContain(`styleId="${builtin}"`)
    expect(pasted.rows[0]![0]!.fill).toBeDefined()
  })
})

describe('paste: ids, connectors, relationships', () => {
  it('connector endpoints follow the renumbered shapes; unrelated references are dropped', async () => {
    const src = await sourceDeck()
    const s = src.deck.slides[0]!
    const box = (id: number, name: string, x: number) =>
      `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${x * EMU}" y="${50 * EMU}"/><a:ext cx="${40 * EMU}" cy="${30 * EMU}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:sp>`
    const cxn = (id: number, st: number, en: number) =>
      `<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="${id}" name="C${id}"/><p:cNvCxnSpPr><a:stCxn id="${st}" idx="3"/><a:endCxn id="${en}" idx="1"/></p:cNvCxnSpPr><p:nvPr/></p:nvCxnSpPr><p:spPr><a:xfrm><a:off x="${80 * EMU}" y="${60 * EMU}"/><a:ext cx="${60 * EMU}" cy="${10 * EMU}"/></a:xfrm><a:prstGeom prst="straightConnector1"><a:avLst/></a:prstGeom></p:spPr></p:cxnSp>`
    appendRawElements(src, 0, [box(901, 'A', 40), box(902, 'B', 140), cxn(903, 901, 902)])
    const items = s.elements
      .filter((e) => /id="90[123]"/.test(e.anchor.originalXml))
      .map((e) => copyElementData(src, s, e))
    expect(items).toHaveLength(3)

    const dst = await sourceDeck()
    const r = pasteElements(dst, 0, items, ZERO)!
    const created = r.slide.elements.slice(-3)
    const ids = created.map((e) => /<p:cNvPr\s[^>]*\bid="(\d+)"/.exec(e.anchor.originalXml)![1]!)
    const cx = created[2]!.anchor.originalXml
    expect(cx).toContain(`<a:stCxn id="${ids[0]}"`)
    expect(cx).toContain(`<a:endCxn id="${ids[1]}"`)
    expect(new Set(cNvPrIds(dst, 0)).size).toBe(cNvPrIds(dst, 0).length)

    // the connector alone: its endpoints name shapes that were not copied → dropped, not misdirected
    const alone = pasteElements(dst, 0, [items[2]!], ZERO)!
    const lone = alone.slide.elements.at(-1)!.anchor.originalXml
    expect(lone).not.toContain('stCxn')
    expect(lone).not.toContain('endCxn')
  })

  it('two rId remaps never chain (rId2→rId3 must not be re-mapped by rId3→rId5)', async () => {
    const src = await sourceDeck()
    const s = src.deck.slides[0]!
    const a = addPicture(src, s, {
      bytes: PNG,
      ext: 'png',
      offset: { x: 0, y: 0, cx: 10 * EMU, cy: 10 * EMU },
    })!
    const copy = copyElementData(src, s, a)
    // an item whose rId collides with an id the target already uses for something else
    const dst = await sourceDeck()
    const first = dst.deck.slides[0]!
    addPicture(dst, first, {
      bytes: Uint8Array.from([...PNG, 0]),
      ext: 'png',
      offset: { x: 0, y: 0, cx: 5 * EMU, cy: 5 * EMU },
    })
    const r = pasteElements(dst, 0, [copy], ZERO)!
    const pasted = r.slide.elements.at(-1)!
    expect(pasted.type).toBe('picture')
    const reopened = await openPptx(await savePptx(dst))
    const el = reopened.deck.slides[0]!.elements.at(-1) as { mediaRef?: string }
    expect(Buffer.from(reopened.archive.readBytes(el.mediaRef!)!)).toEqual(Buffer.from(PNG))
  })

  it('SmartArt data models are carried: drawing rel, data/layout/style/colors parts, relId follows', async () => {
    const src = await sourceDeck()
    const s = src.deck.slides[0]!
    const rels = src.archive.readText(`ppt/slides/_rels/${s.path.split('/').pop()}.rels`)!
    const add = (target: string, type: string, id: string, ct: string, body: string) => {
      src.archive.entries.set(`ppt/${target}`, Buffer.from(body, 'utf8'))
      const ct0 = src.archive.readText('[Content_Types].xml')!
      src.archive.entries.set(
        '[Content_Types].xml',
        Buffer.from(
          ct0.replace(
            '</Types>',
            `<Override PartName="/ppt/${target}" ContentType="${ct}"/></Types>`,
          ),
          'utf8',
        ),
      )
      return `<Relationship Id="${id}" Type="${type}" Target="../${target}"/>`
    }
    const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    const lines = [
      add(
        'diagrams/data1.xml',
        `${R}/diagramData`,
        'rId91',
        'application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml',
        '<dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"><dgm:extLst><a:ext xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><dsp:dataModelExt xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram" relId="rId95" minVer="http://schemas.openxmlformats.org/drawingml/2006/diagram"/></a:ext></dgm:extLst></dgm:dataModel>',
      ),
      add(
        'diagrams/layout1.xml',
        `${R}/diagramLayout`,
        'rId92',
        'application/vnd.openxmlformats-officedocument.drawingml.diagramLayout+xml',
        '<dgm:layoutDef xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"/>',
      ),
      add(
        'diagrams/quickStyle1.xml',
        `${R}/diagramQuickStyle`,
        'rId93',
        'application/vnd.openxmlformats-officedocument.drawingml.diagramStyle+xml',
        '<dgm:styleDef xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"/>',
      ),
      add(
        'diagrams/colors1.xml',
        `${R}/diagramColors`,
        'rId94',
        'application/vnd.openxmlformats-officedocument.drawingml.diagramColors+xml',
        '<dgm:colorsDef xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"/>',
      ),
      add(
        'diagrams/drawing1.xml',
        'http://schemas.microsoft.com/office/2007/relationships/diagramDrawing',
        'rId95',
        'application/vnd.ms-office.drawingml.diagramDrawing+xml',
        '<dsp:drawing xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram"/>',
      ),
    ]
    src.archive.entries.set(
      `ppt/slides/_rels/${s.path.split('/').pop()}.rels`,
      Buffer.from(rels.replace('</Relationships>', `${lines.join('')}</Relationships>`), 'utf8'),
    )
    const frame =
      '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="950" name="Diagram"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram"><dgm:relIds xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" r:dm="rId91" r:lo="rId92" r:qs="rId93" r:cs="rId94"/></a:graphicData></a:graphic></p:graphicFrame>'
    appendRawElements(src, 0, [frame])
    const el = s.elements.at(-1)!
    const item = copyElementData(src, s, el)
    expect(item.rels.map((r) => r.rid).sort()).toEqual([
      'rId91',
      'rId92',
      'rId93',
      'rId94',
      'rId95',
    ])
    expect(
      Object.keys(item.parts ?? {}).filter((p) => p.includes('diagrams/') && !p.includes('_rels'))
        .length,
    ).toBe(5)

    const dst = await sourceDeck()
    const r = pasteElements(dst, 0, [item], ZERO)!
    const xml = r.slide.elements.at(-1)!.anchor.originalXml
    const dstRels = dst.archive.readRels(dst.deck.slides[0]!.path)
    for (const attr of ['dm', 'lo', 'qs', 'cs']) {
      const rid = new RegExp(`r:${attr}="(rId\\d+)"`).exec(xml)![1]!
      expect(dstRels.get(rid), `${attr} relationship`).toBeDefined()
    }
    const dataRel = dstRels.get(/r:dm="(rId\d+)"/.exec(xml)![1]!)!
    const dataPath = `ppt/diagrams/${dataRel.target.split('/').pop()}`
    const dataXml = dst.archive.readText(dataPath)!
    const drawingRid = /relId="(rId\d+)"/.exec(dataXml)![1]!
    expect(dstRels.get(drawingRid)?.type).toBe(
      'http://schemas.microsoft.com/office/2007/relationships/diagramDrawing',
    )
  })

  it('a chart pasted twice into one deck owns two chart parts (not one shared part)', async () => {
    const src = await sourceDeck()
    const c = addChart(src, 0, {
      kind: 'bar',
      categories: ['a', 'b'],
      series: [{ name: 's', values: [1, 2] }],
      offset: { x: 0, y: 0, cx: 200 * EMU, cy: 100 * EMU },
    })!
    const el = src.deck.slides[0]!.elements.find((e) => e.id === c.elementId)!
    const item = copyElementData(src, src.deck.slides[0]!, el)
    const r1 = pasteElements(src, 0, [item], ZERO)!
    const r2 = pasteElements(src, 0, [item], ZERO)!
    expect(r1.elementIds.length + r2.elementIds.length).toBe(2)
    const charts = [...src.archive.entries.keys()].filter((p) =>
      /^ppt\/charts\/chart\d+\.xml$/.test(p),
    )
    expect(charts.length).toBe(3)
    const reopened = await openPptx(await savePptx(src))
    expect(reopened.deck.slides[0]!.elements.filter((e) => e.type === 'chart')).toHaveLength(3)
  })
})

describe('placeholder copy', () => {
  it('copying a title placeholder yields a free text box that keeps its frame and look', async () => {
    const src = await sourceDeck()
    const s = src.deck.slides[0]!
    const title = s.elements.find((e) =>
      /<p:ph[^>]*type="(title|ctrTitle)"/.test(e.anchor.originalXml),
    )!
    expect(title).toBeDefined()
    const item = copyElementData(src, s, title, { demotePlaceholder: true })
    expect(item.xml).not.toContain('<p:ph')
    expect(item.xml).toContain('<a:xfrm')
    const dst = await sourceDeck()
    const before = dst.deck.slides[0]!.elements.length
    const r = pasteElements(dst, 0, [item], ZERO)!
    expect(r.slide.elements.length).toBe(before + 1)
    const pasted = r.slide.elements.at(-1) as TextElement
    expect(pasted.placeholder).toBeUndefined()
    expect(pasted.transform.offset).toEqual(title.transform.offset)
    const t0 = (title as TextElement).text!.paragraphs[0]!.runs[0]!
    const t1 = pasted.text!.paragraphs[0]!.runs[0]!
    expect(t1.text).toBe(t0.text)
    expect(t1.fontSize).toBe(t0.fontSize)
  })
})

describe('PowerPoint native clipboard formats', () => {
  it('reads a PowerPoint GVML textbox into an editable element', async () => {
    const parsed = (await parseGvml(clip('textbox.gvml')))!
    expect(parsed.items).toHaveLength(1)
    expect(parsed.items[0]!.xml).toContain('<p:txBody>')
    expect(parsed.items[0]!.xml).not.toContain('<a:txSp')
    const dst = await sourceDeck()
    const r = pasteElements(dst, 0, parsed.items, ZERO)!
    const el = r.slide.elements.at(-1) as TextElement
    expect(el.type === 'text' || el.type === 'shape').toBe(true)
    expect(el.text!.paragraphs[0]!.runs[0]!.text).toBe('Bold red Georgia 24')
    expect(el.text!.paragraphs[0]!.runs[0]!.bold).toBe(true)
    expect(el.text!.paragraphs[1]!.bullet?.type).toBe('char')
    expect(parsed.bbox.cx).toBe(300 * EMU)
  })

  it('reads gradient shape, connector, picture (media rels) and chart packages', async () => {
    const dst = await sourceDeck()
    for (const name of ['gradient-shape.gvml', 'connector.gvml', 'picture.gvml', 'chart.gvml']) {
      const parsed = (await parseGvml(clip(name)))!
      expect(parsed, name).not.toBeNull()
      const r = pasteElements(dst, 0, parsed.items, ZERO)
      expect(r, name).not.toBeNull()
    }
    const els = dst.deck.slides[0]!.elements
    const grad = els.find((e) => e.anchor.originalXml.includes('Gradient Shape')) as TextElement
    expect(grad.fill?.type).toBe('gradient')
    expect(grad.shadow).toBeDefined()
    expect(els.some((e) => e.type === 'picture')).toBe(true)
    expect(els.some((e) => e.type === 'chart')).toBe(true)
    // round-trips through a save: PowerPoint-style media + chart land under ppt/
    const reopened = await openPptx(await savePptx(dst))
    expect(
      [...reopened.archive.entries.keys()].some((p) => /^ppt\/charts\/chart\d+\.xml$/.test(p)),
    ).toBe(true)
    expect(reopened.deck.slides[0]!.elements.some((e) => e.type === 'chart')).toBe(true)
    const ct = reopened.archive.readText('[Content_Types].xml')!
    expect(ct).toContain('drawingml.chart+xml')
    expect(new Set(cNvPrIds(reopened, 0)).size).toBe(cNvPrIds(reopened, 0).length)
  })

  it('reads a PowerPoint table package into a native table', async () => {
    const parsed = (await parseTableClip(clip('table.tblclip')))!
    expect(parsed.items[0]!.xml).toContain('<p:graphicFrame>')
    const dst = await sourceDeck()
    const r = pasteElements(dst, 0, parsed.items, ZERO)!
    const t = r.slide.elements.at(-1) as TableElement
    expect(t.type).toBe('table')
    expect(t.rows).toHaveLength(3)
    expect(t.colWidths).toHaveLength(3)
    expect(t.rows[0]![0]!.gridSpan).toBe(2)
  })

  it('writes GVML that is a well-formed OPC package PowerPoint can read (no p: vocabulary left)', async () => {
    const src = await sourceDeck()
    const s = src.deck.slides[0]!
    const tb = addElement(s, {
      kind: 'roundRect',
      offset: { x: 50 * EMU, y: 60 * EMU, cx: 100 * EMU, cy: 50 * EMU },
      fillColor: '#FFC000',
      paragraphs: [{ runs: [{ text: 'hello', bold: true }] }],
    })
    const pic = addPicture(src, s, {
      bytes: PNG,
      ext: 'png',
      offset: { x: 200 * EMU, y: 60 * EMU, cx: 40 * EMU, cy: 40 * EMU },
    })!
    const items = [tb, pic].map((e) => copyElementData(src, s, e))
    const built = await buildNativeClipboard(items, src.archive.readText('ppt/theme/theme1.xml'))
    expect(built.gvml).toBeDefined()
    const zip = await JSZip.loadAsync(built.gvml!)
    const names = Object.keys(zip.files)
    expect(names).toContain('[Content_Types].xml')
    expect(names).toContain('clipboard/drawings/drawing1.xml')
    expect(names).toContain('clipboard/theme/theme1.xml')
    expect(names.some((n) => n.startsWith('clipboard/media/'))).toBe(true)
    const drawing = await zip.file('clipboard/drawings/drawing1.xml')!.async('string')
    expect(drawing).toContain('<lc:lockedCanvas')
    expect(drawing).toContain('<a:txSp>')
    expect(drawing).not.toMatch(/<\/?p:/)
    // group frame = selection bounding box
    expect(drawing).toContain(`<a:chOff x="${50 * EMU}" y="${60 * EMU}"/>`)
    expect(drawing).toContain(`<a:chExt cx="${190 * EMU}" cy="${50 * EMU}"/>`)
    const rels = await zip.file('clipboard/drawings/_rels/drawing1.xml.rels')!.async('string')
    expect(rels).toContain('../media/')
    expect(rels).toContain('/theme')
    const ct = await zip.file('[Content_Types].xml')!.async('string')
    expect(ct).toContain('Extension="png"')
  })

  it('GenOffice → clipboard → GenOffice keeps text, fill, picture and size (round trip)', async () => {
    const src = await sourceDeck()
    const s = src.deck.slides[0]!
    const sh = addElement(s, {
      kind: 'ellipse',
      offset: { x: 50 * EMU, y: 60 * EMU, cx: 100 * EMU, cy: 50 * EMU },
      fillColor: '#00B050',
      paragraphs: [{ runs: [{ text: 'round trip', italic: true }] }],
    })
    const pic = addPicture(src, s, {
      bytes: PNG,
      ext: 'png',
      offset: { x: 200 * EMU, y: 60 * EMU, cx: 40 * EMU, cy: 40 * EMU },
    })!
    const items = [sh, pic].map((e) => copyElementData(src, s, e))
    const built = await buildNativeClipboard(items)
    const parsed = (await parseGvml(built.gvml!))!
    expect(parsed.items).toHaveLength(2)
    const dst = await sourceDeck()
    const r = pasteElements(dst, 0, parsed.items, ZERO)!
    const [a, b] = r.slide.elements.slice(-2)
    const shape = a as TextElement
    expect(shape.presetGeometry).toBe('ellipse')
    expect(shape.fill).toEqual({ type: 'solid', color: '#00B050' })
    expect(shape.text!.paragraphs[0]!.runs[0]!.text).toBe('round trip')
    expect(shape.text!.paragraphs[0]!.runs[0]!.italic).toBe(true)
    expect(shape.transform.offset).toEqual({
      x: 50 * EMU,
      y: 60 * EMU,
      cx: 100 * EMU,
      cy: 50 * EMU,
    })
    expect(b!.type).toBe('picture')
    expect(Buffer.from(dst.archive.readBytes((b as { mediaRef: string }).mediaRef)!)).toEqual(
      Buffer.from(PNG),
    )
  })

  it('a lone table is written in the table format and reads back with its style id', async () => {
    const src = await sourceDeck()
    ensureTableStylePart(src, STYLE, STYLE_DEF('accent2'))
    appendRawElements(src, 0, [tableXml(900, STYLE)])
    const t = src.deck.slides[0]!.elements.find((e) => e.type === 'table')!
    const built = await buildNativeClipboard([copyElementData(src, src.deck.slides[0]!, t)])
    expect(built.gvml).toBeUndefined()
    expect(built.table).toBeDefined()
    const zip = await JSZip.loadAsync(built.table!)
    const drawing = await zip.file('clipboard/drawings/drawing1.xml')!.async('string')
    expect(drawing.startsWith('<?xml')).toBe(true)
    expect(drawing).toContain('<a:tbl>')
    expect(drawing).not.toMatch(/<\/?p:/)
    const parsed = (await parseTableClip(built.table!))!
    expect(parsed.items[0]!.tableStyles?.[0]?.id).toBe(STYLE)
    expect(parsed.bbox.cx).toBe(200 * EMU)
  })
})

describe('text / HTML renditions and HTML table ingestion', () => {
  it('renders a copied table as TSV and an HTML table (Excel / Numbers paste a grid)', async () => {
    const src = await sourceDeck()
    appendRawElements(src, 0, [tableXml(900, '{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}')])
    const t = src.deck.slides[0]!.elements.find((e) => e.type === 'table') as TableElement
    expect(elementPlainText(t)).toBe('a\tb\nc\td')
    const html = tableToHtml(t)
    expect(html).toContain('<table')
    expect(html.match(/<tr>/g)).toHaveLength(2)
    expect(html).toContain('>a</td>')
  })

  it('an Excel-style HTML table becomes a native table with merges and colours', async () => {
    const html = `<html><body><table><tr><th colspan=2 style="background:#FFC000;font-weight:700">Title &amp; more</th></tr>
      <tr><td style='color:rgb(192,0,0)'>1,5</td><td style="text-align:right">x<br>y</td></tr>
      <tr><td rowspan=2>m</td><td>z</td></tr><tr><td>w</td></tr></table></body></html>`
    const item = htmlTableToItem(html)!
    const dst = await sourceDeck()
    const r = pasteElements(dst, 0, [item], ZERO)!
    const t = r.slide.elements.at(-1) as TableElement
    expect(t.type).toBe('table')
    expect(t.rows).toHaveLength(4)
    expect(t.colWidths).toHaveLength(2)
    expect(t.rows[0]![0]!.gridSpan).toBe(2)
    expect(t.rows[0]![0]!.fill).toEqual({ type: 'solid', color: '#FFC000' })
    expect(t.rows[0]![0]!.text!.paragraphs[0]!.runs[0]!.text).toBe('Title & more')
    expect(t.rows[1]![0]!.text!.paragraphs[0]!.runs[0]!.color).toBe('#C00000')
    expect(t.rows[1]![1]!.text!.paragraphs.map((p) => p.runs[0]?.text)).toEqual(['x', 'y'])
    expect(t.rows[2]![0]!.rowSpan).toBe(2)
    const reopened = await openPptx(await savePptx(dst))
    expect(reopened.deck.slides[0]!.elements.at(-1)!.type).toBe('table')
  })

  it('an unformatted HTML table takes the default table style (header row, banding)', async () => {
    const item = htmlTableToItem(
      '<table><tr><th>A</th><th>B</th></tr><tr><td>1</td><td>2</td></tr></table>',
    )!
    expect(item.xml).toContain('firstRow="1"')
    expect(item.tableStyles).toEqual([{ id: '{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}' }])
    expect(htmlTableToItem('<p>no table</p>')).toBeNull()
  })

  it('bakeForExternal turns a custom table style into explicit cell formatting', async () => {
    const src = await sourceDeck()
    ensureTableStylePart(src, STYLE, STYLE_DEF('accent2'))
    appendRawElements(src, 0, [tableXml(900, STYLE)])
    const t = src.deck.slides[0]!.elements.find((e) => e.type === 'table') as TableElement
    const plain = copyElementData(src, src.deck.slides[0]!, t)
    const baked = copyElementData(src, src.deck.slides[0]!, t, { bakeForExternal: true })
    expect(plain.xml).toContain(STYLE)
    expect(baked.xml).not.toContain('tableStyleId')
    expect(baked.xml).toMatch(/<a:tcPr[^>]*>[^]*<a:solidFill><a:srgbClr val="[0-9A-F]{6}"\/>/)
    // a built-in style stays a style reference: PowerPoint knows its GUID
    const dst = await sourceDeck()
    appendRawElements(dst, 0, [tableXml(901, '{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}')])
    const bt = dst.deck.slides[0]!.elements.find((e) => e.type === 'table')!
    expect(copyElementData(dst, dst.deck.slides[0]!, bt, { bakeForExternal: true }).xml).toContain(
      '5C22544A',
    )
  })
})

describe('slide copy across decks', () => {
  it('carries the custom table style and the speaker notes', async () => {
    const src = await sourceDeck()
    ensureTableStylePart(src, STYLE, STYLE_DEF('accent2'))
    appendRawElements(src, 0, [tableXml(900, STYLE)])
    expect(setSlideNotes(src, 0, 'remember to smile')).toBe(true)
    const bundle = copySlide(src, 0)!
    expect(bundle.tableStyles?.[0]?.id).toBe(STYLE)
    expect(bundle.notes?.length).toBeGreaterThan(0)

    const dst = await sourceDeck()
    ensureTableStylePart(
      dst,
      '{AAAAAAAA-0000-0000-0000-000000000000}',
      STYLE_DEF('accent1').replace(STYLE, '{AAAAAAAA-0000-0000-0000-000000000000}'),
    )
    const slide = pasteSlide(dst, 0, bundle)!
    const t = slide.elements.find((e) => e.type === 'table') as TableElement
    expect(t.styleId).toBe(STYLE)
    expect(dst.archive.readText('ppt/tableStyles.xml')).toContain(`styleId="${STYLE}"`)
    expect(t.rows[0]![0]!.fill).toBeDefined()
    const idx = dst.deck.slides.indexOf(slide)
    expect(getSlideNotes(dst.archive, dst.deck.slides[idx]!.path)).toContain('remember to smile')
    const reopened = await openPptx(await savePptx(dst))
    expect(getSlideNotes(reopened.archive, reopened.deck.slides[idx]!.path)).toContain(
      'remember to smile',
    )
  })
})

describe('formatted text from HTML', () => {
  it('keeps emphasis, colour, size and lists; plain text stays on the plain path', async () => {
    const html =
      '<html><body><!--StartFragment--><p style="font-size:18pt">Hello <b>bold</b> and <span style="color:#C00000;font-size:24pt">red big</span></p>' +
      '<ul><li>one</li><li><i>two</i></li></ul><!--EndFragment--></body></html>'
    const item = htmlTextToItem(html, { widthEmu: 5_000_000 })!
    expect(item).not.toBeNull()
    const dst = await sourceDeck()
    const r = pasteElements(dst, 0, [item], ZERO)!
    const el = r.slide.elements.at(-1) as TextElement
    const paras = el.text!.paragraphs
    expect(paras).toHaveLength(3)
    expect(paras[0]!.runs.map((x) => x.text).join('')).toBe('Hello bold and red big')
    expect(paras[0]!.runs.find((x) => x.text === 'bold')!.bold).toBe(true)
    const red = paras[0]!.runs.find((x) => x.text === 'red big')!
    expect(red.fontSize).toBe(24)
    expect(red.color).toBe('#C00000')
    expect(paras[1]!.bullet?.type).toBe('char')
    expect(paras[2]!.runs[0]!.italic).toBe(true)
    expect(htmlTextToItem('<p>just words</p>', { widthEmu: 1 })).toBeNull()
    expect(htmlTextToItem('<table><tr><td>x</td></tr></table>', { widthEmu: 1 })).toBeNull()
  })
})

describe('animations travel with a copied shape', () => {
  it('the effect is re-targeted at the pasted shape (same slide and another deck)', async () => {
    const src = await sourceDeck()
    const s = src.deck.slides[0]!
    const box = addElement(s, {
      kind: 'rect',
      offset: { x: 10 * EMU, y: 10 * EMU, cx: 50 * EMU, cy: 50 * EMU },
      fillColor: '#00B0F0',
    })
    const other = addElement(s, {
      kind: 'rect',
      offset: { x: 100 * EMU, y: 10 * EMU, cx: 50 * EMU, cy: 50 * EMU },
      fillColor: '#92D050',
    })
    // materialise so ids exist
    appendRawElements(src, 0, [])
    const live = src.deck.slides[0]!
    const mine = live.elements.find(
      (e) =>
        e.anchor.originalXml.includes(`name="Shape`) &&
        e.type === 'shape' &&
        e.transform.offset.x === 10 * EMU,
    )!
    const spid = elementSpid(mine)!
    const otherSpid = elementSpid(live.elements.find((e) => e.transform.offset.x === 100 * EMU)!)!
    setSlideAnimations(live, [
      { spid, effect: 'fade', trigger: 'onClick', durationMs: 500, delayMs: 0 },
      { spid: otherSpid, effect: 'appear', trigger: 'onClick', durationMs: 500, delayMs: 0 },
    ])
    void box
    void other
    const item = copyElementData(src, live, mine)
    expect(item.animations?.map((a) => a.effect)).toEqual(['fade'])

    const dst = await sourceDeck()
    const r = pasteElements(dst, 0, [item], ZERO)!
    const pasted = r.slide.elements.at(-1)!
    const anims = getSlideAnimations(dst.deck.slides[0]!)
    expect(anims.map((a) => a.effect)).toEqual(['fade'])
    expect(anims[0]!.spid).toBe(elementSpid(pasted))
    const reopened = await openPptx(await savePptx(dst))
    expect(getSlideAnimations(reopened.deck.slides[0]!).length).toBe(1)
  })
})

describe('PowerPoint slide clipboard package', () => {
  it('reads the slides PowerPoint copies and pastes them into another deck', async () => {
    const pkg = (await openSlidesClipboardPackage(clip('slides.pkg')))!
    expect(pkg).not.toBeNull()
    expect(pkg.deck.slides).toHaveLength(1)
    expect(pkg.deck.slides[0]!.elements.length).toBeGreaterThan(5)
    const bundle = copySlide(pkg, 0)!
    const dst = await sourceDeck()
    const before = dst.deck.slides.length
    const slide = pasteSlide(dst, 0, bundle)!
    expect(dst.deck.slides.length).toBe(before + 1)
    expect(slide.elements.some((e) => e.type === 'chart')).toBe(true)
    const reopened = await openPptx(await savePptx(dst))
    expect(reopened.deck.slides[1]!.elements.length).toBe(slide.elements.length)
  })

  it('writes a slides package PowerPoint can read back (round trip, all parts under clipboard/)', async () => {
    const src = await sourceDeck()
    appendRawElements(src, 0, [tableXml(900, '{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}')])
    const bytes = (await buildSlidesClipboardPackage(src, [0, 1]))!
    const zip = await JSZip.loadAsync(bytes)
    const names = Object.keys(zip.files)
    expect(names).toContain('clipboard/presentation.xml')
    expect(names).toContain('clipboard/slides/slide1.xml')
    expect(names).toContain('clipboard/slides/slide2.xml')
    expect(names.some((n) => n.startsWith('clipboard/slideLayouts/'))).toBe(true)
    expect(names.some((n) => n.startsWith('ppt/'))).toBe(false)
    const back = (await openSlidesClipboardPackage(bytes))!
    expect(back.deck.slides).toHaveLength(2)
    expect(back.deck.slides[0]!.elements.length).toBe(src.deck.slides[0]!.elements.length)
    expect(back.deck.size).toEqual(src.deck.size)
  })
})
