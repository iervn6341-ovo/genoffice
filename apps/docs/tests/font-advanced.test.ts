import JSZip from 'jszip'
import { Editor } from '@tiptap/core'
import { TextSelection } from '@tiptap/pm/state'
import { describe, expect, it } from 'vitest'
import { parseDocx, patchRunPropertiesXml, saveDocx } from '@genoffice/docx-engine'
import {
  buildDocx,
  IMAGE_PARAGRAPH_XML,
  MATH_PARAGRAPH_XML,
} from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { sameBodyAfterManagedRPrRebase } from '../src/renderer/editor/save-rebase'
import {
  applyFontChanges,
  commonFontValue,
  enabledStylisticSets,
  fontSelectionSamples,
} from '../src/renderer/editor/font-advanced'

const importedRPr =
  '<w:rFonts w:ascii="Cambria" w:eastAsia="SimSun" w:cs="Arial"/>' +
  '<w:u w:val="double" w:color="C0504D"/>' +
  '<w:spacing w:val="30"/><w:w w:val="125"/><w:kern w:val="28"/>' +
  '<w:position w:val="4"/><w:smallCaps/><w:lang w:val="en-US"/>'

async function open(extraRPr = '', rootNamespace?: string) {
  const bytes = await buildDocx({
    ...(rootNamespace || extraRPr.includes('<w14:')
      ? {
          docRootExtraAttrs:
            rootNamespace ?? 'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"',
        }
      : {}),
    bodyXml:
      `<w:p><w:r><w:rPr>${importedRPr}${extraRPr}</w:rPr><w:t>Alpha</w:t></w:r>` +
      '<w:r><w:t xml:space="preserve"> Beta</w:t></w:r>' +
      '<w:fldSimple w:instr=" DATE "><w:r><w:t>October</w:t></w:r></w:fldSimple>' +
      '</w:p>',
  })
  const parsed = await parseDocx(bytes)
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
  const ranges: Record<string, { from: number; to: number }> = {}
  editor.state.doc.descendants((node, pos) => {
    if (node.isText && node.text) ranges[node.text.trim()] = { from: pos, to: pos + node.nodeSize }
  })
  const savedBytes = async () => {
    const plan = pmDocToSavePlan(editor.state.doc.toJSON() as PmNode, parsed.blocks)
    return saveDocx(parsed, plan.saveBlocks)
  }
  const xml = async () =>
    (await JSZip.loadAsync(await savedBytes())).file('word/document.xml')!.async('string')
  return { editor, ranges, xml, savedBytes, parsed }
}

describe('Font Advanced authoring', () => {
  it('patches only requested raw run-property children, retaining unrelated imported XML', () => {
    const original = `<w:rPr>${importedRPr}<w:shd w:fill="EEECE1"/></w:rPr>`
    const patched = patchRunPropertiesXml(original, {
      'w:spacing': '<w:spacing w:val="-20"/>',
      'w:kern': '<w:kern w:val="32"/>',
    })
    expect(patched).toContain('<w:rFonts w:ascii="Cambria" w:eastAsia="SimSun" w:cs="Arial"/>')
    expect(patched).toContain('<w:u w:val="double" w:color="C0504D"/>')
    expect(patched).toContain('<w:shd w:fill="EEECE1"/>')
    expect(patched).toContain('<w:spacing w:val="-20"/>')
    expect(patched).toContain('<w:kern w:val="32"/>')
    expect(patched).not.toContain('<w:spacing w:val="30"/>')
  })

  it('backfills mixed selection without applying untouched values', async () => {
    const { editor, ranges, xml } = await open()
    const samples = fontSelectionSamples(editor, ranges.Alpha.from, ranges.Beta.to)
    expect(commonFontValue(samples, (sample) => sample.attrs.charSpacingTwips).mixed).toBe(true)
    expect(commonFontValue(samples, (sample) => sample.attrs.fontAscii).mixed).toBe(true)
    editor.view.dispatch(
      editor.state.tr.setSelection(
        TextSelection.create(editor.state.doc, ranges.Alpha.from, ranges.Beta.to),
      ),
    )
    applyFontChanges(editor, ranges.Alpha.from, ranges.Beta.to, { sizeHalfPoints: 28 })
    const savedXml = await xml()
    expect(savedXml).toContain('<w:spacing w:val="30"/>')
    expect(savedXml).toContain('<w:w w:val="125"/>')
    expect(savedXml).toContain('<w:kern w:val="28"/>')
    expect(savedXml).toContain('<w:u w:val="double" w:color="C0504D"/>')
    expect(savedXml).toContain('<w:instrText xml:space="preserve"> DATE </w:instrText>')
    expect(savedXml).toContain('<w:t xml:space="preserve">October</w:t>')
    editor.destroy()
  })

  it('saves edited underline and advanced values, then restores them through Undo/Redo', async () => {
    const { editor, ranges, xml, savedBytes } = await open()
    editor.view.dispatch(
      editor.state.tr.setSelection(
        TextSelection.create(editor.state.doc, ranges.Alpha.from, ranges.Alpha.to),
      ),
    )
    applyFontChanges(editor, ranges.Alpha.from, ranges.Alpha.to, {
      underlineStyle: 'wave',
      underlineColor: '00AA55',
      charScalePct: 90,
      charSpacingTwips: -20,
      positionHalfPoints: -4,
      kernHalfPoints: 32,
      dstrike: true,
    })
    const after = await xml()
    const reparsed = await parseDocx(await savedBytes())
    const rebasedJson = blocksToPmDoc(reparsed.blocks)
    expect(editor.state.doc.content.eq(editor.schema.nodeFromJSON(rebasedJson).content)).toBe(true)
    for (const fragment of [
      '<w:u w:val="wave" w:color="00AA55"/>',
      '<w:w w:val="90"/>',
      '<w:spacing w:val="-20"/>',
      '<w:position w:val="-4"/>',
      '<w:kern w:val="32"/>',
      '<w:dstrike w:val="1"/>',
      '<w:lang w:val="en-US"/>',
      '<w:rFonts w:ascii="Cambria" w:eastAsia="SimSun" w:cs="Arial"/>',
    ])
      expect(after).toContain(fragment)
    expect(after).toContain('<w:instrText xml:space="preserve"> DATE </w:instrText>')
    editor.commands.undo()
    expect(await xml()).toContain('<w:u w:val="double" w:color="C0504D"/>')
    editor.commands.redo()
    expect(await xml()).toContain('<w:u w:val="wave" w:color="00AA55"/>')
    editor.destroy()
  })

  it('authors Word 2010 typography values with a bound w14 namespace and retains the imported field', async () => {
    const { editor, ranges, xml, savedBytes } = await open()
    applyFontChanges(editor, ranges.Alpha.from, ranges.Alpha.to, {
      ligatures: 'all',
      numberSpacing: 'tabular',
      numberForm: 'oldStyle',
      contextualAlternates: true,
      snapToGrid: false,
    })
    const savedXml = await xml()
    expect(savedXml).toContain('xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"')
    expect(savedXml).toContain('mc:Ignorable="w14"')
    expect(savedXml).toContain('<w14:ligatures w14:val="all"/>')
    expect(savedXml).toContain('<w14:numSpacing w14:val="tabular"/>')
    expect(savedXml).toContain('<w14:numForm w14:val="oldStyle"/>')
    expect(savedXml).toContain('<w14:cntxtAlts/>')
    expect(savedXml).toContain('<w:snapToGrid w:val="0"/>')
    expect(savedXml).toContain('<w:instrText xml:space="preserve"> DATE </w:instrText>')
    const reparsed = await parseDocx(await savedBytes())
    expect(
      editor.state.doc.content.eq(
        editor.schema.nodeFromJSON(blocksToPmDoc(reparsed.blocks)).content,
      ),
    ).toBe(true)
    editor.destroy()
  })

  it('preserves imported multiple stylistic sets until touched, then supports set, Undo/Redo, and Default', async () => {
    const original =
      '<w14:stylisticSets><w14:styleSet w14:id="1"/><w14:styleSet w14:id="4"/></w14:stylisticSets>'
    const { editor, ranges, xml } = await open(original)
    expect(enabledStylisticSets(original)).toEqual([1, 4])
    expect(editor.view.dom.innerHTML).toContain('&quot;ss01&quot; 1,&quot;ss04&quot; 1')
    expect(
      enabledStylisticSets('<w14:stylisticSets><w14:styleSet w14:id="22"/></w14:stylisticSets>'),
    ).toBeNull()
    applyFontChanges(editor, ranges.Alpha.from, ranges.Alpha.to, { sizeHalfPoints: 30 })
    expect(alphaRPr(await xml())).toContain(original)
    applyFontChanges(editor, ranges.Alpha.from, ranges.Alpha.to, { stylisticSet: 7 })
    expect(editor.view.dom.innerHTML).toContain('&quot;ss07&quot; 1')
    expect(alphaRPr(await xml())).toContain(
      '<w14:stylisticSets><w14:styleSet w14:id="7"/></w14:stylisticSets>',
    )
    expect(alphaRPr(await xml())).not.toContain(original)
    editor.commands.undo()
    expect(alphaRPr(await xml())).toContain(original)
    editor.commands.redo()
    expect(alphaRPr(await xml())).toContain('<w14:styleSet w14:id="7"/>')
    applyFontChanges(editor, ranges.Alpha.from, ranges.Alpha.to, { stylisticSet: 0 })
    expect(alphaRPr(await xml())).not.toContain('<w14:stylisticSets>')
    expect(alphaRPr(await xml())).toContain('<w:lang w:val="en-US"/>')
    editor.destroy()
  })

  it('recognizes a Word 2010 namespace alias and replaces or clears only the aliased child', async () => {
    const aliased = '<x:stylisticSets><x:styleSet x:id="4"/></x:stylisticSets>'
    const { editor, ranges, xml, savedBytes } = await open(
      aliased,
      'xmlns:w14="urn:custom:reserved" xmlns:x="http://schemas.microsoft.com/office/word/2010/wordml"',
    )
    const sample = fontSelectionSamples(editor, ranges.Alpha.from, ranges.Alpha.to)[0]
    expect(sample.attrs.rawRPrW14Prefixes).toContain('x')
    expect(enabledStylisticSets(String(sample.attrs.rawRPr), ['x'])).toEqual([4])
    expect(editor.view.dom.innerHTML).toContain('&quot;ss04&quot; 1')
    applyFontChanges(editor, ranges.Alpha.from, ranges.Alpha.to, { sizeHalfPoints: 30 })
    expect(alphaRPr(await xml())).toContain(aliased)
    applyFontChanges(editor, ranges.Alpha.from, ranges.Alpha.to, { stylisticSet: 7 })
    const changed = alphaRPr(await xml())
    expect(changed).toMatch(/<x:stylisticSets\b[^>]*><x:styleSet x:id="7"\/><\/x:stylisticSets>/)
    expect(changed).not.toContain('<w14:stylisticSets>')
    expect(editor.view.dom.innerHTML).toContain('&quot;ss07&quot; 1')
    const reopened = await parseDocx(await savedBytes())
    const reopenedPm = editor.schema.nodeFromJSON(blocksToPmDoc(reopened.blocks))
    expect(editor.state.doc.content.eq(reopenedPm.content)).toBe(false)
    expect(sameBodyAfterManagedRPrRebase(editor.state.doc, reopenedPm)).toBe(true)
    const changedOpaque = reopenedPm.toJSON()
    const style = changedOpaque.content?.[0]?.content?.[0]?.marks?.find(
      (mark: { type: string }) => mark.type === 'docTextStyle',
    )
    if (!style?.attrs) throw new Error('Reopened run has no style mark')
    style.attrs.rawRPr = String(style.attrs.rawRPr).replace('x:id="7"', 'x:id="8"')
    expect(
      sameBodyAfterManagedRPrRebase(editor.state.doc, editor.schema.nodeFromJSON(changedOpaque)),
    ).toBe(false)
    const reopenedRun = reopened.blocks.flatMap((block) =>
      block.type === 'paragraph' ? block.runs : [],
    )[0]
    if (!reopenedRun) throw new Error('Reopened document has no text run')
    expect(reopenedRun.rawRPr).toContain('<x:styleSet x:id="7"/>')
    expect(reopenedRun.rawRPrW14Prefixes).toContain('x')
    editor.commands.undo()
    expect(alphaRPr(await xml())).toContain(aliased)
    editor.commands.redo()
    expect(alphaRPr(await xml())).toContain('<x:styleSet x:id="7"/>')
    applyFontChanges(editor, ranges.Alpha.from, ranges.Alpha.to, { stylisticSet: 0 })
    expect(alphaRPr(await xml())).not.toContain('stylisticSets')
    expect(alphaRPr(await xml())).toContain('<w:lang w:val="en-US"/>')
    editor.destroy()
  })

  it('keeps wrapper-local Word 2010 aliases bound and ignorable when replacing a set', async () => {
    const w14 = 'http://schemas.microsoft.com/office/word/2010/wordml'
    const local = `<x:stylisticSets xmlns:x="${w14}"><x:styleSet x:id="4"/></x:stylisticSets>`
    const { editor, ranges, xml, savedBytes } = await open(local, 'xmlns:w14="urn:custom:reserved"')
    applyFontChanges(editor, ranges.Alpha.from, ranges.Alpha.to, {
      stylisticSet: 7,
      ligatures: 'all',
    })
    const saved = await xml()
    expect(saved).toContain('<x:styleSet x:id="7"/>')
    expect(saved).toMatch(/<x:ligatures\b[^>]*x:val="all"/)
    expect(saved).toContain(`xmlns:x="${w14}"`)
    expect(saved).toMatch(/<x:stylisticSets\b[^>]*mc:Ignorable="x"/)
    const rPr = alphaRPr(saved)
    expect(rPr.indexOf('<w:rFonts')).toBeLessThan(rPr.indexOf('<w:lang'))
    expect(rPr.indexOf('<w:lang')).toBeLessThan(rPr.indexOf('<x:ligatures'))
    expect(rPr.indexOf('<x:ligatures')).toBeLessThan(rPr.indexOf('<x:stylisticSets'))
    const parsedXml = new DOMParser().parseFromString(saved, 'application/xml')
    expect(parsedXml.querySelector('parsererror')).toBeNull()
    const reopened = await parseDocx(await savedBytes())
    const run = reopened.blocks.flatMap((block) =>
      block.type === 'paragraph' ? block.runs : [],
    )[0]
    expect(run?.rawRPr).toContain('<x:styleSet x:id="7"/>')
    expect(enabledStylisticSets(run?.rawRPr)).toEqual([7])
    editor.commands.undo()
    expect(alphaRPr(await xml())).toContain(local)
    editor.destroy()
  })

  it('materializes explicitly changed theme-backed font and color without losing save Undo', async () => {
    const themedRPr =
      '<w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi"' +
      ' w:eastAsiaTheme="majorEastAsia" w:ascii="Legacy Latin" w:hAnsi="Legacy Latin"' +
      ' w:eastAsia="Legacy East"/>' +
      '<w:color w:val="365F91" w:themeColor="accent1" w:themeShade="BF"/>' +
      '<w:lang w:val="en-US"/>'
    const themeXml =
      '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Test">' +
      '<a:themeElements><a:clrScheme name="Test"><a:accent1><a:srgbClr val="4472C4"/>' +
      '</a:accent1></a:clrScheme><a:fontScheme name="Test">' +
      '<a:majorFont><a:latin typeface="Heading Latin"/><a:ea typeface="Heading East"/>' +
      '<a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="Body Latin"/>' +
      '<a:ea typeface="Body East"/><a:cs typeface=""/></a:minorFont>' +
      '</a:fontScheme></a:themeElements></a:theme>'
    const bytes = await buildDocx({
      bodyXml:
        `<w:p><w:r><w:rPr>${themedRPr}</w:rPr><w:t>Alpha</w:t></w:r>` +
        '<w:fldSimple w:instr=" DATE "><w:r><w:t>October</w:t></w:r></w:fldSimple></w:p>',
      extraParts: [
        {
          path: 'word/theme/theme1.xml',
          xml: themeXml,
          contentType: 'application/vnd.openxmlformats-officedocument.theme+xml',
        },
      ],
    })
    const parsed = await parseDocx(bytes)
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })
    const before = fontSelectionSamples(editor, 1, 6)[0]
    expect(before.attrs.themeRFonts).toBeTruthy()
    expect(before.attrs.themeColor).toBeTruthy()
    applyFontChanges(editor, 1, 6, { fontAscii: 'Arial', color: '4472C4' })
    const edited = fontSelectionSamples(editor, 1, 6)[0]
    expect(edited.attrs.themeColor).toBeNull()
    expect(JSON.parse(String(edited.attrs.themeRFonts))).toEqual({ font: 'Heading East' })
    const plan = pmDocToSavePlan(editor.state.doc.toJSON() as PmNode, parsed.blocks)
    const saved = await saveDocx(parsed, plan.saveBlocks)
    const savedXml = await (await JSZip.loadAsync(saved)).file('word/document.xml')!.async('string')
    expect(alphaRPr(savedXml)).toContain('w:ascii="Arial"')
    expect(alphaRPr(savedXml)).not.toContain('asciiTheme')
    expect(alphaRPr(savedXml)).not.toContain('hAnsiTheme')
    expect(alphaRPr(savedXml)).toContain('w:eastAsiaTheme="majorEastAsia"')
    expect(alphaRPr(savedXml)).toContain('<w:color w:val="4472C4"/>')
    expect(alphaRPr(savedXml)).not.toContain('themeColor')
    expect(savedXml).toContain('<w:instrText xml:space="preserve"> DATE </w:instrText>')
    const reopened = await parseDocx(saved)
    const rebased = editor.schema.nodeFromJSON(blocksToPmDoc(reopened.blocks))
    expect(sameBodyAfterManagedRPrRebase(editor.state.doc, rebased)).toBe(true)
    editor.commands.undo()
    const undoPlan = pmDocToSavePlan(editor.state.doc.toJSON() as PmNode, reopened.blocks)
    const undoSaved = await saveDocx(reopened, undoPlan.saveBlocks)
    const undoXml = await (
      await JSZip.loadAsync(undoSaved)
    )
      .file('word/document.xml')!
      .async('string')
    expect(alphaRPr(undoXml)).toContain('w:asciiTheme="minorHAnsi"')
    expect(alphaRPr(undoXml)).toContain('w:themeColor="accent1"')
    expect(alphaRPr(undoXml)).toContain('w:eastAsiaTheme="majorEastAsia"')
    editor.destroy()
  })

  it('changes a Latin theme font without inventing an East Asian slot', async () => {
    const bytes = await buildDocx({
      bodyXml:
        '<w:p><w:r><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi"' +
        ' w:ascii="Legacy" w:hAnsi="Legacy"/></w:rPr><w:t>Latin</w:t></w:r></w:p>',
    })
    const parsed = await parseDocx(bytes)
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })
    applyFontChanges(editor, 1, 6, { fontAscii: 'Arial' })
    const plan = pmDocToSavePlan(editor.state.doc.toJSON() as PmNode, parsed.blocks)
    const saved = await saveDocx(parsed, plan.saveBlocks)
    const xml = await (await JSZip.loadAsync(saved)).file('word/document.xml')!.async('string')
    expect(xml).toContain('<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>')
    expect(xml).not.toContain('asciiTheme')
    expect(xml).not.toContain('eastAsia=')
    const reopened = await parseDocx(saved)
    const rebased = editor.schema.nodeFromJSON(blocksToPmDoc(reopened.blocks))
    expect(sameBodyAfterManagedRPrRebase(editor.state.doc, rebased)).toBe(true)
    editor.commands.undo()
    const undoPlan = pmDocToSavePlan(editor.state.doc.toJSON() as PmNode, reopened.blocks)
    const undoSaved = await saveDocx(reopened, undoPlan.saveBlocks)
    const undoXml = await (
      await JSZip.loadAsync(undoSaved)
    )
      .file('word/document.xml')!
      .async('string')
    expect(undoXml).toContain('w:asciiTheme="minorHAnsi"')
    editor.destroy()
  })

  it('does not apply a selection-wide font change to a protected field result', async () => {
    const { editor, ranges, xml } = await open()
    applyFontChanges(editor, ranges.Alpha.from, ranges.October.to, { color: 'AABBCC' })
    const savedXml = await xml()
    const fieldResult = savedXml.match(
      /<w:instrText[^>]*> DATE <\/w:instrText>[\s\S]*?<w:t[^>]*>October<\/w:t>/,
    )?.[0]
    expect(fieldResult).toBeTruthy()
    expect(fieldResult).not.toContain('AABBCC')
    expect(alphaRPr(savedXml)).toContain('<w:color w:val="AABBCC"/>')
    editor.destroy()
  })

  it('keeps one font-dialog Undo event across a save with field, picture, and equation siblings', async () => {
    const bytes = await buildDocx({
      docRootExtraAttrs: 'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"',
      withImage: true,
      bodyXml:
        `<w:p><w:r><w:rPr>${importedRPr}<w14:ligatures w14:val="all"/>` +
        '<w14:stylisticSets><w14:styleSet w14:id="4"/></w14:stylisticSets>' +
        '</w:rPr><w:t>Alpha</w:t></w:r></w:p>' +
        '<w:p><w:fldSimple w:instr=" DATE "><w:r><w:t>October</w:t></w:r></w:fldSimple></w:p>' +
        IMAGE_PARAGRAPH_XML +
        MATH_PARAGRAPH_XML,
    })
    const parsed = await parseDocx(bytes)
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })
    let range: { from: number; to: number } | undefined
    editor.state.doc.descendants((node, pos) => {
      if (node.isText && node.text === 'Alpha') range = { from: pos, to: pos + node.nodeSize }
    })
    if (!range) throw new Error('Fixture text not found')
    applyFontChanges(editor, range.from, range.to, {
      charScalePct: 90,
      positionHalfPoints: -4,
      kernHalfPoints: 0,
      underlineStyle: 'wave',
      stylisticSet: 7,
    })
    const plan = pmDocToSavePlan(editor.state.doc.toJSON() as PmNode, parsed.blocks)
    const reopened = await parseDocx(await saveDocx(parsed, plan.saveBlocks))
    const rebased = editor.schema.nodeFromJSON(blocksToPmDoc(reopened.blocks))
    expect(sameBodyAfterManagedRPrRebase(editor.state.doc, rebased)).toBe(true)
    editor.commands.undo()
    const undoPlan = pmDocToSavePlan(editor.state.doc.toJSON() as PmNode, reopened.blocks)
    const undoXml = (await JSZip.loadAsync(await saveDocx(reopened, undoPlan.saveBlocks))).file(
      'word/document.xml',
    )!
    const xml = await undoXml.async('string')
    expect(alphaRPr(xml)).toContain('<w:w w:val="125"/>')
    expect(alphaRPr(xml)).toContain('<w14:styleSet w14:id="4"/>')
    expect(xml).toContain('<w:drawing>')
    expect(xml).toContain('<m:oMath>')
    editor.destroy()
  })
})

function alphaRPr(xml: string): string {
  const run = xml.match(/<w:r>(?:(?!<\/w:r>).)*<w:t[^>]*>Alpha<\/w:t><\/w:r>/s)?.[0]
  return run?.match(/<w:rPr>.*?<\/w:rPr>/s)?.[0] ?? ''
}
