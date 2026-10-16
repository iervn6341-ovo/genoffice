import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { mergePPrFormat, parseDocx, saveDocx } from '../src/index'
import { buildDocx } from './helpers/build-docx'

const raw =
  '<w:pPr><w:keepNext w:val="0"/><w:keepLines/><w:pageBreakBefore w:val="0"/>' +
  '<w:widowControl w:val="0"/><w:suppressLineNumbers/>' +
  '<w:tabs><w:tab w:val="right" w:pos="1440" w:leader="heavy"/></w:tabs>' +
  '<w:suppressAutoHyphens/><w:kinsoku/><w:wordWrap w:val="0"/>' +
  '<w:overflowPunct/><w:topLinePunct/><w:autoSpaceDE w:val="0"/>' +
  '<w:autoSpaceDN/><w:spacing w:after="240"/><w:contextualSpacing/>' +
  '<w:textAlignment w:val="baseline"/><w:rPr><w:lang w:val="ja-JP"/></w:rPr></w:pPr>'

describe('advanced paragraph properties', () => {
  it('parses native pagination, tab, and Asian typography controls without rewriting untouched XML', async () => {
    const doc = await parseDocx(
      await buildDocx({ bodyXml: `<w:p>${raw}<w:r><w:t>本文</w:t></w:r></w:p>` }),
    )
    const format = doc.blocks[0].format!
    expect(format).toMatchObject({
      keepNext: false,
      keepLines: true,
      pageBreakBefore: false,
      widowControl: false,
      suppressLineNumbers: true,
      suppressAutoHyphens: true,
      kinsoku: true,
      wordWrap: false,
      overflowPunct: true,
      topLinePunct: true,
      autoSpaceDE: false,
      autoSpaceDN: true,
      textAlignment: 'baseline',
      tabStops: [{ pos: 1440, val: 'right', leader: 'heavy' }],
    })
    expect(mergePPrFormat(raw, format, format)).toBe(raw)
  })

  it('changes only touched native children and keeps opaque paragraph formatting', async () => {
    const doc = await parseDocx(
      await buildDocx({ bodyXml: `<w:p>${raw}<w:r><w:t>本文</w:t></w:r></w:p>` }),
    )
    const format = doc.blocks[0].format!
    const changed = mergePPrFormat(
      raw,
      {
        ...format,
        keepNext: true,
        widowControl: true,
        suppressAutoHyphens: false,
        kinsoku: false,
        topLinePunct: false,
        autoSpaceDN: false,
        textAlignment: 'center',
        tabStops: [{ pos: 1800, val: 'decimal', leader: 'dot' }],
      },
      format,
    )
    expect(changed).toContain('<w:keepNext/>')
    expect(changed).toContain('<w:widowControl/>')
    expect(changed).toContain('<w:suppressAutoHyphens w:val="0"/>')
    expect(changed).toContain('<w:kinsoku w:val="0"/>')
    expect(changed).toContain('<w:topLinePunct w:val="0"/>')
    expect(changed).toContain('<w:autoSpaceDN w:val="0"/>')
    expect(changed).toContain('<w:tab w:val="decimal" w:pos="1800" w:leader="dot"/>')
    expect(changed).toContain('<w:textAlignment w:val="center"/>')
    for (const unchanged of [
      '<w:keepLines/>',
      '<w:pageBreakBefore w:val="0"/>',
      '<w:spacing w:after="240"/>',
      '<w:contextualSpacing/>',
      '<w:rPr><w:lang w:val="ja-JP"/></w:rPr>',
    ])
      expect(changed).toContain(unchanged)
    expect(mergePPrFormat(raw, { ...format, tabStops: [] }, format)).not.toContain('<w:tabs>')
  })

  it('removes a newly authored flag when Undo restores an absent direct value after Save', async () => {
    const afterSave =
      '<w:pPr><w:suppressAutoHyphens/><w:kinsoku w:val="0"/>' +
      '<w:topLinePunct/><w:textAlignment w:val="center"/>' +
      '<w:contextualSpacing/></w:pPr>'
    const doc = await parseDocx(
      await buildDocx({ bodyXml: `<w:p>${afterSave}<w:r><w:t>undo</w:t></w:r></w:p>` }),
    )
    const reverted = mergePPrFormat(afterSave, {}, doc.blocks[0].format)
    for (const tag of ['suppressAutoHyphens', 'kinsoku', 'topLinePunct', 'textAlignment'])
      expect(reverted).not.toContain(`<w:${tag}`)
    expect(reverted).toContain('<w:contextualSpacing/>')
  })

  it('round-trips the Word Mac document-wide compression modes without changing other settings', async () => {
    const source = await buildDocx({ bodyXml: '<w:p><w:r><w:t>本文</w:t></w:r></w:p>' })
    const zip = await JSZip.loadAsync(source)
    zip.file(
      'word/settings.xml',
      '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
        '<w:defaultTabStop w:val="720"/><w:characterSpacingControl w:val="compressPunctuation"/>' +
        '<w:autoHyphenation/></w:settings>',
    )
    const parsed = await parseDocx(await zip.generateAsync({ type: 'uint8array' }))
    expect(parsed.characterSpacingControl).toBe('compressPunctuation')
    const settingsOnly = await saveDocx(parsed, [{ kind: 'original', docxIndex: 0 }], {
      characterSpacingControl: 'doNotCompress',
    })
    const settingsOnlyZip = await JSZip.loadAsync(settingsOnly)
    expect(await settingsOnlyZip.file('word/document.xml')!.async('string')).toBe(
      await zip.file('word/document.xml')!.async('string'),
    )
    expect(await settingsOnlyZip.file('word/settings.xml')!.async('string')).toContain(
      '<w:characterSpacingControl w:val="doNotCompress"/>',
    )
    const save = async (mode: 'doNotCompress' | 'compressPunctuationAndJapaneseKana') => {
      const result = await saveDocx(parsed, [{ kind: 'original', docxIndex: 0 }], {
        characterSpacingControl: mode,
        defaultTabStopTwips: 480,
      })
      const reopened = await parseDocx(result)
      const settings = await (
        await JSZip.loadAsync(result)
      )
        .file('word/settings.xml')!
        .async('string')
      expect(reopened.characterSpacingControl).toBe(mode)
      expect(settings).toContain('<w:defaultTabStop w:val="480"/>')
      expect(reopened.defaultTabStopTwips).toBe(480)
      expect(settings).toContain('<w:autoHyphenation/>')
      expect(settings).toContain(`<w:characterSpacingControl w:val="${mode}"/>`)
      expect(settings.match(/<w:characterSpacingControl/g)).toHaveLength(1)
    }
    await save('doNotCompress')
    await save('compressPunctuationAndJapaneseKana')
  })
})
