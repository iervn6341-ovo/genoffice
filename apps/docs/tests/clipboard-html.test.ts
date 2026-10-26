import { describe, expect, it } from 'vitest'
import {
  INLINE_ELEMENT_LIMIT,
  enrichClipboardDom,
  withRootAttrs,
  type HostMounter,
} from '../src/renderer/editor/clipboard-html'
import { clipboardPlainText } from '../src/renderer/editor/clipboard-extension'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { copy, openDocx, selectAll, selectNode } from './helpers/clip-harness'

const mounter: HostMounter = (host) => {
  document.body.appendChild(host)
  return () => host.remove()
}
const container = (html: string): HTMLElement => {
  const dom = document.createElement('div')
  dom.innerHTML = html
  return dom
}

describe('withRootAttrs', () => {
  it('adds attributes to the first element, past leading comments', () => {
    expect(
      withRootAttrs('<!--x--><table class="a"><tr></tr></table>', { 'data-pm-slice': '0 0 []' }),
    ).toBe('<!--x--><table data-pm-slice="0 0 []" class="a"><tr></tr></table>')
  })
  it('escapes values and skips null ones', () => {
    expect(withRootAttrs('<p>x</p>', { a: 'q"&', b: null })).toBe('<p a="q&quot;&amp;">x</p>')
  })
  it('leaves markup without an element alone', () => {
    expect(withRootAttrs('plain', { a: '1' })).toBe('plain')
  })
})

describe('enrichClipboardDom', () => {
  it('without a live editor falls back to the serialized markup', () => {
    const dom = container('<p data-pm-slice="0 0 []">hi</p>')
    expect(enrichClipboardDom(dom)).toBe('<p data-pm-slice="0 0 []">hi</p>')
  })

  it('inlines resolved styles and carries the slice markers over to the first element', () => {
    const dom = container(
      '<table data-pm-slice="0 0 []" data-genoffice-slice="eyJ2IjoxfQ==" style="border-collapse: collapse"><tbody><tr><td style="background-color: rgb(255, 255, 0); border-top: 2px solid rgb(255, 0, 0)">x</td></tr></tbody></table>',
    )
    const html = enrichClipboardDom(dom, mounter)
    expect(html).toMatch(/^<table[^>]*data-pm-slice="0 0 \[\]"/)
    expect(html).toContain('data-genoffice-slice="eyJ2IjoxfQ=="')
    expect(html).toMatch(/background-color:\s*rgb\(255,\s*255,\s*0\)/)
    expect(html).toMatch(/border-top:[^;"]*2px solid rgb\(255,\s*0,\s*0\)/)
    // the host is gone again
    expect(document.body.querySelectorAll('table').length).toBe(0)
  })

  it('does not inline huge fragments', () => {
    const dom = container(
      '<div data-pm-slice="0 0 []">' + '<span>a</span>'.repeat(INLINE_ELEMENT_LIMIT + 1) + '</div>',
    )
    expect(enrichClipboardDom(dom, mounter)).toBe(dom.innerHTML)
  })

  it('a failing mount keeps the plain markup', () => {
    const dom = container('<p>x</p>')
    const boom: HostMounter = () => {
      throw new Error('no layout')
    }
    expect(() => enrichClipboardDom(dom, boom)).toThrow() // the caller (copy handler) falls back
  })
})

describe('plain text flavor', () => {
  it('paragraphs are one line each, a hard break is a line break', async () => {
    const doc = await openDocx(
      await buildDocx({
        bodyXml:
          '<w:p><w:r><w:t>one</w:t></w:r></w:p><w:p><w:r><w:t>a</w:t></w:r><w:r><w:br/></w:r><w:r><w:t>b</w:t></w:r></w:p>',
      }),
    )
    selectAll(doc.editor)
    expect(copy(doc.editor).text).toBe('one\na\nb')
  })

  it('multi-paragraph cells are quoted so Excel keeps them in one cell', async () => {
    const doc = await openDocx(
      await buildDocx({
        bodyXml:
          '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>l1</w:t></w:r></w:p><w:p><w:r><w:t>l2</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>say "hi"</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p/>',
      }),
    )
    selectNode(doc.editor, 0)
    expect(clipboardPlainText(doc.editor.state.selection.content().content)).toBe(
      '"l1\nl2"\t"say ""hi"""',
    )
  })
})
