import { expect } from '@playwright/test'
import { copyFile, writeFile } from 'node:fs/promises'
import { join, resolve } from 'node:path'
import { PDFDocument, StandardFonts } from 'pdf-lib'
import { buildDocx } from '../packages/docx-engine/tests/helpers/build-docx'
import type { LaunchedApp } from './helpers'

export const kinds = ['docs', 'sheets', 'slides', 'pdf', 'markdown', 'html'] as const

export async function openFreshEditor(
  launched: LaunchedApp,
  path: string,
  kind: (typeof kinds)[number],
) {
  const before = new Set(launched.app.windows())
  await launched.page.evaluate((path) => (window as any).aiOffice.openPath(path), path)
  // Closed Docs views are deliberately detached instead of destroyed to avoid
  // an Electron native-loop freeze. Drive the newly attached view, not an orphan.
  await expect
    .poll(() =>
      launched.app
        .windows()
        .some((page) => !before.has(page) && page.url().includes(`://${kind}/`)),
    )
    .toBe(true)
  return launched.app
    .windows()
    .find((page) => !before.has(page) && page.url().includes(`://${kind}/`))!
}

export async function fixture(root: string, kind: (typeof kinds)[number]) {
  const extension = {
    docs: 'docx',
    sheets: 'xlsx',
    slides: 'pptx',
    pdf: 'pdf',
    markdown: 'md',
    html: 'html',
  }[kind]
  const path = join(root, `read-only.${extension}`)
  if (kind === 'docs') {
    await writeFile(
      path,
      await buildDocx({
        bodyXml:
          '<w:p><w:r><w:t>Preview without editing</w:t></w:r><w:r><w:footnoteReference w:id="8"/></w:r></w:p>' +
          '<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Formatted text</w:t></w:r></w:p><w:p/>' +
          '<w:p><w:r><w:t>Body paragraph for scrolling.</w:t></w:r></w:p>'.repeat(80),
        extraRels:
          '<Relationship Id="rFootnotes" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>',
        extraParts: [
          {
            path: 'word/footnotes.xml',
            contentType:
              'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml',
            xml: '<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:id="8"><w:p><w:r><w:footnoteRef/><w:t>Footnote body</w:t></w:r></w:p></w:footnote></w:footnotes>',
          },
        ],
      }),
    )
  } else if (kind === 'sheets')
    await copyFile(resolve('apps/sheets/fixtures/generated/compatibility-basic.xlsx'), path)
  else if (kind === 'slides')
    await copyFile(resolve('packages/pptx-engine/tests/fixtures/01_standard_business.pptx'), path)
  else if (kind === 'pdf') {
    const doc = await PDFDocument.create()
    const font = await doc.embedFont(StandardFonts.Helvetica)
    for (let i = 0; i < 3; i++) {
      const page = doc.addPage()
      page.drawText('Preview without editing', { x: 50, y: 700, font, size: 16 })
    }
    await writeFile(path, await doc.save())
  } else if (kind === 'markdown')
    await writeFile(
      path,
      '# Preview without editing\n\n' +
        '**Formatted text**\n\nBody paragraph for scrolling.\n\n'.repeat(80),
    )
  else
    await writeFile(
      path,
      '<!doctype html><html><body><h1>Preview without editing</h1>' +
        '<p>Body paragraph for scrolling.</p>'.repeat(80) +
        '</body></html>',
    )
  return path
}
