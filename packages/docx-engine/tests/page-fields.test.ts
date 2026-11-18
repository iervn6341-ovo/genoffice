import { expect, it } from 'vitest'
import { removePageFields } from '../src/page-fields'
it('removes PAGE in native layout tables and preserves NUMPAGES and surrounding text', () => {
  const page = '<w:fldSimple w:instr=" PAGE "><w:r><w:t>9</w:t></w:r></w:fldSimple>'
  const total = page.replace(' PAGE ', ' NUMPAGES ')
  const text = '<w:r><w:t>Private</w:t></w:r>'
  expect(
    removePageFields(`<w:tbl><w:tr><w:tc><w:p>${text}${page}${total}</w:p></w:tc></w:tr></w:tbl>`),
  ).toBe(`<w:tbl><w:tr><w:tc><w:p>${text}${total}</w:p></w:tc></w:tr></w:tbl>`)
  const complex =
    '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>5</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>'
  expect(removePageFields(`<w:p>${text}${complex}${total}</w:p>`)).toBe(
    `<w:p>${text}${total}</w:p>`,
  )
})

it('keeps visible text sharing the boundary runs of a PAGE field', () => {
  const xml =
    '<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Before</w:t><w:fldChar w:fldCharType="begin"/></w:r>' +
    '<w:r><w:instrText> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>5</w:t></w:r>' +
    '<w:r><w:rPr><w:i/></w:rPr><w:fldChar w:fldCharType="end"/><w:t>After</w:t></w:r></w:p>'
  expect(removePageFields(xml)).toBe(
    '<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Before</w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t>After</w:t></w:r></w:p>',
  )
})
