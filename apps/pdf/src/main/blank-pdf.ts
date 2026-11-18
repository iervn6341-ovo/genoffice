import { PDFDocument } from 'pdf-lib'

/** A4 portrait in PDF points */
const A4_SIZE: [number, number] = [595.28, 841.89]

/** Minimal single-page A4 PDF kept in the view's memory until the user saves. */
export async function blankPdfBuffer(): Promise<Buffer> {
  const doc = await PDFDocument.create()
  doc.addPage(A4_SIZE)
  return Buffer.from(await doc.save())
}
