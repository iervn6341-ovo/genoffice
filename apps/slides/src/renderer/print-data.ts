/**
 * Everything a print job needs besides the options: slide bitmaps, notes text, native notes/handout
 * fields, master artwork bitmaps and comments. Shared by the in-app print sheet (preview + print)
 * and the macOS direct-to-panel route.
 */
import type { RenderFill, RenderNode, RenderSlide } from '@genoffice/pptx-render'
import type { NotesHandoutsPrintFields } from '../shared/ipc'
import type { PrintCommentItem } from '../shared/print-html'
import { renderSlidesToPngBase64 } from './export-render'

export interface SlidesPrintData {
  /** print-quality (2x) bitmap per deck slide, base64 PNG */
  rendered: string[]
  notes: string[]
  nativeFields: NotesHandoutsPrintFields
  masterArtworkBase64: Record<string, string>
  firstSlideNum: number
  comments: PrintCommentItem[]
}

export async function loadSlidesPrintData(
  slides: RenderSlide[],
  images: Map<string, HTMLImageElement>,
  current: number,
): Promise<SlidesPrintData> {
  const rendered = await renderSlidesToPngBase64(slides, images)
  const [notes, nativeFields, hfState] = await Promise.all([
    Promise.all(slides.map((_s, i) => window.slidesApi.getNotes(i))),
    window.slidesApi.getNotesHandoutsPrintFields(),
    window.slidesApi.getHeaderFooter(current),
  ])
  const artwork = nativeFields.masterArtwork ?? {}
  const artworkSlides = Object.values(artwork)
  const artworkImages = new Map(images)
  const urls = new Set<string>()
  const addFill = (fill: RenderFill | undefined) => {
    if (fill?.kind === 'image' && fill.dataUrl) urls.add(fill.dataUrl)
  }
  const walk = (nodes: readonly RenderNode[]) => {
    for (const node of nodes) {
      if (node.type === 'picture' && node.dataUrl) urls.add(node.dataUrl)
      if ((node.type === 'shape' || node.type === 'text') && node.fill) addFill(node.fill)
      if (node.type === 'group') walk(node.children)
    }
  }
  for (const slide of artworkSlides) {
    addFill(slide.background)
    walk(slide.nodes)
  }
  await Promise.all(
    [...urls]
      .filter((url) => !artworkImages.has(url))
      .map(async (url) => {
        const image = new Image()
        image.src = url
        await image.decode()
        artworkImages.set(url, image)
      }),
  )
  const artworkPngs = artworkSlides.length
    ? await renderSlidesToPngBase64(artworkSlides, artworkImages)
    : []
  const masterArtworkBase64 = Object.fromEntries(
    Object.keys(artwork).map((path, i) => [path, artworkPngs[i]!]),
  )
  const commentLists = await Promise.all(slides.map((_s, i) => window.slidesApi.getComments(i)))
  return {
    rendered,
    notes,
    nativeFields,
    masterArtworkBase64,
    firstSlideNum: hfState.firstSlideNum,
    comments: commentLists.flatMap((list, i) =>
      list.map((c) => ({ number: i + 1, author: c.author, text: c.text })),
    ),
  }
}
