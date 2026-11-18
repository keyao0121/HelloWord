// Registers the fonts a layout draws with as page fonts, so Canvas 2D can use the very files
// the engine measured with. Font ids are local to one layout, so faces are cached by identity.
//
// The faces are declared with CSS @font-face rules over blob URLs. In acdert 0.0.9 a FontFace
// built from an ArrayBuffer reports "loaded" but neither DOM text nor canvas text uses it.

import type { Doc, FontInfo } from './engine.ts'

const families = new Map<string, string>()
let sheet: HTMLStyleElement | null = null
let probe: CanvasRenderingContext2D | null = null

/** Latin letters of clearly different widths: a proportional face cannot match a monospace on them. */
const SAMPLE = 'mmmmmmmmmmiiiiiiiiii'
/**
 * How many times to look for a face, 16 milliseconds apart, before drawing without it. A face
 * with no Latin letters, such as a symbol font, never shows up in the measurement, so the
 * wait is short; and it runs on a timer, which keeps going in a window that is not on screen
 * while animation frames stop.
 */
const WAIT_TRIES = 20

function identity(font: FontInfo): string {
  return `${font.family}|${font.bold}|${font.italic}|${font.faceIndex}|${font.bytes}`
}

/**
 * Resolves once canvas text really uses a face. document.fonts.load() resolves earlier than
 * that in a window on screen, and a page drawn in between shows the fallback font at the
 * engine's advances: every word too tight or too loose.
 */
async function inUse(family: string, weight: string, style: string): Promise<void> {
  probe ??= document.createElement('canvas').getContext('2d')
  if (!probe) return
  const prefix = `${style} ${weight} 40px`
  probe.font = `${prefix} monospace`
  const fallback = probe.measureText(SAMPLE).width
  for (let attempt = 0; attempt < WAIT_TRIES; attempt++) {
    probe.font = `${prefix} ${family}, monospace`
    if (probe.measureText(SAMPLE).width !== fallback) return
    await new Promise((resolve) => setTimeout(resolve, 16))
  }
}

/** How canvas text draws with one font of a layout. */
export interface PageFont {
  /** CSS family name of the face. */
  family: string
  /** Bold is drawn by thickening the outlines of a face that is not bold. */
  synBold: boolean
  /** Italic is drawn by slanting a face that is not italic. */
  synItalic: boolean
}

/** Returns how to draw with each font id of the document's current layout. */
export async function registerFonts(doc: Doc): Promise<Map<number, PageFont>> {
  const byId = new Map<number, PageFont>()
  const loading: Promise<unknown>[] = []
  for (const font of doc.fonts()) {
    const key = identity(font)
    let family = families.get(key)
    if (!family) {
      family = `rdocx-font-${families.size}`
      families.set(key, family)
      const url = URL.createObjectURL(new Blob([doc.fontData(font.id)], { type: 'font/ttf' }))
      const weight = font.bold ? 'bold' : 'normal'
      const style = font.italic ? 'italic' : 'normal'
      if (!sheet) {
        sheet = document.createElement('style')
        document.head.append(sheet)
      }
      sheet.append(
        `@font-face{font-family:${family};src:url(${url});font-weight:${weight};font-style:${style}}\n`,
      )
      const name = family
      loading.push(document.fonts.load(`${style} ${weight} 12px ${name}`).then(() => inUse(name, weight, style)))
    }
    byId.set(font.id, { family, synBold: font.synBold, synItalic: font.synItalic })
  }
  await Promise.all(loading)
  return byId
}
