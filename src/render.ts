// Canvas 2D backend: draws the positioned elements of one laid-out page.
// The context is scaled so that one unit is one point, the unit the engine lays out in.

import type { PageData, PageElement, Paint, Rgba, TextElement } from './engine.ts'
import type { PageFont } from './fonts.ts'

export interface RenderContext {
  /** How to draw with a font id of the layout the page belongs to. */
  font(id: number): PageFont | undefined
  /** A decoded image, or undefined while it is still loading. */
  image(id: string): CanvasImageSource | undefined
}

function css(color: Rgba): string {
  const [r, g, b, a] = color
  return `rgba(${Math.round(r * 255)},${Math.round(g * 255)},${Math.round(b * 255)},${a})`
}

function style(ctx: CanvasRenderingContext2D, paint: Paint): string | CanvasGradient | null {
  switch (paint.k) {
    case 'solid':
      return css(paint.c)
    case 'linear': {
      const gradient = ctx.createLinearGradient(paint.x1, paint.y1, paint.x2, paint.y2)
      for (const [offset, color] of paint.stops) gradient.addColorStop(offset, css(color))
      return gradient
    }
    case 'radial': {
      const gradient = ctx.createRadialGradient(paint.cx, paint.cy, 0, paint.cx, paint.cy, paint.r)
      for (const [offset, color] of paint.stops) gradient.addColorStop(offset, css(color))
      return gradient
    }
    default:
      return null
  }
}

/** The slant of a synthetic italic, the one the engine's own backends use. */
const SLANT = 0.2126

/**
 * How far, in points, a string drawn in one call may end from where the engine's advances
 * put its end. The canvas kerns some pairs that the engine does not, by up to a point.
 */
const WIDTH_SLACK = 0.25

/** The width the canvas itself gives a string, by font and text. Measuring is not free, and a page is drawn many times. */
const naturalWidths = new Map<string, number>()

/**
 * Draws a stretch of text: characters of one style that follow each other on a line, with
 * the advance the engine gave each. The stretch is `style`'s own text, or that of several
 * runs that continue one another.
 *
 * The engine's advances decide where every character is, so that lines are as wide as the
 * layout says. Drawing character by character does that, but each call has a cost, and on
 * Linux the runtime takes close to a millisecond for one Latin letter: a page took half a
 * second. So a stretch is drawn in one call wherever the canvas spaces it as the engine
 * does, and is halved, again and again, where it does not.
 */
function drawStretch(
  ctx: CanvasRenderingContext2D,
  style: TextElement,
  characters: string[],
  advances: number[],
  x: number,
  context: RenderContext,
): void {
  const font = context.font(style.f)
  // Ask for the weight and style the face really has: where the engine synthesises them, so
  // does this, the same way on every platform, and nothing is thickened or slanted twice.
  const bold = style.b && !font?.synBold
  const italic = style.i && !font?.synItalic
  const cssFont = `${italic ? 'italic ' : ''}${bold ? 'bold ' : ''}${style.s}px ${font?.family ?? 'sans-serif'}`
  ctx.font = cssFont
  ctx.fillStyle = css(style.c)
  const thicken = style.b && font?.synBold
  const slant = style.i && font?.synItalic
  if (thicken || slant) ctx.save()
  if (slant) ctx.transform(1, 0, -SLANT, 1, style.y * SLANT, 0)
  if (thicken) {
    ctx.strokeStyle = css(style.c)
    ctx.lineWidth = style.s / 24
    ctx.lineJoin = 'round'
  }
  const draw = (text: string, at: number): void => {
    if (!text.trim()) return
    ctx.fillText(text, at, style.y)
    if (thicken) ctx.strokeText(text, at, style.y)
  }
  const natural = (text: string): number => {
    const key = `${cssFont}\n${text}`
    let width = naturalWidths.get(key)
    if (width === undefined) {
      if (naturalWidths.size > 50000) naturalWidths.clear()
      width = ctx.measureText(text).width
      naturalWidths.set(key, width)
    }
    return width
  }
  if (characters.length !== advances.length) {
    // Shaping merged or split glyphs: the advances no longer map to characters one to one.
    draw(characters.join(''), x)
  } else {
    // starts[k] is where character k begins.
    const starts = [x]
    for (const advance of advances) starts.push(starts[starts.length - 1] + advance)
    const place = (from: number, to: number): void => {
      const text = characters.slice(from, to).join('')
      if (to - from <= 1 || Math.abs(natural(text) - (starts[to] - starts[from])) <= WIDTH_SLACK) {
        draw(text, starts[from])
        return
      }
      const middle = (from + to) >> 1
      place(from, middle)
      place(middle, to)
    }
    if (characters.length > 0) place(0, characters.length)
  }
  if (thicken || slant) ctx.restore()
}

/** Whether a text element can be split into its characters, one advance each. */
function plain(run: TextElement, characters: string[]): boolean {
  return characters.length === run.adv.length
}

/** Whether two text elements look the same, so that one may be drawn on from the other. */
function sameStyle(a: TextElement, b: TextElement): boolean {
  return (
    a.f === b.f &&
    a.s === b.s &&
    a.b === b.b &&
    a.i === b.i &&
    a.y === b.y &&
    a.c[0] === b.c[0] &&
    a.c[1] === b.c[1] &&
    a.c[2] === b.c[2] &&
    a.c[3] === b.c[3]
  )
}

/**
 * Draws the text element at `index` together with the ones after it that continue it on its
 * line in the same style, and returns the index of the last one drawn. The engine cuts a
 * line into many runs, often one a word; joined, they take far fewer calls.
 */
function drawText(ctx: CanvasRenderingContext2D, elements: PageElement[], index: number, context: RenderContext): number {
  const first = elements[index] as TextElement
  const characters = Array.from(first.txt)
  if (!plain(first, characters)) {
    drawStretch(ctx, first, characters, first.adv, first.x, context)
    return index
  }
  const advances = [...first.adv]
  let end = first.x + first.adv.reduce((sum, advance) => sum + advance, 0)
  let last = index
  for (;;) {
    const next = elements[last + 1]
    if (next?.t !== 'text' || !sameStyle(first, next) || Math.abs(next.x - end) > 0.02) break
    const more = Array.from(next.txt)
    if (!plain(next, more)) break
    characters.push(...more)
    for (const advance of next.adv) {
      advances.push(advance)
      end += advance
    }
    last++
  }
  drawStretch(ctx, first, characters, advances, first.x, context)
  return last
}

function drawElements(ctx: CanvasRenderingContext2D, elements: PageElement[], context: RenderContext): void {
  for (let index = 0; index < elements.length; index++) {
    const element = elements[index]
    switch (element.t) {
      case 'text':
        index = drawText(ctx, elements, index, context)
        break
      case 'rect':
        ctx.fillStyle = css(element.c)
        ctx.fillRect(element.x, element.y, element.w, element.h)
        break
      case 'line':
        ctx.save()
        ctx.strokeStyle = css(element.c)
        ctx.lineWidth = element.w
        if (element.dash) ctx.setLineDash(element.dash)
        ctx.beginPath()
        ctx.moveTo(element.x1, element.y1)
        ctx.lineTo(element.x2, element.y2)
        ctx.stroke()
        ctx.restore()
        break
      case 'img': {
        const image = context.image(element.id)
        if (image) ctx.drawImage(image, element.x, element.y, element.w, element.h)
        break
      }
      case 'path': {
        const path = new Path2D(element.d)
        const fill = element.fill && style(ctx, element.fill)
        if (fill) {
          ctx.fillStyle = fill
          ctx.fill(path, element.rule ?? 'nonzero')
        }
        const stroke = element.stroke && style(ctx, element.stroke.paint)
        if (element.stroke && stroke) {
          ctx.save()
          ctx.strokeStyle = stroke
          ctx.lineWidth = element.stroke.w
          ctx.lineCap = element.stroke.cap
          ctx.lineJoin = element.stroke.join
          if (element.stroke.dash) ctx.setLineDash(element.stroke.dash)
          ctx.stroke(path)
          ctx.restore()
        }
        break
      }
      case 'group':
        ctx.save()
        ctx.transform(...element.m)
        ctx.globalAlpha *= element.op
        if (element.clip) ctx.clip(new Path2D(element.clip))
        drawElements(ctx, element.els, context)
        ctx.restore()
        break
    }
  }
}

/** Draws a page onto a canvas whose backing store is `scale` device pixels per point. */
export function renderPage(canvas: HTMLCanvasElement, page: PageData, scale: number, context: RenderContext): void {
  const ctx = canvas.getContext('2d')
  if (!ctx) return
  ctx.setTransform(scale, 0, 0, scale, 0, 0)
  ctx.fillStyle = page.bg ? css(page.bg) : '#ffffff'
  ctx.fillRect(0, 0, page.w, page.h)
  ctx.textBaseline = 'alphabetic'
  drawElements(ctx, page.els, context)
}

/** The ids of every image a page draws, groups included. */
export function imageIds(elements: PageElement[], into: Map<string, string> = new Map()): Map<string, string> {
  for (const element of elements) {
    if (element.t === 'img') into.set(element.id, element.mime)
    else if (element.t === 'group') imageIds(element.els, into)
  }
  return into
}
