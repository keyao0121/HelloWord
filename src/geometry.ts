// Geometry between the laid-out page and the document: which character a point on a page is
// closest to, and where on a page a document position is drawn. A position is a paragraph,
// by its index in the layout's list of editable paragraphs, and a character offset inside
// it: the same addresses the engine reports and edits by. Every length is in points.

import type { PageData, Rgba, Structure } from './engine.ts'

export interface Caret {
  /** Index of the paragraph in `Structure.paras`. */
  para: number
  /** Character offset inside the paragraph's text. */
  offset: number
}

/** A drawn stretch of one paragraph's text. An empty paragraph draws one with no characters. */
interface RunBox {
  para: number
  start: number
  end: number
  baseline: number
  size: number
  bold: boolean
  italic: boolean
  color: Rgba
  /** edges[k] is the x of the boundary before character start + k; one more entry than characters. */
  edges: number[]
}

/** One line of text on a page. In a table it runs across the cells of a row. */
export interface Line {
  top: number
  bottom: number
  runs: RunBox[]
}

export interface CaretRect {
  /** Zero-based page index. */
  page: number
  line: number
  x: number
  top: number
  height: number
}

/** The runs of a page that belong to editable paragraphs. */
function pageRuns(page: PageData, structure: Structure): RunBox[] {
  const runs: RunBox[] = []
  for (const element of page.els) {
    if (element.t !== 'text' || !element.src) continue
    const para = structure.nodes[element.src[0] - 1] ?? -1
    if (para < 0) continue
    const [, start, end] = element.src
    const count = end - start
    if (count < 0) continue
    const edges = [element.x]
    if (element.adv.length === count) {
      for (const advance of element.adv) edges.push(edges[edges.length - 1] + advance)
    } else {
      // Shaping merged or split glyphs: spread the run's width evenly over its characters.
      const width = element.adv.reduce((sum, advance) => sum + advance, 0)
      for (let index = 1; index <= count; index++) edges.push(element.x + (width * index) / count)
    }
    runs.push({
      para,
      start,
      end,
      baseline: element.y,
      size: element.s,
      bold: element.b,
      italic: element.i,
      color: element.c,
      edges,
    })
  }
  return runs
}

/** The lines of one page that belong to editable paragraphs, top to bottom. */
export function pageLines(page: PageData, structure: Structure): Line[] {
  const byBaseline = new Map<number, RunBox[]>()
  for (const run of pageRuns(page, structure)) {
    const key = Math.round(run.baseline * 10)
    const runs = byBaseline.get(key)
    if (runs) runs.push(run)
    else byBaseline.set(key, [run])
  }
  const lines: Line[] = []
  for (const runs of byBaseline.values()) {
    runs.sort((a, b) => a.edges[0] - b.edges[0])
    const size = Math.max(...runs.map((run) => run.size))
    const baseline = runs[0].baseline
    lines.push({ top: baseline - size * 0.9, bottom: baseline + size * 0.3, runs })
  }
  // A picture stands on the baseline of its line and is most of the line's height: a click
  // on it, or beside it, belongs to that line.
  for (const element of page.els) {
    if (element.t !== 'img') continue
    const line = lines.find((candidate) => Math.abs(candidate.runs[0].baseline - (element.y + element.h)) < 0.5)
    if (line) line.top = Math.min(line.top, element.y)
  }
  return lines.sort((a, b) => a.top - b.top)
}

/** A picture drawn in the text of an editable paragraph. */
export interface PageImage {
  para: number
  /** Which of the paragraph's pictures on this page it is, from 0, in reading order. */
  order: number
  /** The character offset right after the picture. */
  offset: number
  x: number
  y: number
  width: number
  height: number
}

/** The pictures of a page that stand in the text of editable paragraphs, in reading order. */
export function pageImages(page: PageData, lines: Line[]): PageImage[] {
  const images: PageImage[] = []
  const counts = new Map<number, number>()
  for (const element of page.els) {
    if (element.t !== 'img') continue
    const right = element.x + element.w
    // The text that follows the picture on its line starts where the picture ends. In a
    // table one baseline runs across the cells of a row, so the closest text decides.
    let best: { para: number; offset: number } | null = null
    let distance = Infinity
    for (const line of lines) {
      for (const run of line.runs) {
        if (Math.abs(run.baseline - (element.y + element.h)) >= 0.5) continue
        run.edges.forEach((edge, index) => {
          const gap = Math.abs(edge - right)
          if (gap < distance) {
            distance = gap
            best = { para: run.para, offset: run.start + index }
          }
        })
      }
    }
    if (!best) continue
    const { para, offset } = best as { para: number; offset: number }
    const order = counts.get(para) ?? 0
    counts.set(para, order + 1)
    images.push({ para, order, offset, x: element.x, y: element.y, width: element.w, height: element.h })
  }
  return images
}

/** The top and the bottom, on one page, of the lines of a stretch of paragraphs; null when it draws none there. */
export function extentOf(lines: Line[], first: number, last: number): { top: number; bottom: number } | null {
  let extent: { top: number; bottom: number } | null = null
  for (const line of lines) {
    if (!line.runs.some((run) => run.para >= first && run.para <= last)) continue
    extent = extent
      ? { top: Math.min(extent.top, line.top), bottom: Math.max(extent.bottom, line.bottom) }
      : { top: line.top, bottom: line.bottom }
  }
  return extent
}

/** The position in one line closest to x. */
export function hitLine(line: Line, x: number): Caret {
  let best: Caret = { para: line.runs[0].para, offset: line.runs[0].start }
  let distance = Infinity
  for (const run of line.runs) {
    run.edges.forEach((edge, index) => {
      const gap = Math.abs(edge - x)
      if (gap < distance) {
        distance = gap
        best = { para: run.para, offset: run.start + index }
      }
    })
  }
  return best
}

/** The position on a page closest to a point, or null when the page has nothing editable. */
export function hitPage(lines: Line[], x: number, y: number): Caret | null {
  let best: Line | null = null
  let distance = Infinity
  for (const line of lines) {
    const gap = y < line.top ? line.top - y : y > line.bottom ? y - line.bottom : 0
    if (gap < distance) {
      distance = gap
      best = line
    }
  }
  return best && hitLine(best, x)
}

/** The zero-based pages a paragraph may be drawn on. */
function pagesOf(para: number, structure: Structure): number[] {
  return (structure.paras[para]?.pages ?? []).map((page) => page - 1)
}

/**
 * Where a position is drawn. At a line wrap the position belongs to the end of the earlier
 * line, matching typed text, which continues the run before the caret.
 */
export function caretRect(
  caret: Caret,
  structure: Structure,
  linesOf: (pageIndex: number) => Line[],
): CaretRect | null {
  let atStart: CaretRect | null = null
  let before: { rect: CaretRect; end: number } | null = null
  for (const page of pagesOf(caret.para, structure)) {
    const lines = linesOf(page)
    for (let index = 0; index < lines.length; index++) {
      for (const run of lines[index].runs) {
        if (run.para !== caret.para) continue
        const rect = (x: number): CaretRect => ({
          page,
          line: index,
          x,
          top: run.baseline - run.size * 0.85,
          height: run.size * 1.15,
        })
        if (caret.offset > run.start && caret.offset <= run.end) return rect(run.edges[caret.offset - run.start])
        if (caret.offset === run.start) atStart ??= rect(run.edges[0])
        else if (run.end < caret.offset && (!before || run.end > before.end)) {
          // Text the engine drew nothing for, such as a space at the end of a line.
          before = { rect: rect(run.edges[run.edges.length - 1]), end: run.end }
        }
      }
    }
  }
  return atStart ?? before?.rect ?? null
}

/** A rectangle on a page, in points. */
export interface PageRect {
  page: number
  x: number
  y: number
  width: number
  height: number
}

/** Orders two positions by where they are in the document. */
export function ordered(a: Caret, b: Caret): [Caret, Caret] {
  return a.para < b.para || (a.para === b.para && a.offset <= b.offset) ? [a, b] : [b, a]
}

/** The rectangles that highlight the text between two positions. */
export function selectionRects(
  from: Caret,
  to: Caret,
  structure: Structure,
  linesOf: (pageIndex: number) => Line[],
): PageRect[] {
  const pages = new Set<number>()
  for (let para = from.para; para <= to.para; para++) {
    for (const page of pagesOf(para, structure)) pages.add(page)
  }
  const rects: PageRect[] = []
  for (const page of [...pages].sort((a, b) => a - b)) {
    for (const line of linesOf(page)) {
      const height = line.bottom - line.top
      for (const run of line.runs) {
        if (run.para < from.para || run.para > to.para) continue
        if (run.start === run.end) {
          // An empty paragraph inside the selection shows as a sliver, like its paragraph mark.
          if (run.para < to.para) rects.push({ page, x: run.edges[0], y: line.top, width: 4, height })
          continue
        }
        const start = Math.max(run.start, run.para === from.para ? from.offset : 0)
        const end = Math.min(run.end, run.para === to.para ? to.offset : Infinity)
        if (start >= end) continue
        const x = run.edges[start - run.start]
        rects.push({ page, x, y: line.top, width: run.edges[end - run.start] - x, height })
      }
    }
  }
  return rects
}

/** How a character is drawn, as layout resolved it from its own format and its styles. */
export interface DrawnStyle {
  bold: boolean
  italic: boolean
  size: number
  color: Rgba
}

/** How the character at a position is drawn. */
export function styleAt(
  position: Caret,
  structure: Structure,
  linesOf: (pageIndex: number) => Line[],
): DrawnStyle | null {
  let empty: DrawnStyle | null = null
  for (const page of pagesOf(position.para, structure)) {
    for (const line of linesOf(page)) {
      for (const run of line.runs) {
        if (run.para !== position.para) continue
        const style = { bold: run.bold, italic: run.italic, size: run.size, color: run.color }
        if (position.offset >= run.start && position.offset < run.end) return style
        // An empty paragraph has a format of its own, the one typed text will take.
        if (run.start === run.end) empty = style
      }
    }
  }
  return empty
}

/** The box a paragraph's text takes on one page, in points. */
export interface ParagraphBox {
  para: number
  x: number
  y: number
  width: number
  height: number
  /** Whether the paragraph begins on this page, and not on an earlier one. */
  starts: boolean
}

/** The boxes of the editable paragraphs drawn on a page, top to bottom. */
export function paragraphBoxes(page: PageData, structure: Structure): ParagraphBox[] {
  const boxes = new Map<number, { left: number; top: number; right: number; bottom: number; starts: boolean }>()
  for (const run of pageRuns(page, structure)) {
    const left = run.edges[0]
    const right = run.edges[run.edges.length - 1]
    const top = run.baseline - run.size * 0.9
    const bottom = run.baseline + run.size * 0.3
    const box = boxes.get(run.para)
    if (!box) {
      boxes.set(run.para, { left, top, right, bottom, starts: run.start === 0 })
      continue
    }
    box.left = Math.min(box.left, left)
    box.top = Math.min(box.top, top)
    box.right = Math.max(box.right, right)
    box.bottom = Math.max(box.bottom, bottom)
    box.starts ||= run.start === 0
  }
  return [...boxes]
    .map(([para, box]) => ({
      para,
      x: box.left,
      y: box.top,
      width: Math.max(box.right - box.left, 4),
      height: box.bottom - box.top,
      starts: box.starts,
    }))
    .sort((a, b) => a.y - b.y || a.x - b.x)
}

/** A line of text on a page with the box it is drawn in, in points. */
export interface TextLine {
  x: number
  y: number
  width: number
  height: number
  text: string
}

/**
 * The lines a page draws outside the editable paragraphs: headers and footers, list
 * numbers, text boxes. They have no place in the structure the editor works on, so the page
 * reports them line by line.
 */
export function foreignLines(page: PageData, structure: Structure): TextLine[] {
  const byBaseline = new Map<number, { x: number; right: number; y: number; size: number; text: string }[]>()
  for (const element of page.els) {
    if (element.t !== 'text' || !element.txt.trim()) continue
    if (element.src && (structure.nodes[element.src[0] - 1] ?? -1) >= 0) continue
    const right = element.x + element.adv.reduce((sum, advance) => sum + advance, 0)
    const key = Math.round(element.y * 10)
    const run = { x: element.x, right, y: element.y, size: element.s, text: element.txt }
    const runs = byBaseline.get(key)
    if (runs) runs.push(run)
    else byBaseline.set(key, [run])
  }
  const lines: TextLine[] = []
  for (const runs of byBaseline.values()) {
    runs.sort((a, b) => a.x - b.x)
    const size = Math.max(...runs.map((run) => run.size))
    const x = runs[0].x
    lines.push({
      x,
      y: runs[0].y - size * 0.9,
      width: Math.max(...runs.map((run) => run.right)) - x,
      height: size * 1.2,
      text: runs.map((run) => run.text).join(''),
    })
  }
  return lines.sort((a, b) => a.y - b.y || a.x - b.x)
}
