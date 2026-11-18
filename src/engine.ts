// The document engine: rdocx compiled to wasm. All document state lives inside the wasm module;
// this file loads it and gives the rest of the app typed access to the layout it produces.

import init, { WasmDocument } from './wasm/rdocx_wasm.js'
import wasmUrl from './wasm/rdocx_wasm_bg.wasm?url'

export type Rgba = [number, number, number, number]

export type Paint =
  | { k: 'solid'; c: Rgba }
  | { k: 'linear'; x1: number; y1: number; x2: number; y2: number; stops: [number, Rgba][] }
  | { k: 'radial'; cx: number; cy: number; r: number; stops: [number, Rgba][] }
  | { k: 'none' }

export interface TextElement {
  t: 'text'
  x: number
  y: number
  /** Font id of the current layout; see fonts(). */
  f: number
  /** Font size in points. */
  s: number
  c: Rgba
  txt: string
  /** One advance per glyph, in points. */
  adv: number[]
  b: boolean
  i: boolean
  /** [source node id, first character, end character] when the run is a direct projection of document text. */
  src?: [number, number, number]
}

export interface LineElement {
  t: 'line'
  x1: number
  y1: number
  x2: number
  y2: number
  w: number
  c: Rgba
  dash?: [number, number]
}

export interface RectElement {
  t: 'rect'
  x: number
  y: number
  w: number
  h: number
  c: Rgba
}

export interface ImageElement {
  t: 'img'
  x: number
  y: number
  w: number
  h: number
  mime: string
  id: string
}

export interface PathElement {
  t: 'path'
  /** SVG path data. */
  d: string
  rule?: 'evenodd'
  fill?: Paint
  stroke?: { paint: Paint; w: number; cap: CanvasLineCap; join: CanvasLineJoin; dash?: number[] }
}

export interface GroupElement {
  t: 'group'
  /** Affine transform [a, b, c, d, e, f]. */
  m: [number, number, number, number, number, number]
  op: number
  clip?: string
  els: PageElement[]
}

export type PageElement = TextElement | LineElement | RectElement | ImageElement | PathElement | GroupElement

/** One laid-out page. Every length is in points (1/72 inch). */
export interface PageData {
  w: number
  h: number
  bg?: Rgba
  els: PageElement[]
}

/** One paragraph a caret can be put in, as the current layout placed it. */
export interface ParaInfo {
  /** Length in characters. */
  len: number
  /** Heading level; 0 when the paragraph is not a heading. */
  h: number
  /** 1 when the next paragraph in the list is this one's next sibling, and so can be joined onto it. */
  m: 0 | 1
  /** 1 when the paragraph is in a table cell. */
  t: 0 | 1
  /** 1 when the paragraph is an item of a list. */
  l: 0 | 1
  /** The page numbers, from 1, of the top-level block the paragraph belongs to. */
  pages: number[]
}

/** What maps the layout back to the document. */
export interface Structure {
  /** nodes[id - 1] is the paragraph a text run's source node id belongs to, or -1 when not editable. */
  nodes: number[]
  /**
   * Every editable paragraph in document order: top-level ones and those in table cells.
   * A paragraph's index here is its address in every editing call.
   */
  paras: ParaInfo[]
}

export type CharacterProperty =
  | 'bold'
  | 'italic'
  | 'underline'
  | 'strike'
  | 'script'
  | 'size'
  | 'color'
  | 'highlight'
  | 'font'
/** Where text stands on its line: raised small, lowered small, or neither. */
export type Script = 'super' | 'sub'
export type ListKind = 'bullet' | 'number'
export type Alignment = 'left' | 'center' | 'right' | 'justify'

export interface ParagraphFormat {
  underline: boolean
  strike: boolean
  script: Script | null
  /** The text's own highlight, by its Word name, or null. */
  highlight: string | null
  /** The font the text itself names, or null when it takes its style's. */
  font: string | null
  /** The colour the text itself sets, as six hex digits, or null when it takes its style's. */
  color: string | null
  list: ListKind | null
  /** Left indent in points. */
  indent: number
  /** Line spacing as a multiple of single spacing, or null when the paragraph takes its style's. */
  spacing: number | null
  /** null when the paragraph inherits its alignment from its style. */
  align: Alignment | null
  /** 0 when the paragraph is not a heading. */
  heading: number
  /** How far the first line starts in from the rest, in points. */
  firstLine: number
  /** Space before and after the paragraph in points, or null when it takes its style's. */
  before: number | null
  after: number | null
}

export interface FontInfo {
  id: number
  family: string
  bold: boolean
  italic: boolean
  /** The file is not a bold face: the engine expects bold to be drawn by thickening its outlines. */
  synBold: boolean
  /** The file is not an italic face: the engine expects italic to be drawn by slanting it. */
  synItalic: boolean
  faceIndex: number
  bytes: number
}

let ready: Promise<void> | null = null

/** Fetches and instantiates the wasm once. acdert has no instantiateStreaming, so the bytes are passed in. */
export function loadEngine(): Promise<void> {
  ready ??= (async () => {
    const response = await fetch(wasmUrl)
    if (!response.ok) throw new Error(`could not load the document engine (${response.status})`)
    await init({ module_or_path: await response.arrayBuffer() })
  })()
  return ready
}

/** Copies bytes out of the wasm memory view into a buffer of their own. */
function ownBuffer(bytes: Uint8Array): ArrayBuffer {
  const copy = new ArrayBuffer(bytes.byteLength)
  new Uint8Array(copy).set(bytes)
  return copy
}

export class Doc {
  private inner: WasmDocument

  private constructor(inner: WasmDocument) {
    this.inner = inner
  }

  static create(): Doc {
    return new Doc(new WasmDocument())
  }

  static open(bytes: Uint8Array): Doc {
    return new Doc(WasmDocument.fromBytes(bytes))
  }

  addHeading(text: string, level: number): void {
    this.inner.addHeading(text, level)
  }

  addParagraph(text: string): void {
    this.inner.addParagraph(text)
  }

  addTable(rows: number, columns: number): void {
    this.inner.addTable(rows, columns)
  }

  /** The document's plain text, paragraphs separated by line feeds. */
  text(): string {
    return this.inner.getText()
  }

  /** Lays the document out if it changed since the last call. */
  pageCount(): number {
    return this.inner.pageCount()
  }

  /** [width, height] in points for every page. */
  pageSizes(): [number, number][] {
    const flat = this.inner.pageSizes()
    const sizes: [number, number][] = []
    for (let index = 0; index + 1 < flat.length; index += 2) sizes.push([flat[index], flat[index + 1]])
    return sizes
  }

  page(index: number): PageData {
    return JSON.parse(this.inner.pageJson(index)) as PageData
  }

  /** A page as the text `page` parses: equal texts are pages that draw the same. */
  pageJson(index: number): string {
    return this.inner.pageJson(index)
  }

  structure(): Structure {
    return JSON.parse(this.inner.structureJson()) as Structure
  }

  insertText(para: number, offset: number, text: string): void {
    this.inner.insertText(para, offset, text)
  }

  /** Deletes the characters [start, end) of a paragraph. */
  deleteRange(para: number, start: number, end: number): void {
    this.inner.deleteRange(para, start, end)
  }

  /** Splits a paragraph at an offset; the second half becomes paragraph `para + 1`. */
  splitParagraph(para: number, offset: number): void {
    this.inner.splitParagraph(para, offset)
  }

  /** Joins paragraph `para + 1` onto `para` and returns the offset of the join. */
  mergeParagraphs(para: number): number {
    return this.inner.mergeParagraphs(para)
  }

  /** Marks the start of one undoable edit; every change until commitEdit is one step. */
  beginEdit(): void {
    this.inner.beginEdit()
  }

  /**
   * Closes the edit. Returns 0 when nothing changed, 1 for a new undo step, and 2 when
   * `merge` folded the change into the previous step.
   */
  commitEdit(merge: boolean): number {
    return this.inner.commitEdit(merge)
  }

  /** Takes back the last step. Returns false when there is none. */
  undo(): boolean {
    return this.inner.undo()
  }

  redo(): boolean {
    return this.inner.redo()
  }

  canUndo(): boolean {
    return this.inner.canUndo()
  }

  canRedo(): boolean {
    return this.inner.canRedo()
  }

  /** Replaces every occurrence of a text in the document and returns how many there were. */
  replaceText(find: string, replacement: string): number {
    return this.inner.replaceText(find, replacement)
  }

  /** Removes a paragraph. A table cell keeps its last one. */
  removeParagraph(para: number): void {
    this.inner.removeParagraph(para)
  }

  /** Inserts a table after the top-level paragraph `para`; its first cell becomes paragraph `para + 1`. */
  insertTable(para: number, rows: number, columns: number): void {
    this.inner.insertTable(para, rows, columns)
  }

  /** Inserts a picture in a paragraph of its own, after the top-level block that holds `para`. The size is in points. */
  insertImage(para: number, data: Uint8Array, filename: string, width: number, height: number): void {
    this.inner.insertImage(para, data, filename, width, height)
  }

  /** The size in points, `[width, height]`, of each inline picture of a paragraph, in document order. */
  images(para: number): [number, number][] {
    return JSON.parse(this.inner.imagesJson(para)) as [number, number][]
  }

  /** Sets the size, in points, of picture `index` of a paragraph. */
  resizeImage(para: number, index: number, width: number, height: number): void {
    this.inner.resizeImage(para, index, width, height)
  }

  /** Deletes picture `index` of a paragraph; the paragraph stays. */
  deleteImage(para: number, index: number): void {
    this.inner.deleteImage(para, index)
  }

  /**
   * Moves picture `index` of a paragraph into a paragraph of its own, before or after the
   * top-level block that holds paragraph `target`, and returns that new paragraph's index.
   * A paragraph the picture leaves empty goes with it.
   */
  moveImage(para: number, index: number, target: number, after: boolean): number {
    return this.inner.moveImage(para, index, target, after)
  }

  /** A paragraph's text, in the characters its offsets count. */
  paragraphText(para: number): string {
    return this.inner.paragraphText(para)
  }

  /** Sets one character property on the characters [start, end) of a paragraph. */
  formatRange(para: number, start: number, end: number, property: CharacterProperty, value: string): void {
    this.inner.formatRange(para, start, end, property, value)
  }

  setAlignment(para: number, alignment: Alignment): void {
    this.inner.setAlignment(para, alignment)
  }

  /** Makes a paragraph an item of a bullet or a numbered list, or takes it out of its list. */
  setList(para: number, kind: ListKind | 'none'): void {
    this.inner.setList(para, kind)
  }

  /** Sets a paragraph's left indent in points; 0 removes it. */
  setIndent(para: number, points: number): void {
    this.inner.setIndent(para, points)
  }

  /** Sets a paragraph's line spacing as a multiple of single spacing. */
  setLineSpacing(para: number, multiple: number): void {
    this.inner.setLineSpacing(para, multiple)
  }

  /** Sets how far a paragraph's first line starts in from the rest, in points; 0 removes it. */
  setFirstLineIndent(para: number, points: number): void {
    this.inner.setFirstLineIndent(para, points)
  }

  /** Sets the space before and after a paragraph in points; a negative number leaves that side to the style. */
  setParagraphSpacing(para: number, before: number, after: number): void {
    this.inner.setParagraphSpacing(para, before, after)
  }

  /** Takes every character format the text itself carries off a stretch of a paragraph. */
  clearFormat(para: number, start: number, end: number): void {
    this.inner.formatRange(para, start, end, 'clear', '')
  }

  // The table edits below take a paragraph in a cell of a table and return the index of the
  // paragraph the caret belongs in afterwards.

  /** Inserts an empty row above or below the paragraph's row. */
  insertTableRow(para: number, below: boolean): number {
    return this.inner.insertTableRow(para, below)
  }

  /** Deletes the paragraph's row; the table goes with its last row. */
  deleteTableRow(para: number): number {
    return this.inner.deleteTableRow(para)
  }

  /** Inserts an empty column left or right of the paragraph's cell. */
  insertTableColumn(para: number, after: boolean): number {
    return this.inner.insertTableColumn(para, after)
  }

  /** Deletes the column of the paragraph's cell; the table goes with its last column. */
  deleteTableColumn(para: number): number {
    return this.inner.deleteTableColumn(para)
  }

  /** Merges the cells between the cells of two paragraphs into one. */
  mergeTableCells(from: number, to: number): number {
    return this.inner.mergeTableCells(from, to)
  }

  /** Splits the merged cell of a paragraph back into the cells it covers. */
  splitTableCell(para: number): number {
    return this.inner.splitTableCell(para)
  }

  /** Starts a paragraph on a new page, or stops doing so. */
  setPageBreakBefore(para: number, on: boolean): void {
    this.inner.setPageBreakBefore(para, on)
  }

  /** Gives a paragraph the heading style of a level from 1, or the default style for 0. */
  setHeading(para: number, level: number): void {
    this.inner.setHeading(para, level)
  }

  /** What the document itself says about a position: underline, own alignment, heading level. */
  format(para: number, offset: number): ParagraphFormat {
    return JSON.parse(this.inner.formatJson(para, offset)) as ParagraphFormat
  }

  fonts(): FontInfo[] {
    return JSON.parse(this.inner.fontsJson()) as FontInfo[]
  }

  fontData(id: number): ArrayBuffer {
    return ownBuffer(this.inner.fontData(id))
  }

  mediaData(page: number, id: string): ArrayBuffer {
    return ownBuffer(this.inner.mediaData(page, id))
  }

  toDocx(): ArrayBuffer {
    return ownBuffer(this.inner.toDocxBytes())
  }

  toPdf(): ArrayBuffer {
    return ownBuffer(this.inner.toPdf())
  }

  free(): void {
    this.inner.free()
  }
}
