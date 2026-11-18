// The editing layer: a caret and a selection in the document, moved by the pointer and the
// keyboard, and text typed, deleted, split and formatted at them. Every edit goes to the
// engine, which lays the document out again; the view then redraws and the caret and the
// selection are placed from the new layout.

import type { Alignment, CharacterProperty, Doc, ListKind, Rgba, Script, Structure } from './engine.ts'
import {
  type Caret,
  type Line,
  type PageImage,
  caretRect,
  extentOf,
  hitLine,
  hitPage,
  ordered,
  pageImages,
  pageLines,
  selectionRects,
  styleAt,
} from './geometry.ts'
import type { Corner, PageView, PointerInfo } from './view.ts'

/** What a toolbar shows for the caret or the start of the selection. */
export interface FormatState {
  bold: boolean
  italic: boolean
  underline: boolean
  strike: boolean
  script: Script | null
  /** Font size in points, or null where no text is drawn. */
  size: number | null
  /** The font the text itself names, or null when it takes its style's. */
  font: string | null
  /** The colour the text is drawn in, as six hex digits, or null where no text is drawn. */
  color: string | null
  /** The text's own highlight, by its Word name, or null. */
  highlight: string | null
  /** null when the paragraph inherits its alignment. */
  align: Alignment | null
  heading: number
  list: ListKind | null
  /** Left indent in points. */
  indent: number
  /** Line spacing as a multiple of single spacing, or null when inherited. */
  spacing: number | null
  /** How far the first line starts in from the rest, in points. */
  firstLine: number
  /** Space before and after the paragraph in points, or null when inherited. */
  before: number | null
  after: number | null
  /** Whether the caret is in a table. */
  inTable: boolean
  /** Whether the format painter holds a format for the text selected next. */
  painting: boolean
}

/** The paragraph formats the paragraph dialog sets; a side left out stays as it is. */
export interface ParagraphSettings {
  /** Points; 0 removes the indent. */
  firstLine?: number
  /** Points; null gives the side back to the paragraph's style. */
  before?: number | null
  after?: number | null
}

/** Where a search found its text. */
export interface FindResult {
  /** Which occurrence is selected, from 1; 0 when there is none. */
  index: number
  total: number
}

interface Match {
  para: number
  start: number
  end: number
}

const WORD_CHARACTER = /[\p{L}\p{N}_]/u

/** Splits text into words the way its language does; Chinese has no spaces to go by. */
const segmenter =
  typeof Intl !== 'undefined' && 'Segmenter' in Intl ? new Intl.Segmenter('zh', { granularity: 'word' }) : null

/** As many steps as the engine keeps. */
const UNDO_LIMIT = 200

/** One step of the indent buttons, in points: half an inch. */
export const INDENT_STEP = 36

/** The shorter side a picture can be dragged down to, in points. */
const MIN_PICTURE = 12
/** How far, in points, the pointer travels with a picture before the picture follows. */
const DRAG_SLACK = 4

/** A picture in the text: its paragraph, and which of that paragraph's pictures it is. */
export interface Picked {
  para: number
  index: number
}

/** What the pointer is doing to the selected picture while its button is held. */
type Drag =
  | {
      kind: 'move'
      page: number
      x: number
      y: number
      /** Whether the pointer has gone far enough for the picture to follow. */
      moving: boolean
      target: { para: number; after: boolean } | null
    }
  | { kind: 'size'; page: number; x: number; y: number; corner: Corner; image: PageImage; width: number; height: number }

function same(a: Caret, b: Caret): boolean {
  return a.para === b.para && a.offset === b.offset
}

function hex(color: Rgba): string {
  return color
    .slice(0, 3)
    .map((channel) => Math.round(channel * 255).toString(16).padStart(2, '0'))
    .join('')
    .toUpperCase()
}

interface Ends {
  anchor: Caret
  caret: Caret
}

/** Where the selection was before and after one undo step. */
interface Step {
  before: Ends
  after: Ends
}

export class Editor {
  /** Called with a message when an edit is refused. */
  onMessage: (message: string) => void = () => {}
  /** Called after every edit that changed the document. */
  onChange: () => void = () => {}
  /** Called when the caret or the selection moved, or the format under it changed. */
  onState: (state: FormatState | null) => void = () => {}

  private doc: Doc
  private view: PageView
  private structure: Structure
  private lines = new Map<number, Line[]>()
  private images = new Map<number, PageImage[]>()
  /** The selected picture, shown in a frame; the caret then waits right after it. */
  private picked: Picked | null = null
  private drag: Drag | null = null
  /** Paragraph texts read since the last edit. */
  private texts = new Map<number, string>()
  /** Where the selection started; equal to the caret when nothing is selected. */
  private anchor: Caret | null = null
  private caret: Caret | null = null
  private composing = false
  /** Whether this editor has the view: of the editors of the open documents, one does. */
  private active = false
  private disposed = false
  /** One entry per undo step the engine holds, in the same order. */
  private undone: Step[] = []
  private redone: Step[] = []
  /** Where the caret was left by the last typed text; typing on from there joins its step. */
  private typedAt: Caret | null = null
  /** Formats chosen with nothing selected: the next typed text takes them. */
  private pending: Partial<Record<CharacterProperty, string>> = {}
  /** The format the painter picked up; the text selected next takes it. */
  private painter: Partial<Record<CharacterProperty, string>> | null = null
  /** The horizontal position a run of up and down moves keeps returning to. */
  private goalX: number | null = null
  /** Edits run one after another: each one waits for the redraw of the one before. */
  private queue: Promise<void> = Promise.resolve()

  constructor(doc: Doc, view: PageView) {
    this.doc = doc
    this.view = view
    this.structure = view.structure
  }

  /**
   * Takes the view, which must already show this editor's document: the pointer and the
   * keyboard now edit it. The caret, the selection and the undo history are where they were
   * left; a document shown for the first time gets its caret at the start.
   */
  activate(): void {
    if (this.disposed || this.active) return
    this.active = true
    const { view } = this
    this.structure = view.structure
    this.lines.clear()
    this.images.clear()
    this.texts.clear()
    this.drag = null
    view.onPointer = (page, x, y, info) => this.pointer(page, x, y, info)
    const input = view.input
    input.value = ''
    input.onkeydown = (event) => this.keyDown(event)
    input.oninput = () => {
      if (this.composing) return
      this.takeInput(input.value)
    }
    input.addEventListener('compositionstart', this.compositionStart)
    input.addEventListener('compositionend', this.compositionEnd)
    if (!this.caret || !this.anchor) {
      this.moveToEdge(false)
      return
    }
    this.anchor = this.valid(this.anchor)
    this.caret = this.valid(this.caret)
    this.render()
  }

  /** Gives the view up to the editor of another document. Nothing of this one runs until it is activated again. */
  deactivate(): void {
    if (!this.active) return
    this.active = false
    this.composing = false
    this.drag = null
    this.view.onPointer = () => {}
    this.view.input.onkeydown = null
    this.view.input.oninput = null
    this.view.input.removeEventListener('compositionstart', this.compositionStart)
    this.view.input.removeEventListener('compositionend', this.compositionEnd)
    this.view.input.parentElement?.classList.remove('composing')
  }

  /** Resolves once every edit asked for so far is done and drawn. */
  idle(): Promise<void> {
    return this.queue
  }

  /**
   * Stops the editor for good, when its document is closed. The document is freed right
   * after, so nothing of it may run again: not a queued edit, not an input method.
   */
  dispose(): void {
    this.deactivate()
    this.disposed = true
  }

  private compositionStart = (): void => {
    this.composing = true
    this.view.input.parentElement?.classList.add('composing')
  }

  private compositionEnd = (event: CompositionEvent): void => {
    this.composing = false
    this.view.input.parentElement?.classList.remove('composing')
    this.takeInput(event.data || this.view.input.value)
  }

  get position(): Caret | null {
    return this.caret
  }

  /** The selected picture, or null when none is. */
  get picture(): Picked | null {
    return this.picked
  }

  private pointer(page: number, x: number, y: number, info: PointerInfo): void {
    if (info.phase === 'up') {
      this.release()
      if (this.painter) this.paint()
      return
    }
    if (info.phase === 'move' && this.drag) {
      this.dragTo(page, x, y)
      return
    }
    if (info.phase === 'down') {
      this.drag = null
      const framed = info.corner && this.picked ? this.placeOf(this.picked) : null
      const image = this.imagesOf(page).find(
        (candidate) =>
          x >= candidate.x && x <= candidate.x + candidate.width && y >= candidate.y && y <= candidate.y + candidate.height,
      )
      if (info.corner && framed?.page === page) {
        const { width, height } = framed.image
        this.drag = { kind: 'size', page, x, y, corner: info.corner, image: framed.image, width, height }
      } else if (image) {
        // The caret waits right after the picture: what is typed next goes there.
        this.collapse({ para: image.para, offset: image.offset })
        this.pending = {}
        this.goalX = null
        this.picked = { para: image.para, index: this.indexOf(page, image) }
        this.drag = { kind: 'move', page, x, y, moving: false, target: null }
        this.render()
      }
      if (this.drag) {
        this.view.input.focus({ preventScroll: true })
        return
      }
    }
    const hit = hitPage(this.linesOf(page), x, y)
    if (!hit) return
    if (info.phase === 'down' && info.clicks === 2) this.selectWord(hit)
    else if (info.phase === 'down' && info.clicks >= 3) this.selectParagraph(hit.para)
    else this.moveTo(hit, info.phase === 'move' || info.shift)
    this.view.input.focus({ preventScroll: true })
  }

  /** Follows the pointer with the picture's new size, or with the place it would move to. */
  private dragTo(page: number, x: number, y: number): void {
    const drag = this.drag
    if (!drag) return
    const lines = this.linesOf(page)
    const data = this.view.pageData(page)
    if (!data) return
    // The text starts at the left margin; the right one is taken to be as wide.
    const margin = Math.min(72, ...lines.map((line) => line.runs[0].edges[0]))
    if (drag.kind === 'size') {
      if (page !== drag.page) return
      const { image, corner } = drag
      const wider = (x - drag.x) * (corner.endsWith('e') ? 1 : -1)
      const taller = (y - drag.y) * (corner.startsWith('s') ? 1 : -1)
      // The picture keeps its proportions; the side dragged further decides.
      const least = MIN_PICTURE / Math.min(image.width, image.height)
      const most = Math.max(1, (data.w - 2 * margin) / image.width)
      const scale = Math.min(most, Math.max(least, (image.width + wider) / image.width, (image.height + taller) / image.height))
      drag.width = image.width * scale
      drag.height = image.height * scale
      this.view.showPicture({
        page,
        x: corner.endsWith('w') ? image.x + image.width - drag.width : image.x,
        y: corner.startsWith('n') ? image.y + image.height - drag.height : image.y,
        width: drag.width,
        height: drag.height,
      })
      return
    }
    if (!drag.moving && (page !== drag.page || Math.hypot(x - drag.x, y - drag.y) > DRAG_SLACK)) drag.moving = true
    if (!drag.moving) return
    const hit = hitPage(lines, x, y)
    // A picture goes between blocks: before or after a paragraph, or a whole table.
    const [first, last] = hit ? this.blockOf(hit.para) : [0, 0]
    const extent = hit && extentOf(lines, first, last)
    if (!hit || !extent) {
      drag.target = null
      this.view.showDrop(null)
      return
    }
    const after = y > (extent.top + extent.bottom) / 2
    drag.target = { para: hit.para, after }
    this.view.showDrop({ page, x: margin, y: after ? extent.bottom + 2 : extent.top - 2, width: data.w - 2 * margin })
  }

  /** Ends a drag: the picture takes the size, or the place, the pointer left it at. */
  private release(): void {
    const drag = this.drag
    const picked = this.picked
    this.drag = null
    this.view.showDrop(null)
    if (!drag || !picked) return
    if (drag.kind === 'size') {
      if (Math.abs(drag.width - drag.image.width) < 0.5) {
        this.render()
        return
      }
      void this.edit(() => {
        this.doc.resizeImage(picked.para, picked.index, drag.width, drag.height)
        this.picked = picked
      })
    } else if (drag.moving && drag.target) {
      const { para, after } = drag.target
      void this.edit(() => {
        const landed = this.doc.moveImage(picked.para, picked.index, para, after)
        this.picked = { para: landed, index: 0 }
        this.collapse({ para: landed, offset: 0 })
      })
    }
  }

  /** The first and the last paragraph of the top-level block a paragraph is in: itself, or its table. */
  private blockOf(para: number): [number, number] {
    const paras = this.structure.paras
    if (!paras[para]?.t) return [para, para]
    let first = para
    let last = para
    while (paras[first - 1]?.t) first--
    while (paras[last + 1]?.t) last++
    return [first, last]
  }

  private imagesOf(page: number): PageImage[] {
    let images = this.images.get(page)
    if (!images) {
      const data = this.view.pageData(page)
      images = data ? pageImages(data, this.linesOf(page)) : []
      this.images.set(page, images)
    }
    return images
  }

  /** Which of its paragraph's pictures a picture on a page is, counting the ones on earlier pages. */
  private indexOf(page: number, image: PageImage): number {
    let index = image.order
    for (const earlier of this.structure.paras[image.para]?.pages ?? []) {
      if (earlier - 1 < page) index += this.imagesOf(earlier - 1).filter((other) => other.para === image.para).length
    }
    return index
  }

  /** Where a picture is drawn, or null when the layout has no such picture. */
  private placeOf(picked: Picked): { page: number; image: PageImage } | null {
    let seen = 0
    for (const numbered of this.structure.paras[picked.para]?.pages ?? []) {
      const own = this.imagesOf(numbered - 1).filter((image) => image.para === picked.para)
      if (picked.index < seen + own.length) return { page: numbered - 1, image: own[picked.index - seen] }
      seen += own.length
    }
    return null
  }

  /** The two ends of the selection in document order, or null when nothing is selected. */
  get range(): [Caret, Caret] | null {
    if (!this.anchor || !this.caret || same(this.anchor, this.caret)) return null
    return ordered(this.anchor, this.caret)
  }

  /** Types text over the selection or at the caret; a line feed starts a new paragraph. */
  insert(text: string): Promise<void> {
    // A run of typed characters is one undo step; a space or a new line starts the next.
    const joins = (): boolean =>
      !this.range && !/\s/.test(text) && this.typedAt !== null && same(this.typedAt, this.caret!)
    return this.edit(() => {
      this.deleteSelection()
      const caret = { ...this.caret! }
      const parts = text.replace(/\r\n?/g, '\n').split('\n')
      parts.forEach((part, index) => {
        if (index > 0) {
          this.doc.splitParagraph(caret.para, caret.offset)
          caret.para += 1
          caret.offset = 0
        }
        if (part) {
          const start = caret.offset
          this.doc.insertText(caret.para, start, part)
          caret.offset += Array.from(part).length
          for (const [property, value] of Object.entries(this.pending)) {
            this.doc.formatRange(caret.para, start, caret.offset, property as CharacterProperty, value)
          }
        }
      })
      // The typed text now carries the formats that were waiting for it.
      this.pending = {}
      this.collapse(caret)
      this.typedAt = { ...caret }
    }, joins)
  }

  get canUndo(): boolean {
    return this.doc.canUndo()
  }

  get canRedo(): boolean {
    return this.doc.canRedo()
  }

  undo(): Promise<void> {
    return this.travel(() => this.doc.undo(), this.undone, this.redone, 'before')
  }

  redo(): Promise<void> {
    return this.travel(() => this.doc.redo(), this.redone, this.undone, 'after')
  }

  /** Turns a character property on or off for the selection, the word around the caret, or the text typed next. */
  toggle(property: 'bold' | 'italic' | 'underline' | 'strike'): Promise<void> {
    const state = this.formatState()
    return this.formatCharacters(property, state?.[property] ? '0' : '1')
  }

  /** Makes the text a superscript or a subscript, or puts it back on the line when it already is one. */
  toggleScript(script: Script): Promise<void> {
    return this.formatCharacters('script', this.formatState()?.script === script ? 'none' : script)
  }

  /** Takes every character format off the selection or the word around the caret, and drops the ones that wait for typed text. */
  clearFormat(): Promise<void> {
    this.pending = {}
    const range = this.range ?? (this.caret && this.wordAt(this.caret, true))
    if (!range) {
      this.onState(this.formatState())
      return Promise.resolve()
    }
    return this.edit(() => this.eachStretch(range, (para, start, end) => this.doc.clearFormat(para, start, end)))
  }

  /**
   * Picks the format of the text at the caret up, for the text that is selected next with
   * the pointer; called again before that, it lets the format go.
   */
  togglePainter(): void {
    const state = this.painter ? null : this.formatState()
    this.painter = state && {
      bold: state.bold ? '1' : '0',
      italic: state.italic ? '1' : '0',
      underline: state.underline ? '1' : '0',
      strike: state.strike ? '1' : '0',
      script: state.script ?? 'none',
      highlight: state.highlight ?? 'none',
      // A raised or lowered character is drawn small; its own size is not that one.
      ...(state.size && !state.script ? { size: String(state.size) } : {}),
      ...(state.color ? { color: state.color } : {}),
      ...(state.font ? { font: state.font } : {}),
    }
    this.onState(this.formatState())
  }

  /** Gives the selection, or the word that was clicked, the format the painter holds. */
  private paint(): void {
    const format = this.painter
    this.painter = null
    const range = this.range ?? (this.caret && this.wordAt(this.caret))
    if (!format || !range) {
      this.onState(this.formatState())
      return
    }
    void this.edit(() =>
      this.eachStretch(range, (para, start, end) => {
        // What the text carried goes first, so that it ends up looking like the source.
        this.doc.clearFormat(para, start, end)
        for (const [property, value] of Object.entries(format)) {
          this.doc.formatRange(para, start, end, property as CharacterProperty, value)
        }
      }),
    )
  }

  /** Calls back for the part of each paragraph that lies between two positions. */
  private eachStretch([from, to]: [Caret, Caret], visit: (para: number, start: number, end: number) => void): void {
    for (let para = from.para; para <= to.para; para++) {
      visit(para, para === from.para ? from.offset : 0, para === to.para ? to.offset : this.lengthOf(para))
    }
  }

  /** Sets the first-line indent and the spacing of every paragraph the selection touches. */
  setParagraph(settings: ParagraphSettings): Promise<void> {
    return this.edit(() => {
      for (const para of this.selectedParagraphs()) {
        if (settings.firstLine !== undefined) this.doc.setFirstLineIndent(para, settings.firstLine)
        if (settings.before !== undefined || settings.after !== undefined) {
          const now = this.doc.format(para, 0)
          const before = settings.before === undefined ? now.before : settings.before
          const after = settings.after === undefined ? now.after : settings.after
          this.doc.setParagraphSpacing(para, before ?? -1, after ?? -1)
        }
      }
    })
  }

  /** Inserts an empty table row above or below the caret's row. */
  insertRow(below: boolean): Promise<void> {
    return this.inTable((para) => this.doc.insertTableRow(para, below))
  }

  /** Deletes the caret's table row. */
  deleteRow(): Promise<void> {
    return this.inTable((para) => this.doc.deleteTableRow(para))
  }

  /** Inserts an empty table column left or right of the caret's cell. */
  insertColumn(after: boolean): Promise<void> {
    return this.inTable((para) => this.doc.insertTableColumn(para, after))
  }

  /** Deletes the table column of the caret's cell. */
  deleteColumn(): Promise<void> {
    return this.inTable((para) => this.doc.deleteTableColumn(para))
  }

  /** Merges the table cells the selection runs through into one. */
  mergeCells(): Promise<void> {
    return this.inTable(() => {
      const [from, to] = this.range ?? [this.caret!, this.caret!]
      if (from.para === to.para) throw new Error('先拖动选中要合并的几个单元格')
      if (!this.structure.paras[from.para]?.t || !this.structure.paras[to.para]?.t) {
        throw new Error('选区要都在同一张表格里')
      }
      return this.doc.mergeTableCells(from.para, to.para)
    })
  }

  /** Splits the merged table cell the caret is in. */
  splitCell(): Promise<void> {
    return this.inTable((para) => this.doc.splitTableCell(para))
  }

  /** Runs a table edit at the caret's paragraph and puts the caret where the edit says. */
  private inTable(change: (para: number) => number): Promise<void> {
    return this.edit(() => {
      const para = this.caret!.para
      if (!this.structure.paras[para]?.t) throw new Error('先把光标放进表格里')
      this.collapse({ para: change(para), offset: 0 })
    })
  }

  setSize(points: number): Promise<void> {
    return this.formatCharacters('size', String(points))
  }

  setFont(family: string): Promise<void> {
    return this.formatCharacters('font', family)
  }

  /** Sets the text colour: six hex digits, or `auto` for the document's own. */
  setColor(color: string): Promise<void> {
    return this.formatCharacters('color', color)
  }

  /** Sets the highlight: a Word highlight colour name, or `none`. */
  setHighlight(color: string): Promise<void> {
    return this.formatCharacters('highlight', color)
  }

  /** Aligns every paragraph the selection touches. */
  setAlignment(alignment: Alignment): Promise<void> {
    return this.edit(() => {
      for (const para of this.selectedParagraphs()) this.doc.setAlignment(para, alignment)
    })
  }

  /** Makes every paragraph the selection touches a heading of a level, or body text for 0. */
  setHeading(level: number): Promise<void> {
    return this.edit(() => {
      for (const para of this.selectedParagraphs()) this.doc.setHeading(para, level)
    })
  }

  /** Makes the paragraphs the selection touches a list of a kind, or plain again when they already are. */
  toggleList(kind: ListKind): Promise<void> {
    return this.edit(() => {
      const paras = this.selectedParagraphs()
      const target = this.doc.format(paras[0], 0).list === kind ? 'none' : kind
      for (const para of paras) this.doc.setList(para, target)
    })
  }

  /** Moves the left indent of the paragraphs the selection touches by a number of points. */
  changeIndent(points: number): Promise<void> {
    return this.edit(() => {
      for (const para of this.selectedParagraphs()) {
        this.doc.setIndent(para, Math.max(0, this.doc.format(para, 0).indent + points))
      }
    })
  }

  setLineSpacing(multiple: number): Promise<void> {
    return this.edit(() => {
      for (const para of this.selectedParagraphs()) this.doc.setLineSpacing(para, multiple)
    })
  }

  /** Ends the page at the caret: what follows starts a new paragraph on the next page. */
  insertPageBreak(): Promise<void> {
    return this.edit(() => {
      if (this.structure.paras[this.caret!.para]?.t) throw new Error('表格里不能插入分页符')
      this.deleteSelection()
      const { para, offset } = this.caret!
      this.doc.splitParagraph(para, offset)
      this.doc.setPageBreakBefore(para + 1, true)
      this.collapse({ para: para + 1, offset: 0 })
    })
  }

  /** Inserts a table after the caret's paragraph and puts the caret in its first cell. */
  insertTable(rows: number, columns: number): Promise<void> {
    return this.edit(() => {
      const para = this.caret!.para
      if (this.structure.paras[para]?.t) throw new Error('表格里暂不支持再插入表格')
      this.doc.insertTable(para, rows, columns)
      this.collapse({ para: para + 1, offset: 0 })
    })
  }

  /**
   * Inserts a picture in a paragraph of its own after the caret's paragraph, or after the
   * table the caret is in. The size is in points.
   */
  insertImage(data: Uint8Array, filename: string, width: number, height: number): Promise<void> {
    return this.edit(() => this.doc.insertImage(this.caret!.para, data, filename, width, height))
  }

  /** Puts the caret at the start or the end of the document. */
  moveToEdge(end: boolean, extend = false): void {
    const last = this.structure.paras.length - 1
    if (last < 0) return
    this.moveTo(end ? { para: last, offset: this.lengthOf(last) } : { para: 0, offset: 0 }, extend)
  }

  /**
   * Selects the next occurrence of a text after the selection, or the one before it, going
   * round at the ends of the document. Case is ignored.
   */
  find(query: string, backwards = false): FindResult {
    const matches = this.matches(query)
    if (matches.length === 0 || !this.caret) return { index: 0, total: matches.length }
    const [from, to] = this.range ?? [this.caret, this.caret]
    const after = (match: Match, position: Caret): boolean =>
      match.para > position.para || (match.para === position.para && match.start >= position.offset)
    const before = (match: Match, position: Caret): boolean =>
      match.para < position.para || (match.para === position.para && match.end <= position.offset)
    let index = backwards ? matches.findLastIndex((match) => before(match, from)) : matches.findIndex((match) => after(match, to))
    if (index < 0) index = backwards ? matches.length - 1 : 0
    const match = matches[index]
    this.anchor = { para: match.para, offset: match.start }
    this.caret = { para: match.para, offset: match.end }
    this.picked = null
    this.pending = {}
    this.goalX = null
    this.render()
    return { index: index + 1, total: matches.length }
  }

  /** How many times a text occurs. */
  count(query: string): number {
    return this.matches(query).length
  }

  /** Replaces the selected occurrence, when the selection is one, and selects the next. */
  async replace(query: string, replacement: string): Promise<FindResult> {
    const range = this.range
    if (range && range[0].para === range[1].para && this.selectedText().toLowerCase() === query.toLowerCase()) {
      await this.insert(replacement)
    }
    return this.find(query)
  }

  /** Replaces every occurrence of a text and resolves to how many there were. */
  async replaceAll(query: string, replacement: string): Promise<number> {
    let count = 0
    await this.edit(() => {
      const matches = this.matches(query)
      count = matches.length
      // From the last to the first, so that no replacement moves one still to come.
      for (const match of matches.reverse()) {
        this.doc.deleteRange(match.para, match.start, match.end)
        if (replacement) this.doc.insertText(match.para, match.start, replacement)
      }
      if (count > 0) this.collapse({ para: this.caret!.para, offset: 0 })
    })
    return count
  }

  selectAll(): void {
    const last = this.structure.paras.length - 1
    if (last < 0) return
    this.anchor = { para: 0, offset: 0 }
    this.caret = { para: last, offset: this.lengthOf(last) }
    this.picked = null
    this.render()
  }

  /** The selected text, paragraphs separated by line feeds. */
  selectedText(): string {
    const range = this.range
    if (!range) return ''
    const [from, to] = range
    const parts: string[] = []
    for (let para = from.para; para <= to.para; para++) {
      const characters = Array.from(this.textOf(para))
      parts.push(
        characters
          .slice(para === from.para ? from.offset : 0, para === to.para ? to.offset : characters.length)
          .join(''),
      )
    }
    return parts.join('\n')
  }

  /**
   * What the toolbar shows: the format at the start of the selection, or before the caret,
   * with the formats that wait for the next typed text on top.
   */
  formatState(): FormatState | null {
    if (!this.caret) return null
    const range = this.range
    const probe = range ? range[0] : { para: this.caret.para, offset: Math.max(0, this.caret.offset - 1) }
    const drawn = styleAt(probe, this.structure, (page) => this.linesOf(page))
    const own = this.doc.format(probe.para, probe.offset)
    const waiting = this.pending
    const flag = (property: 'bold' | 'italic' | 'underline', value: boolean): boolean =>
      waiting[property] === undefined ? value : waiting[property] === '1'
    return {
      bold: flag('bold', drawn?.bold ?? false),
      italic: flag('italic', drawn?.italic ?? false),
      underline: flag('underline', own.underline),
      strike: waiting.strike === undefined ? own.strike : waiting.strike === '1',
      script: waiting.script === undefined ? own.script : waiting.script === 'none' ? null : (waiting.script as Script),
      highlight: waiting.highlight === undefined ? own.highlight : waiting.highlight === 'none' ? null : waiting.highlight,
      size: waiting.size ? Number(waiting.size) : (drawn?.size ?? null),
      font: waiting.font ?? own.font,
      color: waiting.color && waiting.color !== 'auto' ? waiting.color : drawn ? hex(drawn.color) : null,
      align: own.align,
      heading: own.heading,
      list: own.list,
      indent: own.indent,
      spacing: own.spacing,
      firstLine: own.firstLine,
      before: own.before,
      after: own.after,
      inTable: Boolean(this.structure.paras[this.caret.para]?.t),
      painting: this.painter !== null,
    }
  }

  private takeInput(text: string): void {
    this.view.input.value = ''
    if (text) void this.insert(text)
  }

  private keyDown(event: KeyboardEvent): void {
    if (event.isComposing || this.composing || event.altKey) return
    const picked = this.picked
    if (picked && !event.ctrlKey && !event.metaKey) {
      if (event.key === 'Backspace' || event.key === 'Delete') {
        event.preventDefault()
        // The caret already waits where the picture was.
        void this.edit(() => this.doc.deleteImage(picked.para, picked.index))
        return
      }
      if (event.key === 'Escape') {
        event.preventDefault()
        this.picked = null
        this.render()
        return
      }
    }
    if (this.painter && event.key === 'Escape') {
      event.preventDefault()
      this.painter = null
      this.onState(this.formatState())
      return
    }
    const extend = event.shiftKey
    const handlers: Record<string, () => void> =
      event.ctrlKey || event.metaKey
        ? {
            a: () => this.selectAll(),
            b: () => void this.toggle('bold'),
            i: () => void this.toggle('italic'),
            u: () => void this.toggle('underline'),
            c: () => this.copy(),
            x: () => {
              this.copy()
              void this.edit(() => this.deleteSelection())
            },
            z: () => void (extend ? this.redo() : this.undo()),
            y: () => void this.redo(),
            enter: () => void this.insertPageBreak(),
            arrowleft: () => this.moveByWord(-1, extend),
            arrowright: () => this.moveByWord(1, extend),
            home: () => this.moveToEdge(false, extend),
            end: () => this.moveToEdge(true, extend),
          }
        : {
            Backspace: () => void this.edit(() => this.deleteSelection() || this.deleteBackward()),
            Delete: () => void this.edit(() => this.deleteSelection() || this.deleteForward()),
            Enter: () => this.enter(),
            Tab: () => this.tab(extend),
            ArrowLeft: () => this.moveBy(-1, extend),
            ArrowRight: () => this.moveBy(1, extend),
            ArrowUp: () => this.moveLines(-1, extend),
            ArrowDown: () => this.moveLines(1, extend),
            PageUp: () => this.moveLines(-this.linesPerScreen(), extend),
            PageDown: () => this.moveLines(this.linesPerScreen(), extend),
            Home: () => this.moveToLineEdge(false, extend),
            End: () => this.moveToLineEdge(true, extend),
          }
    const handler = handlers[event.ctrlKey || event.metaKey ? event.key.toLowerCase() : event.key]
    if (!handler) return
    event.preventDefault()
    handler()
  }

  /** Enter starts a paragraph, and in an empty list item it ends the list instead. */
  private enter(): void {
    const caret = this.caret
    const info = caret && this.structure.paras[caret.para]
    if (caret && info?.l && info.len === 0 && !this.range) {
      void this.edit(() => this.doc.setList(caret.para, 'none'))
    } else {
      void this.insert('\n')
    }
  }

  /** In a table Tab goes from cell to cell; elsewhere it types a tab. */
  private tab(backwards: boolean): void {
    if (this.caret && this.structure.paras[this.caret.para]?.t) this.moveByCell(backwards ? -1 : 1)
    else if (!backwards) void this.insert('\t')
  }

  private copy(): void {
    const text = this.selectedText()
    if (text) void navigator.clipboard.writeText(text).catch(() => this.onMessage('无法写入剪贴板'))
  }

  /** Runs one edit, then lays out, redraws and places the caret and the selection again. */
  private edit(change: () => void, joins: () => boolean = () => false): Promise<void> {
    return this.enqueue(async () => {
      if (!this.caret || !this.anchor) return
      const before = this.ends()
      const join = joins()
      // An edit lets go of the selected picture, unless the edit is to the picture.
      const picked = this.picked
      this.picked = null
      this.typedAt = null
      this.goalX = null
      this.onMessage('')
      this.doc.beginEdit()
      try {
        change()
      } catch (error) {
        this.anchor = before.anchor
        this.caret = before.caret
        this.picked = picked
        this.onMessage(error instanceof Error ? error.message : String(error))
      }
      // Even a refused edit may have changed part of the document; what it changed is a step too.
      const recorded = this.doc.commitEdit(join)
      if (recorded === 1) {
        this.undone.push({ before, after: this.ends() })
        if (this.undone.length > UNDO_LIMIT) this.undone.shift()
      } else if (recorded === 2 && this.undone.length > 0) {
        this.undone[this.undone.length - 1].after = this.ends()
      }
      if (recorded > 0) this.redone = []
      await this.redraw()
      if (recorded > 0) this.onChange()
    })
  }

  /**
   * Runs a task after the ones before it. A task that fails says so and lets the next one
   * run: a chain of promises that rejected once would otherwise skip every later edit.
   */
  private enqueue(task: () => Promise<void>): Promise<void> {
    this.queue = this.queue
      .then(() => (this.disposed || !this.active ? undefined : task()))
      .catch((error: unknown) => {
        this.onMessage(`出错了：${error instanceof Error ? error.message : String(error)}`)
      })
    return this.queue
  }

  /** Steps through the undo history and puts the selection back where that step had it. */
  private travel(apply: () => boolean, from: Step[], to: Step[], side: 'before' | 'after'): Promise<void> {
    return this.enqueue(async () => {
      if (!apply()) return
      this.picked = null
      this.typedAt = null
      this.pending = {}
      this.goalX = null
      this.onMessage('')
      const step = from.pop()
      if (step) {
        to.push(step)
        this.anchor = { ...step[side].anchor }
        this.caret = { ...step[side].caret }
      }
      await this.redraw()
      this.onChange()
    })
  }

  private ends(): Ends {
    return { anchor: { ...this.anchor! }, caret: { ...this.caret! } }
  }

  /** Lays out, redraws, and places the caret and the selection from the new layout. */
  private async redraw(): Promise<void> {
    await this.view.refresh()
    this.structure = this.view.structure
    this.lines.clear()
    this.images.clear()
    this.texts.clear()
    if (this.anchor) this.anchor = this.valid(this.anchor)
    if (this.caret) this.caret = this.valid(this.caret)
    this.render()
  }

  /** The position itself when the document still has it, or else the closest one it has. */
  private valid(position: Caret): Caret {
    const last = this.structure.paras.length - 1
    if (last < 0) return position
    const para = Math.min(Math.max(position.para, 0), last)
    return { para, offset: para === position.para ? Math.min(position.offset, this.lengthOf(para)) : 0 }
  }

  private formatCharacters(property: CharacterProperty, value: string): Promise<void> {
    // With nothing selected the format goes to the word the caret is inside, as in Word,
    // and when the caret is in no word it waits for the text typed next.
    if (this.caret && !this.range && !this.wordAt(this.caret, true)) {
      this.pending[property] = value
      this.onState(this.formatState())
      return Promise.resolve()
    }
    return this.edit(() => {
      const range = this.range ?? this.wordAt(this.caret!, true)
      if (!range) return
      const [from, to] = range
      for (let para = from.para; para <= to.para; para++) {
        const start = para === from.para ? from.offset : 0
        const end = para === to.para ? to.offset : this.lengthOf(para)
        this.doc.formatRange(para, start, end, property, value)
      }
    })
  }

  /** The paragraphs between the two ends of the selection, or the caret's own. */
  private selectedParagraphs(): number[] {
    const [from, to] = this.range ?? [this.caret!, this.caret!]
    const paras: number[] = []
    for (let para = from.para; para <= to.para; para++) paras.push(para)
    return paras
  }

  /** Deletes the selected text and leaves the caret where it was. Returns whether there was any. */
  private deleteSelection(): boolean {
    const range = this.range
    if (!range) return false
    const [from, to] = range
    this.collapse(from)
    if (from.para === to.para) {
      this.doc.deleteRange(from.para, from.offset, to.offset)
      return true
    }
    this.doc.deleteRange(to.para, 0, to.offset)
    this.doc.deleteRange(from.para, from.offset, this.lengthOf(from.para))
    // Sibling paragraphs collapse into one. Across table cells, or around a table, each
    // paragraph loses its selected text and stays, so that no cell and no table goes.
    let siblings = true
    for (let para = from.para; para < to.para; para++) siblings &&= this.structure.paras[para].m === 1
    for (let para = to.para - 1; para > from.para; para--) {
      if (siblings) this.doc.removeParagraph(para)
      else this.doc.deleteRange(para, 0, this.lengthOf(para))
    }
    if (siblings) {
      try {
        this.doc.mergeParagraphs(from.para)
      } catch (error) {
        // The text is gone either way; only the two paragraphs stay apart.
        this.onMessage(error instanceof Error ? error.message : String(error))
      }
    }
    return true
  }

  /** The picture the caret is drawn right after, when it is. */
  private pictureBefore(caret: Caret): Picked | null {
    const rect = caretRect(caret, this.structure, (page) => this.linesOf(page))
    if (!rect) return null
    const image = this.imagesOf(rect.page).findLast(
      (candidate) =>
        candidate.para === caret.para &&
        candidate.offset === caret.offset &&
        Math.abs(candidate.x + candidate.width - rect.x) < 1,
    )
    return image ? { para: image.para, index: this.indexOf(rect.page, image) } : null
  }

  private deleteBackward(): void {
    const caret = this.caret!
    const picture = this.pictureBefore(caret)
    if (picture) {
      // A picture counts for no character, yet Backspace right after one takes it out.
      this.doc.deleteImage(picture.para, picture.index)
    } else if (caret.offset > 0) {
      this.doc.deleteRange(caret.para, caret.offset - 1, caret.offset)
      this.collapse({ para: caret.para, offset: caret.offset - 1 })
    } else if (this.structure.paras[caret.para]?.l) {
      // At the start of a list item, Backspace first takes the item out of its list.
      this.doc.setList(caret.para, 'none')
    } else if (this.structure.paras[caret.para - 1]?.m) {
      const offset = this.doc.mergeParagraphs(caret.para - 1)
      this.collapse({ para: caret.para - 1, offset })
    }
  }

  private deleteForward(): void {
    const caret = this.caret!
    if (caret.offset < this.lengthOf(caret.para)) {
      this.doc.deleteRange(caret.para, caret.offset, caret.offset + 1)
    } else if (this.structure.paras[caret.para]?.m) {
      this.doc.mergeParagraphs(caret.para)
    }
  }

  private lengthOf(para: number): number {
    return this.structure.paras[para]?.len ?? 0
  }

  private textOf(para: number): string {
    let text = this.texts.get(para)
    if (text === undefined) {
      text = this.doc.paragraphText(para)
      this.texts.set(para, text)
    }
    return text
  }

  private linesOf(page: number): Line[] {
    let lines = this.lines.get(page)
    if (!lines) {
      const data = this.view.pageData(page)
      lines = data ? pageLines(data, this.structure) : []
      this.lines.set(page, lines)
    }
    return lines
  }

  /** Every place a text occurs, in document order. Case is ignored. */
  private matches(query: string): Match[] {
    const needle = query.toLowerCase()
    if (!needle) return []
    const length = Array.from(needle).length
    const found: Match[] = []
    for (let para = 0; para < this.structure.paras.length; para++) {
      if (this.lengthOf(para) === 0) continue
      const text = this.textOf(para)
      const haystack = text.toLowerCase()
      for (let at = haystack.indexOf(needle); at >= 0; at = haystack.indexOf(needle, at + needle.length)) {
        // Offsets count characters, and a string index counts UTF-16 units.
        const start = Array.from(text.slice(0, at)).length
        found.push({ para, start, end: start + length })
      }
    }
    return found
  }

  /** Sets both ends to one position without drawing; an edit draws once it is done. */
  private collapse(caret: Caret): void {
    this.anchor = { ...caret }
    this.caret = { ...caret }
  }

  /** Moves the caret; the other end of the selection follows unless the move extends it. */
  private moveTo(caret: Caret, extend = false, keepGoal = false): void {
    this.caret = caret
    this.picked = null
    if (!extend || !this.anchor) this.anchor = { ...caret }
    // A format chosen for the next typed text belongs to the place it was chosen at.
    this.pending = {}
    if (!keepGoal) this.goalX = null
    this.render()
  }

  /** Draws the caret and the selection from the current layout. */
  private render(): void {
    const linesOf = (page: number): Line[] => this.linesOf(page)
    const rect = this.caret && caretRect(this.caret, this.structure, linesOf)
    const range = this.range
    this.view.showSelection(range ? selectionRects(range[0], range[1], this.structure, linesOf) : [])
    if (rect) this.view.showCaret(rect.page, rect.x, rect.top, rect.height)
    else this.view.hideCaret()
    const place = this.picked && this.placeOf(this.picked)
    if (!place) this.picked = null
    this.view.showPicture(
      place ? { page: place.page, x: place.image.x, y: place.image.y, width: place.image.width, height: place.image.height } : null,
    )
    this.onState(this.formatState())
  }

  /** The words of a paragraph as [start, end) character ranges, in order. */
  private wordsOf(para: number): [number, number][] {
    const text = this.textOf(para)
    const words: [number, number][] = []
    if (segmenter) {
      let offset = 0
      for (const segment of segmenter.segment(text)) {
        const length = Array.from(segment.segment).length
        if (segment.isWordLike) words.push([offset, offset + length])
        offset += length
      }
      return words
    }
    const characters = Array.from(text)
    let start = -1
    characters.forEach((character, index) => {
      const isWord = WORD_CHARACTER.test(character)
      if (isWord && start < 0) start = index
      if (!isWord && start >= 0) {
        words.push([start, index])
        start = -1
      }
    })
    if (start >= 0) words.push([start, characters.length])
    return words
  }

  /** The word a position touches, or with `inside` the word it is strictly within; null when there is none. */
  private wordAt(position: Caret, inside = false): [Caret, Caret] | null {
    const word = this.wordsOf(position.para).find(([start, end]) =>
      inside ? position.offset > start && position.offset < end : position.offset >= start && position.offset <= end,
    )
    return word ? [{ para: position.para, offset: word[0] }, { para: position.para, offset: word[1] }] : null
  }

  private selectWord(position: Caret): void {
    const word = this.wordAt(position)
    if (!word) {
      this.moveTo(position)
      return
    }
    this.anchor = word[0]
    this.caret = word[1]
    this.picked = null
    this.render()
  }

  private selectParagraph(para: number): void {
    this.anchor = { para, offset: 0 }
    this.caret = { para, offset: this.lengthOf(para) }
    this.picked = null
    this.render()
  }

  /** Moves one character, crossing into the neighbouring paragraph at an end. */
  private moveBy(step: 1 | -1, extend: boolean): void {
    const caret = this.caret
    if (!caret) return
    const range = this.range
    if (range && !extend) {
      // A plain arrow key leaves a selection at the end it points to.
      this.moveTo(step < 0 ? range[0] : range[1])
      return
    }
    const offset = caret.offset + step
    if (offset >= 0 && offset <= this.lengthOf(caret.para)) {
      this.moveTo({ para: caret.para, offset }, extend)
      return
    }
    const para = caret.para + step
    if (para < 0 || para >= this.structure.paras.length) return
    this.moveTo({ para, offset: step > 0 ? 0 : this.lengthOf(para) }, extend)
  }

  /** Moves to the end of the next word, or the start of the one before. */
  private moveByWord(step: 1 | -1, extend: boolean): void {
    const caret = this.caret
    if (!caret) return
    const words = this.wordsOf(caret.para)
    if (step > 0) {
      const next = words.find(([, end]) => end > caret.offset)
      if (next) this.moveTo({ para: caret.para, offset: next[1] }, extend)
      else if (caret.offset < this.lengthOf(caret.para)) this.moveTo({ para: caret.para, offset: this.lengthOf(caret.para) }, extend)
      else this.moveBy(1, extend)
    } else {
      const previous = words.findLast(([start]) => start < caret.offset)
      if (previous) this.moveTo({ para: caret.para, offset: previous[0] }, extend)
      else if (caret.offset > 0) this.moveTo({ para: caret.para, offset: 0 }, extend)
      else this.moveBy(-1, extend)
    }
  }

  /** In a table, moves to the next or the previous paragraph: the next cell, for cells of one paragraph. */
  private moveByCell(step: 1 | -1): void {
    const caret = this.caret
    if (!caret || !this.structure.paras[caret.para]?.t) return
    const para = caret.para + step
    if (!this.structure.paras[para]?.t) return
    this.moveTo({ para, offset: 0 })
  }

  /** About how many lines of the caret's size fit the window, less one that stays in view. */
  private linesPerScreen(): number {
    const rect = this.caret && caretRect(this.caret, this.structure, (page) => this.linesOf(page))
    return rect ? Math.max(1, Math.floor(this.view.viewportHeight / (rect.height * 1.3)) - 1) : 1
  }

  /** Moves a number of lines down, or up for a negative number, keeping the horizontal position. */
  private moveLines(count: number, extend: boolean): void {
    const linesOf = (page: number): Line[] => this.linesOf(page)
    let caret = this.caret
    let rect = caret && caretRect(caret, this.structure, linesOf)
    if (!caret || !rect) return
    const x = this.goalX ?? rect.x
    const step = count > 0 ? 1 : -1
    for (let moved = 0; moved < Math.abs(count); moved++) {
      let page: number = rect.page
      let index: number = rect.line + step
      while (index < 0 || index >= linesOf(page).length) {
        page += step
        if (page < 0 || page >= this.view.pageCount) break
        index = step > 0 ? 0 : linesOf(page).length - 1
      }
      const line = linesOf(page)?.[index]
      if (page < 0 || page >= this.view.pageCount || !line) break
      const next: Caret = hitLine(line, x)
      const nextRect = caretRect(next, this.structure, linesOf)
      if (!nextRect) break
      caret = next
      rect = nextRect
    }
    this.goalX = x
    this.moveTo(caret, extend, true)
  }

  private moveToLineEdge(end: boolean, extend: boolean): void {
    const rect = this.caret && caretRect(this.caret, this.structure, (page) => this.linesOf(page))
    if (!rect) return
    const line = this.linesOf(rect.page)[rect.line]
    // A line in a table runs across the row: stay within the caret's own paragraph.
    const own = line.runs.filter((run) => run.para === this.caret!.para)
    this.moveTo(hitLine({ ...line, runs: own.length > 0 ? own : line.runs }, end ? 1e9 : -1e9), extend)
  }
}
