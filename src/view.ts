// The page view: a scrolling column of pages. Each page is a placeholder of the right size;
// a canvas is created and drawn only while the page is near the viewport, so a long document
// costs memory for the few pages on screen, not for all of them.

import type { Doc, PageData, Structure } from './engine.ts'
import { type PageFont, registerFonts } from './fonts.ts'
import { foreignLines, paragraphBoxes } from './geometry.ts'
import { decodeImage } from './images.ts'
import { imageIds, renderPage } from './render.ts'

const CSS_PX_PER_POINT = 96 / 72
/** How far outside the viewport, in CSS pixels, a page keeps its canvas. */
const KEEP_MARGIN = 800
export const MIN_ZOOM = 0.25
export const MAX_ZOOM = 4

interface PageSlot {
  holder: HTMLDivElement
  size: [number, number]
  canvas: HTMLCanvasElement | null
  /** The page's text as elements, for what reads the element tree and not the pixels. */
  text: HTMLDivElement | null
  data: PageData | null
  /** The engine's text for `data`, kept to tell whether an edit changed the page. */
  json: string | null
}

/** A corner of the frame around the selected picture. */
export type Corner = 'nw' | 'ne' | 'sw' | 'se'

export interface PointerInfo {
  /**
   * down: the button went down on a page; move: the pointer moved with the button held;
   * up: the button came up, on page -1 when that was outside every page.
   */
  phase: 'down' | 'move' | 'up'
  shift: boolean
  /** 1 for a click, 2 for a double click, as the system counts them. */
  clicks: number
  /** The handle of the picture frame the button went down on, when it did. */
  corner?: Corner
}

interface PlacedRect {
  page: number
  x: number
  y: number
  width: number
  height: number
}

interface CaretPlace {
  page: number
  x: number
  top: number
  height: number
}

export class PageView {
  /** Called when the page in the middle of the viewport, or the page count, changes. */
  onPageChange: (current: number, total: number) => void = () => {}
  /** Called with a point on a page, in points, when the pointer goes down on it or drags over it. */
  onPointer: (page: number, x: number, y: number, info: PointerInfo) => void = () => {}

  /** Receives the keyboard and the input method; it sits at the caret so candidate windows open there. */
  readonly input: HTMLTextAreaElement

  private scroller: HTMLElement
  private doc: Doc | null = null
  private layout: Structure = { nodes: [], paras: [] }
  private zoomLevel = 1
  private slots: PageSlot[] = []
  private fonts = new Map<number, PageFont>()
  /** Decoded pictures by media id; null while one is still being decoded. */
  private images = new Map<string, HTMLImageElement | null>()
  /** Bumped whenever the document is replaced, so late image loads are dropped. */
  private generation = 0
  private currentPage = -1
  private syncQueued = false
  private caretAnchor: HTMLDivElement
  private caret: CaretPlace | null = null
  private selection: PlacedRect[] = []
  private selectionMarks: HTMLDivElement[] = []
  /** The frame around the selected picture, with a handle on each corner. */
  private frame: HTMLDivElement
  private picture: PlacedRect | null = null
  /** The line that shows where a dragged picture would land. */
  private dropMark: HTMLDivElement
  private drop: { page: number; x: number; y: number; width: number } | null = null
  private unlisten: () => void = () => {}

  constructor(scroller: HTMLElement) {
    this.scroller = scroller

    this.frame = document.createElement('div')
    this.frame.className = 'picture-frame'
    for (const corner of ['nw', 'ne', 'sw', 'se'] satisfies Corner[]) {
      const handle = document.createElement('i')
      handle.dataset.corner = corner
      this.frame.append(handle)
    }
    this.dropMark = document.createElement('div')
    this.dropMark.className = 'drop-line'

    this.caretAnchor = document.createElement('div')
    this.caretAnchor.className = 'caret'
    const bar = document.createElement('div')
    bar.className = 'caret-bar'
    this.input = document.createElement('textarea')
    this.input.setAttribute('aria-label', '文档输入')
    this.input.autocapitalize = 'off'
    this.input.spellcheck = false
    this.caretAnchor.append(bar, this.input)

    // acdert fires no scroll event for a scrollTop set from script, so every programmatic
    // scroll calls sync() itself; this listener covers the user's own scrolling.
    scroller.addEventListener(
      'scroll',
      () => {
        if (this.syncQueued) return
        this.syncQueued = true
        requestAnimationFrame(() => {
          this.syncQueued = false
          this.sync()
        })
      },
      { passive: true },
    )

    const report = (event: MouseEvent, phase: PointerInfo['phase']): boolean => {
      const target = (phase === 'down' ? event.target : document.elementFromPoint(event.clientX, event.clientY)) as HTMLElement | null
      const holder = target?.closest<HTMLElement>('.page')
      if (!holder) return false
      const bounds = holder.getBoundingClientRect()
      const scale = CSS_PX_PER_POINT * this.zoomLevel
      const corner = phase === 'down' ? (target?.dataset.corner as Corner | undefined) : undefined
      this.onPointer(
        Number(holder.dataset.page),
        (event.clientX - bounds.left) / scale,
        (event.clientY - bounds.top) / scale,
        { phase, shift: event.shiftKey, clicks: event.detail, corner },
      )
      return true
    }
    let dragging = false
    scroller.addEventListener('mousedown', (event) => {
      if (event.button !== 0 || !report(event, 'down')) return
      // Keep the focus on the input instead of letting the click take it.
      event.preventDefault()
      dragging = true
    })
    const onMove = (event: MouseEvent): void => {
      if (dragging) report(event, 'move')
    }
    const onUp = (event: MouseEvent): void => {
      if (!dragging) return
      dragging = false
      if (!report(event, 'up')) this.onPointer(-1, 0, 0, { phase: 'up', shift: event.shiftKey, clicks: event.detail })
    }
    window.addEventListener('mousemove', onMove)
    window.addEventListener('mouseup', onUp)
    this.unlisten = () => {
      window.removeEventListener('mousemove', onMove)
      window.removeEventListener('mouseup', onUp)
    }
  }

  /** Lets go of what the view holds outside its own element, when its document is closed. */
  dispose(): void {
    this.unlisten()
    this.generation++
    this.doc = null
  }

  get zoom(): number {
    return this.zoomLevel
  }

  get pageCount(): number {
    return this.slots.length
  }

  /** The height of the visible part of the document, in points at the current zoom. */
  get viewportHeight(): number {
    return this.scroller.clientHeight / (CSS_PX_PER_POINT * this.zoomLevel)
  }

  /** How far down the document is scrolled, in CSS pixels. */
  get scrollTop(): number {
    return this.scroller.scrollTop
  }

  /**
   * Scrolls the document to a position, in CSS pixels, and reports the page again. An
   * element that was out of the page comes back scrolled to its top, so a view that is
   * shown again is put back where it was with this.
   */
  scrollTo(top: number): void {
    this.scroller.scrollTop = top
    this.sync(true)
  }

  /** What maps the layout on screen back to the document. */
  get structure(): Structure {
    return this.layout
  }

  /** Shows a document from its first page. */
  async show(doc: Doc): Promise<void> {
    this.doc = doc
    this.generation++
    this.images.clear()
    this.caret = null
    this.selection = []
    this.picture = null
    this.drop = null
    const sizes = doc.pageSizes()
    this.layout = doc.structure()
    this.fonts = await registerFonts(doc)
    this.build(sizes)
    this.scroller.scrollTop = 0
    this.sync(true)
  }

  /** Draws the document again after an edit, keeping the scroll position. */
  async refresh(): Promise<void> {
    if (!this.doc) return
    const sizes = this.doc.pageSizes()
    this.layout = this.doc.structure()
    this.fonts = await registerFonts(this.doc)
    const unchanged =
      sizes.length === this.slots.length &&
      sizes.every((size, index) => size[0] === this.slots[index].size[0] && size[1] === this.slots[index].size[1])
    if (unchanged) {
      this.slots.forEach((slot, index) => {
        if (!slot.canvas) {
          slot.data = null
          slot.json = null
          return
        }
        // Most edits change one page. A page whose text from the engine is what it was
        // draws the same, and drawing is the costly part: leave it as it is.
        const json = this.doc!.pageJson(index)
        if (json === slot.json) return
        slot.json = json
        slot.data = JSON.parse(json) as PageData
        this.draw(index)
      })
    } else {
      const top = this.scroller.scrollTop
      this.build(sizes)
      this.scroller.scrollTop = top
    }
    this.sync(true)
  }

  setZoom(zoom: number): void {
    const next = Math.min(MAX_ZOOM, Math.max(MIN_ZOOM, zoom))
    if (next === this.zoomLevel || !this.doc) return
    // Keep the same spot of the document in view across the zoom.
    const anchor = this.scroller.scrollHeight > 0 ? this.scroller.scrollTop / this.scroller.scrollHeight : 0
    this.zoomLevel = next
    this.build(this.slots.map((slot) => slot.size))
    this.scroller.scrollTop = anchor * this.scroller.scrollHeight
    this.sync(true)
  }

  /** The layout data of a page, whether or not it is on screen. */
  pageData(index: number): PageData | null {
    const slot = this.slots[index]
    if (!slot || !this.doc) return null
    if (!slot.data) {
      slot.json = this.doc.pageJson(index)
      slot.data = JSON.parse(slot.json) as PageData
    }
    return slot.data
  }

  /** Puts the caret on a page, at a position in points, and scrolls it into view. */
  showCaret(page: number, x: number, top: number, height: number): void {
    this.caret = { page, x, top, height }
    this.placeCaret()
    const slot = this.slots[page]
    if (!slot) return
    const scale = CSS_PX_PER_POINT * this.zoomLevel
    const caretTop = slot.holder.offsetTop + top * scale
    const caretBottom = caretTop + height * scale
    const margin = 24
    if (caretTop - margin < this.scroller.scrollTop) {
      this.scroller.scrollTop = Math.max(0, caretTop - margin)
      this.sync()
    } else if (caretBottom + margin > this.scroller.scrollTop + this.scroller.clientHeight) {
      this.scroller.scrollTop = caretBottom + margin - this.scroller.clientHeight
      this.sync()
    }
  }

  /** Highlights rectangles on pages, in points; an empty list clears the highlight. */
  showSelection(rects: PlacedRect[]): void {
    this.selection = rects
    this.placeSelection()
  }

  private placeSelection(): void {
    for (const mark of this.selectionMarks) mark.remove()
    const scale = CSS_PX_PER_POINT * this.zoomLevel
    this.selectionMarks = this.selection.flatMap((rect) => {
      const slot = this.slots[rect.page]
      if (!slot) return []
      const mark = document.createElement('div')
      mark.className = 'selection'
      mark.style.left = `${rect.x * scale}px`
      mark.style.top = `${rect.y * scale}px`
      mark.style.width = `${rect.width * scale}px`
      mark.style.height = `${rect.height * scale}px`
      slot.holder.append(mark)
      return [mark]
    })
  }

  /** Frames a picture on a page, in points, with its handles; null takes the frame away. */
  showPicture(rect: PlacedRect | null): void {
    this.picture = rect
    this.placePicture()
  }

  /** Shows where a dragged picture would land: a line across a page, in points; null hides it. */
  showDrop(line: { page: number; x: number; y: number; width: number } | null): void {
    this.drop = line
    this.placePicture()
  }

  private placePicture(): void {
    const scale = CSS_PX_PER_POINT * this.zoomLevel
    const frameHolder = this.picture && this.slots[this.picture.page]?.holder
    // While a picture is selected the caret bar hides; the input stays where it is.
    this.caretAnchor.classList.toggle('picking', Boolean(this.picture && frameHolder))
    if (this.picture && frameHolder) {
      this.frame.style.left = `${this.picture.x * scale}px`
      this.frame.style.top = `${this.picture.y * scale}px`
      this.frame.style.width = `${this.picture.width * scale}px`
      this.frame.style.height = `${this.picture.height * scale}px`
      if (this.frame.parentElement !== frameHolder) frameHolder.append(this.frame)
    } else {
      this.frame.remove()
    }
    const dropHolder = this.drop && this.slots[this.drop.page]?.holder
    if (this.drop && dropHolder) {
      this.dropMark.style.left = `${this.drop.x * scale}px`
      this.dropMark.style.top = `${this.drop.y * scale - 1}px`
      this.dropMark.style.width = `${this.drop.width * scale}px`
      if (this.dropMark.parentElement !== dropHolder) dropHolder.append(this.dropMark)
    } else {
      this.dropMark.remove()
    }
  }

  /**
   * Stops showing the caret, for a position that layout draws nothing for. The input the
   * caret carries stays in the page: taking it out would take the keyboard away with it.
   */
  hideCaret(): void {
    this.caret = null
    this.placeCaret()
  }

  private placeCaret(): void {
    const slot = this.caret ? this.slots[this.caret.page] : undefined
    const holder = slot?.holder ?? (this.caretAnchor.isConnected ? null : this.slots[0]?.holder)
    this.caretAnchor.classList.toggle('parked', !this.caret || !slot)
    if (this.caret && slot) {
      const scale = CSS_PX_PER_POINT * this.zoomLevel
      this.caretAnchor.style.left = `${this.caret.x * scale}px`
      this.caretAnchor.style.top = `${this.caret.top * scale}px`
      this.caretAnchor.style.height = `${this.caret.height * scale}px`
    }
    if (holder && this.caretAnchor.parentElement !== holder) {
      const focused = document.activeElement === this.input
      holder.append(this.caretAnchor)
      // Moving an element in the tree drops its focus.
      if (focused) this.input.focus({ preventScroll: true })
    }
    // Restart the blink so the caret is solid right after it moves.
    this.caretAnchor.classList.remove('blink')
    void this.caretAnchor.offsetWidth
    this.caretAnchor.classList.add('blink')
  }

  private build(sizes: [number, number][]): void {
    const focused = document.activeElement === this.input
    this.slots = sizes.map((size, index) => {
      const holder = document.createElement('div')
      holder.className = 'page'
      holder.dataset.page = String(index)
      holder.setAttribute('role', 'group')
      holder.setAttribute('aria-label', `第 ${index + 1} 页`)
      holder.style.width = `${size[0] * CSS_PX_PER_POINT * this.zoomLevel}px`
      holder.style.height = `${size[1] * CSS_PX_PER_POINT * this.zoomLevel}px`
      return { holder, size, canvas: null, text: null, data: null, json: null }
    })
    this.scroller.replaceChildren(...this.slots.map((slot) => slot.holder))
    this.placeSelection()
    this.placePicture()
    this.placeCaret()
    if (focused) this.input.focus({ preventScroll: true })
  }

  /** Gives a canvas to the pages near the viewport, takes it from the others, and reports the page. */
  private sync(force = false): void {
    const top = this.scroller.scrollTop - KEEP_MARGIN
    const bottom = this.scroller.scrollTop + this.scroller.clientHeight + KEEP_MARGIN
    this.slots.forEach((slot, index) => {
      const start = slot.holder.offsetTop
      const near = start + slot.holder.offsetHeight > top && start < bottom
      if (near && !slot.canvas) this.draw(index)
      else if (!near && slot.canvas) this.release(index)
    })
    this.reportPage(force)
  }

  private draw(index: number): void {
    const slot = this.slots[index]
    const data = this.pageData(index)
    if (!slot || !data) return
    if (!slot.canvas) {
      slot.canvas = document.createElement('canvas')
      slot.holder.prepend(slot.canvas)
    }
    const ratio = window.devicePixelRatio || 1
    const cssWidth = slot.size[0] * CSS_PX_PER_POINT * this.zoomLevel
    const cssHeight = slot.size[1] * CSS_PX_PER_POINT * this.zoomLevel
    slot.canvas.width = Math.round(cssWidth * ratio)
    slot.canvas.height = Math.round(cssHeight * ratio)
    this.loadImages(index, data)
    renderPage(slot.canvas, data, CSS_PX_PER_POINT * this.zoomLevel * ratio, {
      font: (id) => this.fonts.get(id),
      image: (id) => this.images.get(id) ?? undefined,
    })
    this.describe(index, data)
  }

  private release(index: number): void {
    const slot = this.slots[index]
    if (!slot?.canvas) return
    slot.canvas.remove()
    slot.canvas = null
    slot.text?.remove()
    slot.text = null
    slot.data = null
    slot.json = null
  }

  /**
   * Puts the page's text into the element tree, over the place it is drawn. The canvas shows
   * pixels only; this layer is what a screen reader hears and what an agent that reads the
   * screen through acde:inspect sees, and a click on one of its elements lands on the text.
   */
  private describe(index: number, data: PageData): void {
    const slot = this.slots[index]
    if (!slot || !this.doc) return
    slot.text?.remove()
    const layer = document.createElement('div')
    layer.className = 'page-text'
    const scale = CSS_PX_PER_POINT * this.zoomLevel
    const add = (tag: string, text: string, x: number, y: number, width: number, height: number): void => {
      const element = document.createElement(tag)
      element.textContent = text
      element.style.left = `${x * scale}px`
      element.style.top = `${y * scale}px`
      element.style.width = `${width * scale}px`
      element.style.height = `${height * scale}px`
      layer.append(element)
    }
    for (const box of paragraphBoxes(data, this.layout)) {
      // A paragraph is reported once, whole, on the page it starts on.
      const info = this.layout.paras[box.para]
      if (!box.starts || !info?.len) continue
      const tag = info.h ? `h${Math.min(info.h, 6)}` : 'p'
      add(tag, this.doc.paragraphText(box.para), box.x, box.y, box.width, box.height)
    }
    for (const line of foreignLines(data, this.layout)) {
      add('div', line.text, line.x, line.y, line.width, line.height)
    }
    // A picture has no text of its own; it is reported as an image where it is drawn.
    for (const element of data.els) {
      if (element.t !== 'img') continue
      add('div', '', element.x, element.y, element.w, element.h)
      const picture = layer.lastElementChild!
      picture.setAttribute('role', 'img')
      picture.setAttribute('aria-label', '图片')
    }
    slot.holder.append(layer)
    slot.text = layer
  }

  /** Starts decoding the page's images; the page is drawn again as each one arrives. */
  private loadImages(index: number, page: PageData): void {
    if (!this.doc) return
    const generation = this.generation
    for (const [id, mime] of imageIds(page.els)) {
      if (this.images.has(id)) continue
      // Claim the id at once, so a redraw before the picture arrives does not decode it again.
      this.images.set(id, null)
      decodeImage(new Blob([this.doc.mediaData(index, id)], { type: mime })).then(
        (image) => {
          if (generation !== this.generation) return
          this.images.set(id, image)
          if (this.slots[index]?.canvas) this.draw(index)
        },
        () => {
          // A picture the runtime cannot decode leaves its place empty.
        },
      )
    }
  }

  private reportPage(force = false): void {
    const middle = this.scroller.scrollTop + this.scroller.clientHeight / 2
    let current = 0
    for (let index = 0; index < this.slots.length; index++) {
      if (this.slots[index].holder.offsetTop <= middle) current = index
      else break
    }
    if (force || current !== this.currentPage) {
      this.currentPage = current
      this.onPageChange(current, this.slots.length)
    }
  }
}
