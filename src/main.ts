import './style.css'
import { load } from '@dopame/acde-api'
import { Editor, type FormatState, INDENT_STEP, type ParagraphSettings } from './editor.ts'
import { type Alignment, Doc, type ListKind, loadEngine } from './engine.ts'
import { FIND_BAR_HTML, setupFindBar } from './findbar.ts'
import { decodeImage } from './images.ts'
import { type Commands, handleRequest } from './requests.ts'
import { PageView } from './view.ts'

const ZOOM_STEPS = [0.5, 0.75, 1, 1.25, 1.5, 2, 3]
const FONT_SIZES = [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 22, 24, 28, 36, 48, 72]
/** Word's names for the fonts the engine has a metric-compatible face for. */
const FONTS = ['Calibri', 'Arial', 'Times New Roman', 'Cambria', 'Georgia', 'Courier New']
const LINE_SPACINGS = [1, 1.15, 1.5, 2, 2.5, 3]
/** Text colours as six hex digits, with the name a screen reader says. */
const TEXT_COLORS: [string, string][] = [
  ['000000', '黑色'],
  ['7F7F7F', '灰色'],
  ['C00000', '深红'],
  ['FF0000', '红色'],
  ['ED7D31', '橙色'],
  ['FFC000', '金色'],
  ['00B050', '绿色'],
  ['00B0F0', '青色'],
  ['0070C0', '蓝色'],
  ['1F3864', '深蓝'],
  ['7030A0', '紫色'],
]
/** Word's highlight colours: the name the file stores, the name a screen reader says, the shade shown. */
const HIGHLIGHTS: [string, string, string][] = [
  ['yellow', '黄色', '#ffff00'],
  ['green', '绿色', '#00ff00'],
  ['cyan', '青色', '#00ffff'],
  ['magenta', '粉色', '#ff00ff'],
  ['red', '红色', '#ff0000'],
  ['blue', '蓝色', '#0000ff'],
  ['lightGray', '浅灰', '#c0c0c0'],
]
const DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
const DOCX_TYPE = { description: 'Word 文档', accept: { [DOCX_MIME]: ['.docx'] } }
const PDF_TYPE = { description: 'PDF 文档', accept: { 'application/pdf': ['.pdf'] } }
const IMAGE_TYPE = { description: '图片', accept: { 'image/png': ['.png'], 'image/jpeg': ['.jpg', '.jpeg'] } }
/** The widest a picture is placed, in points: the text width of an A4 page with common margins. */
const MAX_IMAGE_WIDTH = 432

document.querySelector<HTMLDivElement>('#app')!.innerHTML = `
<nav class="tabs" aria-label="打开的文档">
  <div class="tab-list" id="tabs" role="tablist"></div>
  <button type="button" id="new-tab" aria-label="新标签页">+</button>
</nav>
<header class="toolbar" role="toolbar" aria-label="工具栏">
  <div class="toolbar-row">
    <button type="button" id="new">新建</button>
    <button type="button" id="open">打开</button>
    <button type="button" id="save">保存</button>
    <button type="button" id="save-as">另存为</button>
    <button type="button" id="export-pdf">导出 PDF</button>
    <span class="divider" aria-hidden="true"></span>
    <button type="button" id="undo" disabled>撤销</button>
    <button type="button" id="redo" disabled>重做</button>
    <span class="divider" aria-hidden="true"></span>
    <button type="button" id="insert-table" aria-label="插入表格">表格</button>
    <span class="menu">
      <button type="button" id="table-menu" aria-label="表格操作" aria-haspopup="true" aria-expanded="false">行列</button>
      <div class="popover list" id="table-ops" role="menu" aria-label="表格操作" hidden>
        <button type="button" role="menuitem" data-value="row-above">在上方插入行</button>
        <button type="button" role="menuitem" data-value="row-below">在下方插入行</button>
        <button type="button" role="menuitem" data-value="column-left">在左侧插入列</button>
        <button type="button" role="menuitem" data-value="column-right">在右侧插入列</button>
        <button type="button" role="menuitem" data-value="delete-row">删除行</button>
        <button type="button" role="menuitem" data-value="delete-column">删除列</button>
        <button type="button" role="menuitem" data-value="merge">合并单元格</button>
        <button type="button" role="menuitem" data-value="split">拆分单元格</button>
      </div>
    </span>
    <button type="button" id="insert-image" aria-label="插入图片">图片</button>
    <button type="button" id="page-break" aria-label="插入分页符">分页</button>
    <span class="divider" aria-hidden="true"></span>
    <button type="button" id="find" aria-label="查找和替换">查找</button>
    <span class="divider" aria-hidden="true"></span>
    <button type="button" id="zoom-out" aria-label="缩小">−</button>
    <output id="zoom" aria-label="缩放比例">100%</output>
    <button type="button" id="zoom-in" aria-label="放大">+</button>
    <span class="spacer"></span>
    <output id="pages" aria-label="页码"></output>
  </div>
  <div class="toolbar-row">
    <select id="heading" aria-label="段落样式">
      <option value="0">正文</option>
      <option value="1">标题 1</option>
      <option value="2">标题 2</option>
      <option value="3">标题 3</option>
    </select>
    <select id="font" aria-label="字体">
      <option value="">字体</option>
      ${FONTS.map((font) => `<option value="${font}">${font}</option>`).join('')}
    </select>
    <select id="size" aria-label="字号">
      ${FONT_SIZES.map((size) => `<option value="${size}">${size}</option>`).join('')}
    </select>
    <button type="button" class="mark" id="bold" aria-label="加粗" aria-pressed="false"><b>B</b></button>
    <button type="button" class="mark" id="italic" aria-label="斜体" aria-pressed="false"><i>I</i></button>
    <button type="button" class="mark" id="underline" aria-label="下划线" aria-pressed="false"><u>U</u></button>
    <button type="button" class="mark" id="strike" aria-label="删除线" aria-pressed="false"><s>S</s></button>
    <button type="button" class="mark" id="superscript" aria-label="上标" aria-pressed="false">x<sup>2</sup></button>
    <button type="button" class="mark" id="subscript" aria-label="下标" aria-pressed="false">x<sub>2</sub></button>
    <span class="menu">
      <button type="button" class="mark swatch" id="color" aria-label="文字颜色" aria-haspopup="true" aria-expanded="false">A</button>
      <div class="popover" id="color-menu" role="group" aria-label="文字颜色" hidden>
        <button type="button" class="wide" data-value="auto">自动</button>
        ${TEXT_COLORS.map(([value, name]) => `<button type="button" class="chip" data-value="${value}" aria-label="${name}" style="--chip:#${value}"></button>`).join('')}
      </div>
    </span>
    <span class="menu">
      <button type="button" id="highlight" aria-label="高亮颜色" aria-haspopup="true" aria-expanded="false">高亮</button>
      <div class="popover" id="highlight-menu" role="group" aria-label="高亮颜色" hidden>
        <button type="button" class="wide" data-value="none">无</button>
        ${HIGHLIGHTS.map(([value, name, shade]) => `<button type="button" class="chip" data-value="${value}" aria-label="${name}" style="--chip:${shade}"></button>`).join('')}
      </div>
    </span>
    <button type="button" id="clear-format" aria-label="清除格式">清除</button>
    <button type="button" id="painter" aria-label="格式刷" aria-pressed="false">格式刷</button>
    <span class="divider" aria-hidden="true"></span>
    <button type="button" data-align="left" aria-label="左对齐" aria-pressed="false">左</button>
    <button type="button" data-align="center" aria-label="居中" aria-pressed="false">中</button>
    <button type="button" data-align="right" aria-label="右对齐" aria-pressed="false">右</button>
    <button type="button" data-align="justify" aria-label="两端对齐" aria-pressed="false">两端</button>
    <span class="divider" aria-hidden="true"></span>
    <button type="button" data-list="bullet" aria-label="项目符号" aria-pressed="false">• 列表</button>
    <button type="button" data-list="number" aria-label="编号" aria-pressed="false">1. 编号</button>
    <button type="button" id="outdent" aria-label="减少缩进">缩进 −</button>
    <button type="button" id="indent" aria-label="增加缩进">缩进 +</button>
    <select id="spacing" aria-label="行距">
      <option value="">行距</option>
      ${LINE_SPACINGS.map((spacing) => `<option value="${spacing}">${spacing} 倍</option>`).join('')}
    </select>
    <button type="button" id="paragraph" aria-label="段落设置">段落</button>
  </div>
</header>
${FIND_BAR_HTML}
<div class="workspaces" id="workspaces"></div>
<footer class="status">
  <span id="name"></span>
  <span id="count"></span>
  <span id="message" role="status"></span>
</footer>
<input type="file" id="file" hidden accept=".docx,${DOCX_MIME}">
<input type="file" id="image-file" hidden accept="image/png,image/jpeg">
<div class="modal" id="table-dialog" role="dialog" aria-modal="true" aria-labelledby="table-title" hidden>
  <form class="modal-box">
    <p id="table-title">插入表格</p>
    <div class="modal-fields">
      <label>行数 <input type="number" id="table-rows" min="1" max="100" value="3" required></label>
      <label>列数 <input type="number" id="table-columns" min="1" max="20" value="3" required></label>
    </div>
    <div class="modal-actions">
      <button type="button" data-choice="cancel">取消</button>
      <button type="submit" class="primary">插入</button>
    </div>
  </form>
</div>
<div class="modal" id="paragraph-dialog" role="dialog" aria-modal="true" aria-labelledby="paragraph-title" hidden>
  <form class="modal-box">
    <p id="paragraph-title">段落设置</p>
    <div class="modal-fields">
      <label>首行缩进（字符） <input type="number" id="first-line" min="0" max="20" step="0.5" required></label>
    </div>
    <div class="modal-fields">
      <label>段前（磅） <input type="number" id="space-before" min="0" max="500" step="1" placeholder="默认"></label>
      <label>段后（磅） <input type="number" id="space-after" min="0" max="500" step="1" placeholder="默认"></label>
    </div>
    <div class="modal-actions">
      <button type="button" data-choice="cancel">取消</button>
      <button type="submit" class="primary">确定</button>
    </div>
  </form>
</div>
<div class="modal" id="unsaved" role="dialog" aria-modal="true" aria-labelledby="unsaved-text" hidden>
  <div class="modal-box">
    <p id="unsaved-text"></p>
    <div class="modal-actions">
      <button type="button" data-choice="cancel">取消</button>
      <button type="button" data-choice="discard">不保存</button>
      <button type="button" data-choice="save" class="primary">保存</button>
    </div>
  </div>
</div>
`

function element<T extends HTMLElement>(id: string): T {
  return document.getElementById(id) as T
}

/** Holds the document area of the active tab. The other tabs keep theirs out of the page, drawn as they were. */
const workspaces = element('workspaces')
/** The page view of the active tab. Until the first document opens it is one that shows nothing. */
let view = new PageView(document.createElement('main'))
const toolbar = document.querySelector<HTMLElement>('.toolbar')!
const zoomLabel = element<HTMLOutputElement>('zoom')
const pagesLabel = element<HTMLOutputElement>('pages')
const nameLabel = element<HTMLSpanElement>('name')
const messageLabel = element<HTMLSpanElement>('message')
const fileInput = element<HTMLInputElement>('file')
const headingSelect = element<HTMLSelectElement>('heading')
const sizeSelect = element<HTMLSelectElement>('size')
const alignButtons = [...toolbar.querySelectorAll<HTMLButtonElement>('[data-align]')]
const listButtons = [...toolbar.querySelectorAll<HTMLButtonElement>('[data-list]')]
const fontSelect = element<HTMLSelectElement>('font')
const spacingSelect = element<HTMLSelectElement>('spacing')
const countLabel = element<HTMLSpanElement>('count')

const tabBar = document.querySelector<HTMLElement>('.tabs')!
const tabList = element('tabs')

/** One open document. The tab bar switches the view between them. */
interface Tab {
  id: number
  doc: Doc
  /** Keeps the caret, the selection and the undo history while another document has the view. */
  editor: Editor
  name: string
  /** The file the document was opened from or last saved to, when the app may write to it. */
  handle: FileSystemFileHandle | null
  dirty: boolean
  /** Whether the document started empty in this window, and not from a file or a request. */
  blank: boolean
  /** The document area: the scrolling element with the pages. In the page only while the tab is active. */
  scroller: HTMLElement
  /** Draws the document in `scroller`. A tab keeps its own, so its pages stay drawn while another tab shows. */
  view: PageView
  /** Whether `view` has laid the document out yet; it does so the first time the tab shows. */
  shown: boolean
  /** Where the document was scrolled to when another tab took its place. */
  scrollTop: number
}

const tabs: Tab[] = []
/** The tab whose document has the view, with that document and its editor. */
let active: Tab | null = null
let current: Doc | null = null
let editor: Editor | null = null
let nextTabId = 1
const acdeWindow = await load('window')


function say(message: string): void {
  messageLabel.textContent = message
}

function describe(error: unknown): string {
  return error instanceof Error ? error.message : String(error)
}

/** Shows the active document's name and whether it is saved: in the status bar, the window title and the tab bar. */
function showName(): void {
  const name = active?.name ?? ''
  const dirty = active?.dirty ?? false
  nameLabel.textContent = dirty ? `${name}（未保存）` : name
  acdeWindow?.setTitle(`${dirty ? '● ' : ''}${name} - Word`)
  showTabs()
}

function showTabs(): void {
  tabList.replaceChildren(
    ...tabs.map((tab) => {
      const item = document.createElement('div')
      item.className = 'tab'
      item.setAttribute('role', 'tab')
      item.setAttribute('aria-selected', String(tab === active))
      item.setAttribute('aria-label', tab.name)
      item.dataset.tab = String(tab.id)
      item.title = tab.name
      const label = document.createElement('span')
      label.className = 'tab-name'
      label.textContent = `${tab.dirty ? '● ' : ''}${tab.name}`
      const close = document.createElement('button')
      close.type = 'button'
      close.className = 'tab-close'
      close.setAttribute('aria-label', `关闭 ${tab.name}`)
      close.textContent = '×'
      item.append(label, close)
      return item
    }),
  )
}

function showFormat(state: FormatState | null): void {
  element<HTMLButtonElement>('undo').disabled = !editor?.canUndo
  element<HTMLButtonElement>('redo').disabled = !editor?.canRedo
  for (const mark of ['bold', 'italic', 'underline', 'strike'] as const) {
    element(mark).setAttribute('aria-pressed', String(state?.[mark] ?? false))
  }
  element('superscript').setAttribute('aria-pressed', String(state?.script === 'super'))
  element('subscript').setAttribute('aria-pressed', String(state?.script === 'sub'))
  element('painter').setAttribute('aria-pressed', String(state?.painting ?? false))
  // While the painter holds a format, the pointer over the document says so.
  workspaces.classList.toggle('painting', state?.painting ?? false)
  for (const button of alignButtons) {
    button.setAttribute('aria-pressed', String(state?.align === button.dataset.align))
  }
  for (const button of listButtons) {
    button.setAttribute('aria-pressed', String(state?.list === button.dataset.list))
  }
  headingSelect.value = String(state && state.heading <= 3 ? state.heading : 0)
  showOption(fontSelect, state?.font ?? '')
  showOption(spacingSelect, state?.spacing ? String(state.spacing) : '')
  element('color').style.setProperty('--swatch', state?.color ? `#${state.color}` : 'currentColor')
  if (state?.size) showOption(sizeSelect, String(state.size))
}

/** Selects a value in a list, adding it first when the list does not offer it. */
function showOption(select: HTMLSelectElement, value: string): void {
  if (![...select.options].some((option) => option.value === value)) select.add(new Option(value, value))
  select.value = value
}

/** Chinese characters count one each and a run of Latin letters or digits counts one, as Word counts. */
function countWords(text: string): number {
  return text.match(/[\u3400-\u9fff\uf900-\ufaff]|[\p{L}\p{N}]+(?:['’-][\p{L}\p{N}]+)*/gu)?.length ?? 0
}

let countTimer = 0
/** Shows the word count a moment after the last change, so that typing does not wait for it. */
function showCount(): void {
  clearTimeout(countTimer)
  countTimer = window.setTimeout(() => {
    const words = view.structure.paras.reduce((sum, _, para) => sum + countWords(current?.paragraphText(para) ?? ''), 0)
    countLabel.textContent = `字数 ${words.toLocaleString('zh-CN')}`
  }, 400)
}

let presenting: Promise<unknown> = Promise.resolve()

/**
 * Runs a task that changes which document has the view after the ones before it. Two never
 * overlap, so that the view, the editor and the document on screen always belong together.
 */
function inTurn<T>(task: () => Promise<T>): Promise<T> {
  const run = presenting.then(task)
  presenting = run.catch(() => {})
  return run
}

function makeTab(doc: Doc, name: string, handle: FileSystemFileHandle | null, blank: boolean): Tab {
  const scroller = document.createElement('main')
  scroller.className = 'workspace'
  scroller.setAttribute('aria-label', '文档')
  const tabView = new PageView(scroller)
  const editor = new Editor(doc, tabView)
  const tab: Tab = {
    id: nextTabId++,
    doc,
    editor,
    name,
    handle,
    dirty: false,
    blank,
    scroller,
    view: tabView,
    shown: false,
    scrollTop: 0,
  }
  tabView.onPageChange = (page, total) => {
    if (tab === active) pagesLabel.textContent = total > 0 ? `第 ${page + 1} / ${total} 页` : ''
  }
  tab.editor.onMessage = say
  tab.editor.onState = showFormat
  tab.editor.onChange = () => {
    showCount()
    if (tab.dirty) return
    tab.dirty = true
    showName()
  }
  return tab
}

/** Puts a tab's document area in the page, in place of the one that is there. */
function mount(tab: Tab): void {
  for (const other of tabs) {
    if (other !== tab && other.scroller.isConnected) {
      other.scroller.removeAttribute('id')
      other.scroller.remove()
    }
  }
  // Tests and tools find the document area of the document on screen by this id.
  tab.scroller.id = 'scroller'
  workspaces.append(tab.scroller)
}

/**
 * Shows a tab's document. Runs inside inTurn. The pages of a tab that was shown before are
 * still drawn, so coming back to it draws nothing again: its document area goes back into
 * the page where it was scrolled to.
 */
async function activate(tab: Tab): Promise<void> {
  if (active === tab) return
  const previous = active
  if (previous) {
    // An edit that is still being drawn finishes on its own document first.
    await previous.editor.idle()
    previous.scrollTop = previous.view.scrollTop
    previous.editor.deactivate()
  }
  mount(tab)
  try {
    if (!tab.shown) await tab.view.show(tab.doc)
    tab.shown = true
  } catch (error) {
    if (previous) {
      mount(previous)
      previous.view.scrollTo(previous.scrollTop)
      previous.editor.activate()
    } else {
      tab.scroller.remove()
    }
    throw error
  }
  active = tab
  view = tab.view
  current = tab.doc
  editor = tab.editor
  tab.editor.activate()
  // After the editor, which brings its caret into view: the document shows where it was left.
  view.scrollTo(tab.scrollTop)
  zoomLabel.textContent = `${Math.round(view.zoom * 100)}%`
  showCount()
  showName()
  view.input.focus({ preventScroll: true })
}

/** Closes a tab without asking: stops its editor and frees its document. Another tab has the view by now. */
function discard(tab: Tab): void {
  const index = tabs.indexOf(tab)
  if (index >= 0) tabs.splice(index, 1)
  tab.editor.dispose()
  tab.view.dispose()
  tab.scroller.remove()
  tab.doc.free()
  showTabs()
}

/** Whether a tab still holds the empty document it started with, with nothing done to it. */
function untouched(tab: Tab): boolean {
  return tab.blank && !tab.dirty && !tab.handle && !tab.editor.canUndo && !tab.editor.canRedo
}

/**
 * Opens a document in a tab of its own, next to the active one, and shows it. With
 * `overBlank` it takes the place of the active tab when that one is still untouched, as a
 * file opened in a new window takes the place of the empty document there.
 */
function openTab(
  doc: Doc,
  name: string | (() => string),
  options: { handle?: FileSystemFileHandle | null; blank?: boolean; overBlank?: boolean } = {},
): Promise<Tab> {
  return inTurn(async () => {
    const started = performance.now()
    const pages = doc.pageCount()
    const layoutMs = Math.round(performance.now() - started)
    // A name that depends on the open tabs is made up now, when the tabs before this one are in.
    const tab = makeTab(doc, typeof name === 'string' ? name : name(), options.handle ?? null, options.blank ?? false)
    const spare = options.overBlank && active && untouched(active) ? active : null
    tabs.splice(active ? tabs.indexOf(active) + 1 : tabs.length, 0, tab)
    try {
      await activate(tab)
    } catch (error) {
      discard(tab)
      throw error
    }
    if (spare) discard(spare)
    say(`共 ${pages} 页，排版用时 ${layoutMs} 毫秒`)
    return tab
  })
}

/** A name no open document has: 未命名文档, then 未命名文档 2, and so on. A tab that is about to close does not count. */
function untitledName(closing: Tab | null = null): string {
  for (let number = 1; ; number++) {
    const name = number === 1 ? '未命名文档' : `未命名文档 ${number}`
    if (!tabs.some((tab) => tab !== closing && tab.name === name)) return name
  }
}

/** An empty document: one paragraph to put the caret in. */
function emptyDocument(): Doc {
  const doc = Doc.create()
  doc.addParagraph('')
  return doc
}

function welcome(): Doc {
  const doc = Doc.create()
  doc.addHeading('Word for Dopame OS', 1)
  doc.addParagraph(
    'This page is laid out by rdocx, compiled to WebAssembly, and drawn on a 2D canvas by the app. ' +
      'Nothing here is rendered by a browser engine: every glyph sits where the layout engine put it.',
  )
  doc.addHeading('Try it', 2)
  doc.addParagraph('Click anywhere in the text and type. Enter splits a paragraph, Backspace joins two.')
  doc.addParagraph('Drag to select, or hold Shift with the arrow keys. Double-click selects a word.')
  doc.addParagraph('Use the toolbar or Ctrl+B, Ctrl+I and Ctrl+U to format what is selected.')
  doc.addParagraph('Save with Ctrl+S as a .docx file, or export a PDF.')
  doc.addHeading('Still to come', 2)
  doc.addParagraph('Undo, tables, a full Chinese font, and requests from agents.')
  return doc
}

async function openBytes(bytes: Uint8Array, name: string, handle: FileSystemFileHandle | null = null): Promise<void> {
  try {
    await openTab(Doc.open(bytes), name, { handle, overBlank: true })
  } catch (error) {
    say(`无法打开 ${name}：${describe(error)}`)
  }
}

/** The tab that already has a file open, when one does and the system can tell. */
async function tabOf(handle: FileSystemFileHandle): Promise<Tab | null> {
  for (const tab of tabs) {
    try {
      if (tab.handle && (await tab.handle.isSameEntry?.(handle))) return tab
    } catch {
      // A handle that cannot be compared is taken to be another file.
    }
  }
  return null
}

async function openHandle(handle: FileSystemFileHandle): Promise<void> {
  // A file that is open already is shown, not opened a second time.
  const open = await tabOf(handle)
  if (open) {
    await inTurn(() => activate(open))
    return
  }
  const file = await handle.getFile()
  await openBytes(new Uint8Array(await file.arrayBuffer()), file.name, handle)
}

async function openFile(file: File): Promise<void> {
  await openBytes(new Uint8Array(await file.arrayBuffer()), file.name)
}

function isCancel(error: unknown): boolean {
  return error instanceof DOMException && error.name === 'AbortError'
}

async function open(): Promise<void> {
  if (!window.showOpenFilePicker) {
    fileInput.click()
    return
  }
  try {
    const [handle] = await window.showOpenFilePicker({ types: [DOCX_TYPE] })
    await openHandle(handle)
  } catch (error) {
    if (!isCancel(error)) say(`无法打开：${describe(error)}`)
  }
}

async function write(handle: FileSystemFileHandle, bytes: ArrayBuffer): Promise<void> {
  const writable = await handle.createWritable()
  await writable.write(bytes)
  await writable.close()
}

/** Asks where to save. Returns null when the user cancels. */
async function pickSaveFile(
  name: string,
  extension: string,
  type: typeof DOCX_TYPE | typeof PDF_TYPE,
): Promise<FileSystemFileHandle | null> {
  if (!window.showSaveFilePicker) return null
  const suggestedName = `${name.replace(/\.[^.]+$/, '')}.${extension}`
  try {
    return await window.showSaveFilePicker({ suggestedName, types: [type] })
  } catch (error) {
    if (isCancel(error)) return null
    throw error
  }
}

/** Writes a tab's document to a file and makes that file the tab's own. */
async function saveTo(tab: Tab, handle: FileSystemFileHandle): Promise<void> {
  await write(handle, tab.doc.toDocx())
  tab.handle = handle
  tab.name = handle.name
  tab.dirty = false
  showName()
}

async function saveAs(tab: Tab | null = active): Promise<void> {
  if (!tab) return
  try {
    const handle = await pickSaveFile(tab.name, 'docx', DOCX_TYPE)
    if (!handle) return
    await saveTo(tab, handle)
    say(`已保存到 ${handle.name}`)
  } catch (error) {
    say(`保存失败：${describe(error)}`)
  }
}

async function save(tab: Tab | null = active): Promise<void> {
  if (!tab) return
  if (!tab.handle) {
    await saveAs(tab)
    return
  }
  try {
    await saveTo(tab, tab.handle)
    say(`已保存到 ${tab.handle.name}`)
  } catch {
    // A file that was dropped or handed over read-only cannot be written back.
    await saveAs(tab)
  }
}

async function exportPdf(): Promise<void> {
  if (!active) return
  try {
    const handle = await pickSaveFile(active.name, 'pdf', PDF_TYPE)
    if (!handle) return
    await write(handle, active.doc.toPdf())
    say(`已导出到 ${handle.name}`)
  } catch (error) {
    say(`导出失败：${describe(error)}`)
  }
}

type UnsavedChoice = 'save' | 'discard' | 'cancel'

/** Asks what to do with unsaved changes. confirm() does not exist in acdert, so the page asks. */
function askUnsaved(name: string): Promise<UnsavedChoice> {
  const modal = element('unsaved')
  element('unsaved-text').textContent = `「${name}」有未保存的修改，要保存吗？`
  modal.hidden = false
  modal.querySelector<HTMLButtonElement>('.primary')!.focus()
  return new Promise((resolve) => {
    const finish = (choice: UnsavedChoice): void => {
      modal.hidden = true
      modal.removeEventListener('click', onClick)
      modal.removeEventListener('keydown', onKey)
      refocus()
      resolve(choice)
    }
    const onClick = (event: MouseEvent): void => {
      const choice = (event.target as HTMLElement).closest<HTMLElement>('[data-choice]')?.dataset.choice
      if (choice) finish(choice as UnsavedChoice)
    }
    const onKey = (event: KeyboardEvent): void => {
      if (event.key === 'Escape') finish('cancel')
      // Keys typed in the dialog are not shortcuts of the document behind it.
      event.stopPropagation()
    }
    modal.addEventListener('click', onClick)
    modal.addEventListener('keydown', onKey)
  })
}

/**
 * Whether a tab may be closed: it is saved, or the user let its changes go. The tab is shown
 * first, so that the user sees the document the question is about.
 */
async function mayClose(tab: Tab): Promise<boolean> {
  if (!tab.dirty) return true
  await inTurn(() => activate(tab))
  const choice = await askUnsaved(tab.name)
  if (choice === 'cancel') return false
  if (choice === 'save') await save(tab)
  // A save the user cancelled leaves the document unsaved, and so open.
  return choice === 'discard' || !tab.dirty
}

/** Closes a tab, asking first when it has unsaved changes. The last tab leaves an empty document behind. */
async function closeTab(tab: Tab | null = active): Promise<void> {
  if (!tab || !tabs.includes(tab) || !(await mayClose(tab))) return
  await inTurn(async () => {
    const index = tabs.indexOf(tab)
    if (index < 0) return
    if (tabs.length === 1) tabs.push(makeTab(emptyDocument(), untitledName(tab), null, true))
    if (active === tab) await activate(tabs[index + 1] ?? tabs[index - 1])
    discard(tab)
  })
}

/** Shows the tab after the active one, or the one before it, going round at the ends. */
function stepTab(step: 1 | -1): void {
  if (!active || tabs.length < 2) return
  const next = tabs[(tabs.indexOf(active) + step + tabs.length) % tabs.length]
  void inTurn(() => activate(next))
}

/** Asks for the size of a table. Resolves to null when the user cancels. */
function askTable(): Promise<{ rows: number; columns: number } | null> {
  const modal = element('table-dialog')
  const form = modal.querySelector('form')!
  const rows = element<HTMLInputElement>('table-rows')
  const columns = element<HTMLInputElement>('table-columns')
  modal.hidden = false
  rows.focus()
  rows.select()
  return new Promise((resolve) => {
    const finish = (size: { rows: number; columns: number } | null): void => {
      modal.hidden = true
      form.removeEventListener('submit', onSubmit)
      modal.removeEventListener('click', onClick)
      modal.removeEventListener('keydown', onKey)
      refocus()
      resolve(size)
    }
    const onSubmit = (event: Event): void => {
      // The form checks the numbers against their ranges before it submits.
      event.preventDefault()
      finish({ rows: rows.valueAsNumber, columns: columns.valueAsNumber })
    }
    const onClick = (event: MouseEvent): void => {
      if ((event.target as HTMLElement).closest('[data-choice="cancel"]')) finish(null)
    }
    const onKey = (event: KeyboardEvent): void => {
      if (event.key === 'Escape') finish(null)
      // Keys typed in the dialog are not shortcuts of the document behind it.
      event.stopPropagation()
    }
    form.addEventListener('submit', onSubmit)
    modal.addEventListener('click', onClick)
    modal.addEventListener('keydown', onKey)
  })
}

/** The size the paragraph dialog counts a character of first-line indent as: the text's own, or the usual body size. */
const BODY_SIZE = 11

/** Asks for the paragraph's first-line indent and spacing. Resolves to null when the user cancels. */
function askParagraph(state: FormatState | null): Promise<ParagraphSettings | null> {
  const modal = element('paragraph-dialog')
  const form = modal.querySelector('form')!
  const firstLine = element<HTMLInputElement>('first-line')
  const before = element<HTMLInputElement>('space-before')
  const after = element<HTMLInputElement>('space-after')
  const size = state?.size ?? BODY_SIZE
  firstLine.value = String(Math.round(((state?.firstLine ?? 0) / size) * 2) / 2)
  before.value = state?.before == null ? '' : String(state.before)
  after.value = state?.after == null ? '' : String(state.after)
  modal.hidden = false
  firstLine.focus()
  firstLine.select()
  return new Promise((resolve) => {
    const finish = (settings: ParagraphSettings | null): void => {
      modal.hidden = true
      form.removeEventListener('submit', onSubmit)
      modal.removeEventListener('click', onClick)
      modal.removeEventListener('keydown', onKey)
      refocus()
      resolve(settings)
    }
    const onSubmit = (event: Event): void => {
      // The form checks the numbers against their ranges before it submits.
      event.preventDefault()
      finish({
        firstLine: firstLine.valueAsNumber * size,
        // An empty field gives the spacing back to the paragraph's style.
        before: before.value === '' ? null : before.valueAsNumber,
        after: after.value === '' ? null : after.valueAsNumber,
      })
    }
    const onClick = (event: MouseEvent): void => {
      if ((event.target as HTMLElement).closest('[data-choice="cancel"]')) finish(null)
    }
    const onKey = (event: KeyboardEvent): void => {
      if (event.key === 'Escape') finish(null)
      // Keys typed in the dialog are not shortcuts of the document behind it.
      event.stopPropagation()
    }
    form.addEventListener('submit', onSubmit)
    modal.addEventListener('click', onClick)
    modal.addEventListener('keydown', onKey)
  })
}

async function setParagraph(): Promise<void> {
  const settings = await askParagraph(editor?.formatState() ?? null)
  if (settings) await editor?.setParagraph(settings)
}

/** The table edits, by the name the table menu and the system menu know them by. */
const tableActions: Record<string, () => void> = {
  'row-above': () => void editor?.insertRow(false),
  'row-below': () => void editor?.insertRow(true),
  'column-left': () => void editor?.insertColumn(false),
  'column-right': () => void editor?.insertColumn(true),
  'delete-row': () => void editor?.deleteRow(),
  'delete-column': () => void editor?.deleteColumn(),
  merge: () => void editor?.mergeCells(),
  split: () => void editor?.splitCell(),
}

async function insertTable(): Promise<void> {
  const size = await askTable()
  if (size) await editor?.insertTable(size.rows, size.columns)
}

/** Puts a picture file into the document at 96 pixels to the inch, scaled down to the text width. */
async function placeImage(file: File): Promise<void> {
  try {
    const pixels = await decodeImage(file)
    if (!pixels.naturalWidth || !pixels.naturalHeight) throw new Error('无法读取这张图片')
    const width = Math.min((pixels.naturalWidth * 72) / 96, MAX_IMAGE_WIDTH)
    const height = (width * pixels.naturalHeight) / pixels.naturalWidth
    await editor?.insertImage(new Uint8Array(await file.arrayBuffer()), file.name, width, height)
  } catch (error) {
    say(`无法插入图片：${describe(error)}`)
  }
}

async function insertImage(): Promise<void> {
  if (!window.showOpenFilePicker) {
    element<HTMLInputElement>('image-file').click()
    return
  }
  try {
    const [handle] = await window.showOpenFilePicker({ types: [IMAGE_TYPE] })
    await placeImage(await handle.getFile())
  } catch (error) {
    if (!isCancel(error)) say(`无法插入图片：${describe(error)}`)
  }
}

async function newDocument(): Promise<void> {
  await openTab(emptyDocument(), untitledName, { blank: true })
}

function stepZoom(direction: 1 | -1): void {
  const steps = direction > 0 ? ZOOM_STEPS : [...ZOOM_STEPS].reverse()
  const next = steps.find((step) => (direction > 0 ? step > view.zoom + 0.001 : step < view.zoom - 0.001))
  if (next === undefined) return
  view.setZoom(next)
  zoomLabel.textContent = `${Math.round(view.zoom * 100)}%`
}

function refocus(): void {
  view.input.focus({ preventScroll: true })
}

// A button takes the focus on mouse down; keeping it on the document input keeps the caret
// and the selection alive while the toolbar is used.
toolbar.addEventListener('mousedown', (event) => {
  // While an input method composes, the click is let through to end the composition, as a
  // click anywhere else does; the focus comes back to the document after the click.
  const composing = view.input.parentElement?.classList.contains('composing')
  if (!composing && (event.target as HTMLElement).closest('button')) event.preventDefault()
})
toolbar.addEventListener('click', (event) => {
  const busy = document.querySelector('.modal:not([hidden])') || !element('find-bar').hidden
  if ((event.target as HTMLElement).closest('button') && !busy) refocus()
})

// The tab bar keeps the focus on the document too.
tabBar.addEventListener('mousedown', (event) => {
  if (!view.input.parentElement?.classList.contains('composing')) event.preventDefault()
})
tabBar.addEventListener('click', (event) => {
  const target = event.target as HTMLElement
  const id = target.closest<HTMLElement>('[data-tab]')?.dataset.tab
  const tab = tabs.find((candidate) => String(candidate.id) === id)
  if (target.closest('#new-tab')) void newDocument()
  else if (tab && target.closest('.tab-close')) void closeTab(tab)
  else if (tab) void inTurn(() => activate(tab))
})

// Nothing may fail in silence: a failure nobody sees looks like a button that does nothing.
window.addEventListener('error', (event) => say(`出错了：${event.message}`))
window.addEventListener('unhandledrejection', (event) => {
  if (!isCancel(event.reason)) say(`出错了：${describe(event.reason)}`)
})

element('new').addEventListener('click', () => void newDocument())
element('undo').addEventListener('click', () => void editor?.undo())
element('redo').addEventListener('click', () => void editor?.redo())
element('insert-table').addEventListener('click', () => void insertTable())
element('insert-image').addEventListener('click', () => void insertImage())
element<HTMLInputElement>('image-file').addEventListener('change', (event) => {
  const input = event.target as HTMLInputElement
  const file = input.files?.[0]
  input.value = ''
  if (file) void placeImage(file)
})
element('open').addEventListener('click', () => void open())
element('save').addEventListener('click', () => void save())
element('save-as').addEventListener('click', () => void saveAs())
element('export-pdf').addEventListener('click', () => void exportPdf())
element('zoom-in').addEventListener('click', () => stepZoom(1))
element('zoom-out').addEventListener('click', () => stepZoom(-1))
for (const mark of ['bold', 'italic', 'underline', 'strike'] as const) {
  element(mark).addEventListener('click', () => void editor?.toggle(mark))
}
element('superscript').addEventListener('click', () => void editor?.toggleScript('super'))
element('subscript').addEventListener('click', () => void editor?.toggleScript('sub'))
element('clear-format').addEventListener('click', () => void editor?.clearFormat())
element('painter').addEventListener('click', () => editor?.togglePainter())
element('paragraph').addEventListener('click', () => void setParagraph())
for (const button of alignButtons) {
  button.addEventListener('click', () => void editor?.setAlignment(button.dataset.align as Alignment))
}
for (const button of listButtons) {
  button.addEventListener('click', () => void editor?.toggleList(button.dataset.list as ListKind))
}
element('indent').addEventListener('click', () => void editor?.changeIndent(INDENT_STEP))
element('outdent').addEventListener('click', () => void editor?.changeIndent(-INDENT_STEP))
element('page-break').addEventListener('click', () => void editor?.insertPageBreak())
fontSelect.addEventListener('change', () => {
  if (fontSelect.value) void editor?.setFont(fontSelect.value)
  refocus()
})
spacingSelect.addEventListener('change', () => {
  if (spacingSelect.value) void editor?.setLineSpacing(Number(spacingSelect.value))
  refocus()
})

/** A button that opens a small panel of choices under it; a choice calls `choose` with its value. */
function setupMenu(buttonId: string, menuId: string, choose: (value: string) => void): void {
  const button = element<HTMLButtonElement>(buttonId)
  const menu = element(menuId)
  const close = (): void => {
    menu.hidden = true
    button.setAttribute('aria-expanded', 'false')
  }
  button.addEventListener('click', () => {
    const opening = menu.hidden
    for (const other of toolbar.querySelectorAll<HTMLElement>('.popover')) other.hidden = true
    menu.hidden = !opening
    button.setAttribute('aria-expanded', String(opening))
  })
  menu.addEventListener('click', (event) => {
    const value = (event.target as HTMLElement).closest<HTMLElement>('[data-value]')?.dataset.value
    if (value === undefined) return
    close()
    choose(value)
  })
  window.addEventListener('mousedown', (event) => {
    if (!menu.hidden && !(event.target as HTMLElement).closest('.menu')) close()
  })
  window.addEventListener('keydown', (event) => {
    if (event.key === 'Escape') close()
  })
}
setupMenu('color', 'color-menu', (value) => void editor?.setColor(value))
setupMenu('highlight', 'highlight-menu', (value) => void editor?.setHighlight(value))
setupMenu('table-menu', 'table-ops', (value) => tableActions[value]?.())

const findBar = setupFindBar(() => editor, refocus)
element('find').addEventListener('click', () => findBar.open())

headingSelect.addEventListener('change', () => {
  void editor?.setHeading(Number(headingSelect.value))
  refocus()
})
sizeSelect.addEventListener('change', () => {
  void editor?.setSize(Number(sizeSelect.value))
  refocus()
})

fileInput.addEventListener('change', () => {
  const file = fileInput.files?.[0]
  fileInput.value = ''
  if (file) void openFile(file)
})

workspaces.addEventListener('dragover', (event) => event.preventDefault())
workspaces.addEventListener('drop', (event) => {
  event.preventDefault()
  const file = event.dataTransfer?.files[0]
  if (file) void openFile(file)
})

window.addEventListener('keydown', (event) => {
  if (!event.ctrlKey && !event.metaKey) return
  const actions: Record<string, () => void> = {
    n: () => void newDocument(),
    tab: () => stepTab(event.shiftKey ? -1 : 1),
    o: () => void open(),
    s: () => void (event.shiftKey ? saveAs() : save()),
    f: () => findBar.open(),
    h: () => findBar.open(),
    '=': () => stepZoom(1),
    '+': () => stepZoom(1),
    '-': () => stepZoom(-1),
  }
  const action = actions[event.key.toLowerCase()]
  if (!action) return
  event.preventDefault()
  action()
})

/** What a request from another app or an agent may do; the same actions the toolbar offers. */
const commands: Commands = {
  async newDocument(title, blocks) {
    const doc = blocks.length > 0 ? Doc.create() : emptyDocument()
    for (const block of blocks) {
      if (block.type === 'heading') doc.addHeading(block.text, block.level)
      else doc.addParagraph(block.text)
    }
    // The document gets a tab of its own, so nothing that is open has to give way.
    const tab = await openTab(doc, title, { overBlank: true })
    // The content came with the request and is in no file yet.
    tab.dirty = blocks.length > 0
    showName()
    return true
  },
  async insertText(text, at) {
    if (at !== 'caret') editor?.moveToEdge(at === 'end')
    await editor?.insert(text)
  },
  replaceText: async (find, replacement) => (await editor?.replaceAll(find, replacement)) ?? 0,
  async save(handle) {
    if (!handle) {
      await save()
      return
    }
    if (active) await saveTo(active, handle)
  },
  async exportPdf(handle) {
    if (!handle) await exportPdf()
    else if (active) await write(handle, active.doc.toPdf())
  },
}

/** Resolves once the engine is loaded and the first document is on screen. */
let markStarted: () => void = () => {}
const started = new Promise<void>((resolve) => {
  markStarted = resolve
})

/** Carries out a request and shows its outcome in the status bar, where an agent can read it. */
async function receive(request: unknown): Promise<string> {
  // A request that opened the window arrives before the engine has loaded.
  await started
  let outcome: string
  try {
    outcome = await handleRequest(request, commands)
  } catch (error) {
    outcome = `请求无效：${describe(error)}`
  }
  say(outcome)
  return outcome
}

// What the system and other apps hand to this app. A file arrives from a double-click in the
// file manager or another app's open(); every other kind is one of this app's own requests.
acdeWindow?.on('request', (request) => {
  if (request.kind !== 'open-file') {
    void receive(request)
    return
  }
  const handle = (request as { handle?: FileSystemHandle }).handle
  if (handle?.kind !== 'file') return
  void started.then(() => openHandle(handle as FileSystemFileHandle))
})

// The system waits three seconds for an answer before it closes the window anyway, which is
// less than a person needs: keep the window, ask, and close it once the answer allows.
let closing = false
acdeWindow?.on('beforeclose', () => {
  if (closing || !tabs.some((tab) => tab.dirty)) return true
  void (async () => {
    // Every document with unsaved changes is asked about in turn; one "cancel" keeps the window.
    for (const tab of [...tabs]) {
      if (tabs.includes(tab) && !(await mayClose(tab))) return
    }
    closing = true
    void acdeWindow.close()
  })()
  return false
})

// The same commands as the toolbar, in the desktop's menu bar. The key combinations stay
// with the page's own key handling, so that no command runs twice.
acdeWindow?.setMenus([
  {
    id: 'file',
    label: '文件',
    items: [
      { id: 'new', label: '新建', onSelect: () => void newDocument() },
      { id: 'open', label: '打开…', onSelect: () => void open() },
      { type: 'separator' },
      { id: 'save', label: '保存', onSelect: () => void save() },
      { id: 'save-as', label: '另存为…', onSelect: () => void saveAs() },
      { id: 'export-pdf', label: '导出 PDF…', onSelect: () => void exportPdf() },
      { type: 'separator' },
      { id: 'close', label: '关闭文档', onSelect: () => void closeTab() },
    ],
  },
  {
    id: 'edit',
    label: '编辑',
    items: [
      { id: 'undo', label: '撤销', onSelect: () => void editor?.undo() },
      { id: 'redo', label: '重做', onSelect: () => void editor?.redo() },
      { type: 'separator' },
      { id: 'select-all', label: '全选', onSelect: () => editor?.selectAll() },
      { id: 'find', label: '查找和替换…', onSelect: () => findBar.open() },
    ],
  },
  {
    id: 'format',
    label: '格式',
    items: [
      { id: 'bold', label: '加粗', onSelect: () => void editor?.toggle('bold') },
      { id: 'italic', label: '斜体', onSelect: () => void editor?.toggle('italic') },
      { id: 'underline', label: '下划线', onSelect: () => void editor?.toggle('underline') },
      { id: 'strike', label: '删除线', onSelect: () => void editor?.toggle('strike') },
      { id: 'superscript', label: '上标', onSelect: () => void editor?.toggleScript('super') },
      { id: 'subscript', label: '下标', onSelect: () => void editor?.toggleScript('sub') },
      { id: 'clear-format', label: '清除格式', onSelect: () => void editor?.clearFormat() },
      { id: 'painter', label: '格式刷', onSelect: () => editor?.togglePainter() },
      { type: 'separator' },
      { id: 'align-left', label: '左对齐', onSelect: () => void editor?.setAlignment('left') },
      { id: 'align-center', label: '居中', onSelect: () => void editor?.setAlignment('center') },
      { id: 'align-right', label: '右对齐', onSelect: () => void editor?.setAlignment('right') },
      { id: 'align-justify', label: '两端对齐', onSelect: () => void editor?.setAlignment('justify') },
      { type: 'separator' },
      { id: 'bullets', label: '项目符号', onSelect: () => void editor?.toggleList('bullet') },
      { id: 'numbers', label: '编号', onSelect: () => void editor?.toggleList('number') },
      { id: 'indent', label: '增加缩进', onSelect: () => void editor?.changeIndent(INDENT_STEP) },
      { id: 'outdent', label: '减少缩进', onSelect: () => void editor?.changeIndent(-INDENT_STEP) },
      { id: 'paragraph', label: '段落设置…', onSelect: () => void setParagraph() },
    ],
  },
  {
    id: 'table',
    label: '表格',
    items: [
      { id: 'row-above', label: '在上方插入行', onSelect: tableActions['row-above'] },
      { id: 'row-below', label: '在下方插入行', onSelect: tableActions['row-below'] },
      { id: 'column-left', label: '在左侧插入列', onSelect: tableActions['column-left'] },
      { id: 'column-right', label: '在右侧插入列', onSelect: tableActions['column-right'] },
      { type: 'separator' },
      { id: 'delete-row', label: '删除行', onSelect: tableActions['delete-row'] },
      { id: 'delete-column', label: '删除列', onSelect: tableActions['delete-column'] },
      { type: 'separator' },
      { id: 'merge', label: '合并单元格', onSelect: tableActions.merge },
      { id: 'split', label: '拆分单元格', onSelect: tableActions.split },
    ],
  },
  {
    id: 'insert',
    label: '插入',
    items: [
      { id: 'table', label: '表格…', onSelect: () => void insertTable() },
      { id: 'image', label: '图片…', onSelect: () => void insertImage() },
      { id: 'page-break', label: '分页符', onSelect: () => void editor?.insertPageBreak() },
    ],
  },
])

// A small surface for end-to-end tests and, later, for requests from agents.
declare global {
  interface Window {
    wordApp: {
      openBytes(bytes: Uint8Array, name: string): Promise<void>
      /** The open documents in the order of their tabs. */
      tabs(): { name: string; dirty: boolean; active: boolean }[]
      state(): {
        name: string
        dirty: boolean
        pages: number
        zoom: number
        caret: { para: number; offset: number } | null
        /** The selected picture: its paragraph and which of that paragraph's pictures it is. */
        picture: { para: number; index: number } | null
        selection: string
        format: FormatState | null
        canUndo: boolean
        canRedo: boolean
      }
      /** The document's text, one entry per paragraph. */
      text(): string[]
      /** The text of every paragraph a caret can be in, table cells included, by paragraph index. */
      paragraphs(): string[]
      /** Whether the paragraph after this one is in the same cell or, outside tables, next to it in the body. */
      sibling(para: number): boolean
      /** How many pictures a page draws. */
      images(page: number): number
      /** Every picture in the text: its paragraph, and its width and height in points. */
      pictures(): [number, number, number][]
      /** Where a page draws its pictures: x, y, width and height in points. */
      pictureBoxes(page: number): [number, number, number, number][]
      /** How many elements of a kind a page draws, such as the rectangles of highlights. */
      elements(page: number, kind: string): number
      insert(text: string): Promise<void>
      /** Carries out a request as if another app had sent it, and resolves to its outcome. */
      request(request: unknown): Promise<string>
    }
  }
}
window.wordApp = {
  openBytes,
  tabs: () => tabs.map((tab) => ({ name: tab.name, dirty: tab.dirty, active: tab === active })),
  state: () => ({
    name: active?.name ?? '',
    dirty: active?.dirty ?? false,
    pages: view.pageCount,
    zoom: view.zoom,
    caret: editor?.position ?? null,
    picture: editor?.picture ?? null,
    selection: editor?.selectedText() ?? '',
    format: editor?.formatState() ?? null,
    canUndo: editor?.canUndo ?? false,
    canRedo: editor?.canRedo ?? false,
  }),
  text: () => current?.text().split('\n') ?? [],
  paragraphs: () => view.structure.paras.map((_, para) => current?.paragraphText(para) ?? ''),
  sibling: (para) => view.structure.paras[para]?.m === 1,
  images: (page) => view.pageData(page)?.els.filter((element) => element.t === 'img').length ?? 0,
  pictures: () =>
    view.structure.paras.flatMap((_, para) =>
      (current?.images(para) ?? []).map(([width, height]): [number, number, number] => [para, width, height]),
    ),
  pictureBoxes: (page) =>
    (view.pageData(page)?.els ?? []).flatMap((element): [number, number, number, number][] =>
      element.t === 'img' ? [[element.x, element.y, element.w, element.h]] : [],
    ),
  elements: (page, kind) => view.pageData(page)?.els.filter((element) => element.t === kind).length ?? 0,
  insert: (text) => editor?.insert(text) ?? Promise.resolve(),
  request: receive,
}

say('正在加载排版引擎…')
try {
  await loadEngine()
  await openTab(welcome(), untitledName, { blank: true })
} catch (error) {
  say(`排版引擎加载失败：${describe(error)}`)
}
markStarted()
