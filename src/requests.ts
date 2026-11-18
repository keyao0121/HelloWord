// Requests from other apps and from agents. Another app sends one with
//   import { open } from 'acde:open'
//   await open('word', { kind: 'new-document', text: '# Title\nFirst paragraph' })
// and needs "open": ["word"] in its capabilities. The kinds are listed under "requests" in
// acde.config.jsonc, which is all the runtime checks; the rest of each request is checked here.
// A request gets no answer back: its result shows in the window, where an agent reads it
// from the element tree.

/** One block of a new document. */
export type Block = { type: 'heading'; level: number; text: string } | { type: 'paragraph'; text: string }

export type WordRequest =
  /** Replace the document with a new one. `text` is one paragraph per line, with "# ", "## " and "### " for headings. */
  | { kind: 'new-document'; title?: string; text?: string; blocks?: Block[] }
  /** Type text at the caret, or at the start or the end of the document. A line feed starts a paragraph. */
  | { kind: 'insert-text'; text: string; at?: 'caret' | 'start' | 'end' }
  /** Replace every occurrence of a text. */
  | { kind: 'replace-text'; find: string; replace: string }
  /** Save the document: to `handle` when given, else where it was last saved, else by asking the user. */
  | { kind: 'save-document'; handle?: FileSystemFileHandle }
  /** Write the document as a PDF: to `handle` when given, else by asking the user. */
  | { kind: 'export-pdf'; handle?: FileSystemFileHandle }

/** What the app does for a request. */
export interface Commands {
  /** Resolves to false when the user kept the document on screen. */
  newDocument(title: string, blocks: Block[]): Promise<boolean>
  insertText(text: string, at: 'caret' | 'start' | 'end'): Promise<void>
  /** Resolves to the number of replacements. */
  replaceText(find: string, replacement: string): Promise<number>
  save(handle?: FileSystemFileHandle): Promise<void>
  exportPdf(handle?: FileSystemFileHandle): Promise<void>
}

function text(value: unknown, field: string, required = true): string {
  if (typeof value === 'string' && (value !== '' || !required)) return value
  if (value === undefined && !required) return ''
  throw new Error(`${field} 必须是${required ? '非空' : ''}字符串`)
}

function fileHandle(value: unknown): FileSystemFileHandle | undefined {
  if (value === undefined || value === null) return undefined
  if (typeof value === 'object' && (value as FileSystemHandle).kind === 'file') return value as FileSystemFileHandle
  throw new Error('handle 必须是文件句柄')
}

/** One paragraph per line; "# ", "## " and "### " start a heading; empty lines are left out. */
export function blocksFromText(source: string): Block[] {
  const blocks: Block[] = []
  for (const line of source.replace(/\r\n?/g, '\n').split('\n')) {
    if (!line.trim()) continue
    const heading = /^(#{1,3})\s+(.*)$/.exec(line)
    blocks.push(
      heading ? { type: 'heading', level: heading[1].length, text: heading[2] } : { type: 'paragraph', text: line },
    )
  }
  return blocks
}

function checkedBlocks(value: unknown): Block[] {
  if (!Array.isArray(value)) throw new Error('blocks 必须是数组')
  return value.map((entry, index) => {
    const block = entry as Partial<{ type: string; level: number; text: string }>
    const content = text(block?.text, `blocks[${index}].text`)
    if (block.type === 'paragraph') return { type: 'paragraph', text: content }
    if (block.type === 'heading') {
      const level = Number(block.level ?? 1)
      if (!Number.isInteger(level) || level < 1 || level > 3) throw new Error(`blocks[${index}].level 必须是 1 到 3`)
      return { type: 'heading', level, text: content }
    }
    throw new Error(`blocks[${index}].type 必须是 heading 或 paragraph`)
  })
}

/** Carries out one request and resolves to a line that says what was done. Rejects when the request is malformed. */
export async function handleRequest(request: unknown, commands: Commands): Promise<string> {
  const fields = (request ?? {}) as Record<string, unknown>
  switch (fields.kind) {
    case 'new-document': {
      const blocks = fields.blocks !== undefined ? checkedBlocks(fields.blocks) : blocksFromText(text(fields.text, 'text', false))
      const title = text(fields.title, 'title', false) || '未命名文档'
      const created = await commands.newDocument(title, blocks)
      return created ? `已按请求新建文档「${title}」，共 ${blocks.length} 段` : '新建文档的请求已取消：当前文档未保存'
    }
    case 'insert-text': {
      const at = fields.at ?? 'caret'
      if (at !== 'caret' && at !== 'start' && at !== 'end') throw new Error('at 必须是 caret、start 或 end')
      const inserted = text(fields.text, 'text')
      await commands.insertText(inserted, at)
      return `已按请求插入 ${Array.from(inserted).length} 个字符`
    }
    case 'replace-text': {
      const count = await commands.replaceText(text(fields.find, 'find'), text(fields.replace, 'replace', false))
      return `已按请求替换 ${count} 处`
    }
    case 'save-document':
      await commands.save(fileHandle(fields.handle))
      return '已按请求保存文档'
    case 'export-pdf':
      await commands.exportPdf(fileHandle(fields.handle))
      return '已按请求导出 PDF'
    default:
      throw new Error(`不认识的请求类型 ${String(fields.kind)}`)
  }
}
