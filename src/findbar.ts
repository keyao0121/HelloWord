// The find and replace bar: a strip under the toolbar that searches the document, steps
// through what it finds, and replaces one occurrence or all of them.

import type { Editor, FindResult } from './editor.ts'

export const FIND_BAR_HTML = `
<div class="find-bar" id="find-bar" role="search" aria-label="查找和替换" hidden>
  <input id="find-text" type="text" aria-label="查找内容" placeholder="查找" autocomplete="off" spellcheck="false">
  <output id="find-count" aria-label="查找结果"></output>
  <button type="button" id="find-previous">上一个</button>
  <button type="button" id="find-next">下一个</button>
  <span class="divider" aria-hidden="true"></span>
  <input id="replace-text" type="text" aria-label="替换为" placeholder="替换为" autocomplete="off" spellcheck="false">
  <button type="button" id="replace-one">替换</button>
  <button type="button" id="replace-all">全部替换</button>
  <span class="spacer"></span>
  <button type="button" id="find-close" aria-label="关闭查找">×</button>
</div>
`

export interface FindBar {
  open(): void
  close(): void
}

/** Wires the bar up. `editor` gives the editor of the document on screen, which changes when another is opened. */
export function setupFindBar(editor: () => Editor | null, refocus: () => void): FindBar {
  const get = <T extends HTMLElement>(id: string): T => document.getElementById(id) as T
  const bar = get('find-bar')
  const query = get<HTMLInputElement>('find-text')
  const replacement = get<HTMLInputElement>('replace-text')
  const count = get<HTMLOutputElement>('find-count')

  const show = (result: FindResult): void => {
    count.textContent = result.total === 0 ? (query.value ? '无结果' : '') : `${result.index || '–'} / ${result.total}`
  }
  const step = (backwards: boolean): void => {
    const current = editor()
    if (current && query.value) show(current.find(query.value, backwards))
  }
  const close = (): void => {
    bar.hidden = true
    refocus()
  }

  query.addEventListener('input', () => show({ index: 0, total: editor()?.count(query.value) ?? 0 }))
  get('find-next').addEventListener('click', () => step(false))
  get('find-previous').addEventListener('click', () => step(true))
  get('replace-one').addEventListener('click', () => {
    const current = editor()
    if (current && query.value) void current.replace(query.value, replacement.value).then(show)
  })
  get('replace-all').addEventListener('click', () => {
    const current = editor()
    if (!current || !query.value) return
    void current.replaceAll(query.value, replacement.value).then((replaced) => {
      count.textContent = `已替换 ${replaced} 处`
    })
  })
  get('find-close').addEventListener('click', close)
  bar.addEventListener('keydown', (event) => {
    if (event.key === 'Escape') {
      event.preventDefault()
      close()
    } else if (event.key === 'Enter') {
      event.preventDefault()
      if (event.target === replacement) get('replace-one').click()
      else step(event.shiftKey)
    }
  })
  // The buttons leave the focus in the text boxes, so Enter keeps stepping through the results.
  bar.addEventListener('mousedown', (event) => {
    if ((event.target as HTMLElement).closest('button')) event.preventDefault()
  })

  return {
    open() {
      bar.hidden = false
      query.focus()
      query.select()
      show({ index: 0, total: editor()?.count(query.value) ?? 0 })
    },
    close,
  }
}
