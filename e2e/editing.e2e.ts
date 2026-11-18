// End-to-end checks of editing, formatting and saving, driven through real pointer and key
// events in a headless acdert. Run with: acde test

import { existsSync, readFileSync } from 'node:fs'
import { join } from 'node:path'
import { expect, test } from 'acde/test'

interface State {
  name: string
  dirty: boolean
  pages: number
  caret: { para: number; offset: number } | null
  selection: string
  format: { bold: boolean; italic: boolean; underline: boolean; size: number | null; align: string | null; heading: number } | null
  canUndo: boolean
  canRedo: boolean
}

type App = { evaluate<T>(code: string): Promise<T> }

async function state(app: App): Promise<State> {
  return JSON.parse(await app.evaluate<string>('JSON.stringify(window.wordApp?.state() ?? null)')) as State
}

async function text(app: App): Promise<string[]> {
  return JSON.parse(await app.evaluate<string>('JSON.stringify(window.wordApp.text())')) as string[]
}

/** Waits until a check on the state holds; edits and saves finish a moment after the event. */
async function until(app: App, check: (state: State) => boolean, what: string): Promise<State> {
  for (let attempt = 0; attempt < 100; attempt++) {
    const now = await state(app)
    if (now && check(now)) return now
    await new Promise((resolve) => setTimeout(resolve, 50))
  }
  throw new Error(`timed out waiting for ${what}: ${JSON.stringify(await state(app))}`)
}

/** Page coordinates, in CSS pixels, of a point given in points on the first page. */
async function onPage(app: App, x: number, y: number): Promise<[number, number]> {
  const [left, top] = JSON.parse(
    await app.evaluate<string>(
      '(() => { const b = document.querySelector(".page").getBoundingClientRect(); return JSON.stringify([b.left, b.top]) })()',
    ),
  ) as [number, number]
  return [left + (x * 96) / 72, top + (y * 96) / 72]
}

// The welcome document: a heading at the top, then the paragraph "This page is laid out ..."
// whose first line sits about 100 points down the page, starting at the 72 point margin.
const FIRST_LINE_Y = 100
const MARGIN_X = 72

test('selects a word, formats it, and aligns its paragraph', async ({ app }) => {
  await until(app!, (now) => now.pages === 1 && now.caret !== null, 'the welcome document')

  await app!.mouse.dblclick(...(await onPage(app!, MARGIN_X + 8, FIRST_LINE_Y)))
  const selected = await until(app!, (now) => now.selection !== '', 'a selected word')
  expect(selected.selection).toBe('This')
  expect(selected.format?.bold).toBe(false)

  await app!.getByRole('button', { name: '加粗' }).click()
  await until(app!, (now) => now.format?.bold === true, 'bold')
  await app!.keyboard.press('Control+i')
  await until(app!, (now) => now.format?.italic === true, 'italic')
  await app!.keyboard.press('Control+u')
  const marked = await until(app!, (now) => now.format?.underline === true, 'underline')
  // Formatting keeps the selection, so it can be turned off again the same way.
  expect(marked.selection).toBe('This')
  await app!.keyboard.press('Control+b')
  await until(app!, (now) => now.format?.bold === false, 'bold off')

  await app!.getByRole('button', { name: '居中' }).click()
  const centred = await until(app!, (now) => now.format?.align === 'center', 'centre alignment')
  expect(centred.dirty).toBe(true)
})

test('selects by dragging and with Shift, and typing replaces the selection', async ({ app }) => {
  await until(app!, (now) => now.pages === 1 && now.caret !== null, 'the welcome document')

  const [startX, startY] = await onPage(app!, MARGIN_X, FIRST_LINE_Y)
  const [endX] = await onPage(app!, MARGIN_X + 40, FIRST_LINE_Y)
  await app!.mouse.move(startX, startY)
  await app!.mouse.down()
  await app!.mouse.move(endX, startY, { steps: 4 })
  await app!.mouse.up()
  const dragged = await until(app!, (now) => now.selection.length > 2, 'a dragged selection')
  expect('This page is laid out'.startsWith(dragged.selection)).toBe(true)

  await app!.keyboard.type('Each sheet')
  await until(app!, (now) => now.selection === '' && now.dirty, 'the typed replacement')
  expect((await text(app!))[1].startsWith('Each sheet')).toBe(true)

  await app!.keyboard.press('Shift+ArrowLeft')
  await app!.keyboard.press('Shift+ArrowLeft')
  const extended = await until(app!, (now) => now.selection.length === 2, 'two characters selected')
  expect(extended.selection).toBe('et')
  await app!.keyboard.press('Backspace')
  await until(app!, (now) => now.selection === '', 'the deletion')
  expect((await text(app!))[1].startsWith('Each she')).toBe(true)

  // What an input method commits arrives without key events.
  await app!.keyboard.insertText('中文')
  await until(app!, (now) => now.caret?.offset === 10, 'the committed text')
  expect((await text(app!))[1].startsWith('Each she中文')).toBe(true)

  await app!.keyboard.press('Control+a')
  const all = await until(app!, (now) => now.selection.includes('Still to come'), 'everything selected')
  expect(all.selection.split('\n').length).toBeGreaterThan(5)
})

test('saves a docx, opens it again, and exports a PDF', async ({ app, acde, home }) => {
  await until(app!, (now) => now.pages === 1 && now.caret !== null, 'the welcome document')
  await app!.keyboard.type('Saved: ')
  await until(app!, (now) => now.dirty, 'an unsaved change')

  const docx = join(home, 'saved.docx')
  await acde.chooseSaveFile(docx)
  await app!.keyboard.press('Control+s')
  const saved = await until(app!, (now) => !now.dirty, 'the save')
  expect(saved.name).toBe('saved.docx')
  expect(existsSync(docx)).toBe(true)
  // A docx is a zip archive.
  expect(readFileSync(docx).subarray(0, 2).toString('latin1')).toBe('PK')

  // A second save goes to the same file without asking.
  await app!.keyboard.type('again ')
  await until(app!, (now) => now.dirty, 'a second change')
  await app!.keyboard.press('Control+s')
  await until(app!, (now) => !now.dirty, 'the second save')
  const before = await text(app!)

  // Closing the only tab leaves an empty document; the file then opens in its place.
  await app!.getByRole('button', { name: '关闭 saved.docx' }).click()
  await until(app!, (now) => now.name === '未命名文档', 'an empty document')
  await acde.chooseFiles(docx)
  await app!.getByRole('button', { name: '打开' }).click()
  await until(app!, (now) => now.name === 'saved.docx', 'the reopened file')
  const after = await text(app!)
  expect(after[0]).toBe('Saved: again Word for Dopame OS')
  expect(after).toEqual(before)

  const pdf = join(home, 'saved.pdf')
  await acde.chooseSaveFile(pdf)
  await app!.getByRole('button', { name: '导出 PDF' }).click()
  for (let attempt = 0; attempt < 100 && !existsSync(pdf); attempt++) {
    await new Promise((resolve) => setTimeout(resolve, 50))
  }
  expect(readFileSync(pdf).subarray(0, 4).toString('latin1')).toBe('%PDF')
})

test('undoes and redoes typing, formatting and paragraph edits', async ({ app }) => {
  await until(app!, (now) => now.pages === 1 && now.caret !== null, 'the welcome document')
  const original = await text(app!)
  expect((await state(app!)).canUndo).toBe(false)

  // A run of typed characters is one step.
  await app!.keyboard.type('abc')
  await until(app!, (now) => now.caret?.offset === 3, 'the typed text')
  expect((await text(app!))[0]).toBe('abcWord for Dopame OS')
  await app!.keyboard.press('Control+z')
  const undone = await until(app!, (now) => !now.canUndo, 'the undo')
  expect(await text(app!)).toEqual(original)
  expect(undone.caret).toEqual({ para: 0, offset: 0 })
  expect(undone.canRedo).toBe(true)
  await app!.keyboard.press('Control+y')
  const redone = await until(app!, (now) => !now.canRedo, 'the redo')
  expect((await text(app!))[0]).toBe('abcWord for Dopame OS')
  expect(redone.caret).toEqual({ para: 0, offset: 3 })

  // A new paragraph, a format change and a deletion are each a step of their own.
  await app!.keyboard.press('Enter')
  await until(app!, (now) => now.caret?.para === 1, 'the new paragraph')
  await app!.keyboard.press('ArrowDown')
  await app!.keyboard.press('Home')
  await until(app!, (now) => now.caret?.para === 2 && now.caret.offset === 0, 'the caret in the next paragraph')
  for (let count = 0; count < 4; count++) await app!.keyboard.press('Shift+ArrowRight')
  await until(app!, (now) => now.selection === 'This', 'a selected word')
  await app!.keyboard.press('Control+b')
  await until(app!, (now) => now.format?.bold === true, 'bold')
  await app!.keyboard.press('Backspace')
  await until(app!, (now) => now.selection === '', 'the deletion')
  expect((await text(app!))[2].startsWith(' page is laid out')).toBe(true)

  await app!.keyboard.press('Control+z')
  const restored = await until(app!, (now) => now.selection === 'This', 'the deletion undone')
  expect(restored.format?.bold).toBe(true)
  await app!.keyboard.press('Control+z')
  await until(app!, (now) => now.format?.bold === false, 'the bold undone')
  await app!.keyboard.press('Control+z')
  await until(app!, (now) => now.caret?.para === 0 && now.caret.offset === 3, 'the new paragraph undone')
  await app!.keyboard.press('Control+z')
  await until(app!, (now) => !now.canUndo, 'everything undone')
  expect(await text(app!)).toEqual(original)

  // A new edit drops what could be redone.
  await app!.keyboard.type('z')
  const after = await until(app!, (now) => now.canUndo, 'a new edit')
  expect(after.canRedo).toBe(false)
})

test('asks before a document with unsaved changes is closed', async ({ app, acde, home }) => {
  await until(app!, (now) => now.pages === 1 && now.caret !== null, 'the welcome document')
  const hidden = (): Promise<boolean> => app!.evaluate<boolean>('document.getElementById("unsaved").hidden')
  await app!.keyboard.type('draft ')
  await until(app!, (now) => now.dirty, 'an unsaved change')

  await app!.getByRole('button', { name: '关闭 未命名文档' }).click()
  expect(await hidden()).toBe(false)
  await app!.getByRole('button', { name: '取消' }).click()
  expect(await hidden()).toBe(true)
  expect((await text(app!))[0]).toBe('draft Word for Dopame OS')
  expect((await state(app!)).dirty).toBe(true)

  // Saving from the prompt writes the file, then closes the tab. The last tab to close
  // leaves an empty document behind.
  const docx = join(home, 'draft.docx')
  await acde.chooseSaveFile(docx)
  await app!.getByRole('button', { name: '关闭 未命名文档' }).click()
  await app!.getByRole('button', { name: '保存', exact: true }).last().click()
  await until(app!, (now) => !now.dirty && now.name === '未命名文档' && !now.canUndo, 'the empty document after saving')
  expect(existsSync(docx)).toBe(true)
  expect(await app!.evaluate<string>('JSON.stringify(window.wordApp.paragraphs())')).toBe('[""]')

  await app!.keyboard.type('scrap ')
  await until(app!, (now) => now.dirty, 'another unsaved change')
  await app!.getByRole('button', { name: '关闭 未命名文档' }).click()
  await app!.getByRole('button', { name: '不保存' }).click()
  await until(app!, (now) => !now.dirty, 'the empty document after discarding')
  expect(await app!.evaluate<string>('JSON.stringify(window.wordApp.paragraphs())')).toBe('[""]')
})

test('an editor stops when its document is closed', async ({ app, acde }) => {
  await until(app!, (now) => now.pages === 1 && now.caret !== null, 'the welcome document')
  // Close the document a few times, then open two more without waiting, as hurried clicks do.
  for (let round = 0; round < 3; round++) {
    await app!.keyboard.insertText('旧')
    await until(app!, (now) => now.dirty, 'an unsaved change')
    await app!.getByRole('button', { name: '关闭 未命名文档' }).click()
    await app!.getByRole('button', { name: '不保存' }).click()
    await until(app!, (now) => !now.dirty, 'an empty document')
  }
  await app!.evaluate<void>('(() => { document.getElementById("new").click(); document.getElementById("new").click() })()')
  await until(app!, (now) => now.name === '未命名文档 3' && now.caret !== null, 'the last new document')

  // What an input method commits goes to the document on screen, once, and to no other.
  await app!.keyboard.insertText('新文档')
  await until(app!, (now) => now.caret?.offset === 3, 'the committed text')
  expect(await app!.evaluate<string>('JSON.stringify(window.wordApp.paragraphs())')).toBe('["新文档"]')
  const message = await app!.evaluate<string>('document.getElementById("message").textContent')
  expect(message.startsWith('出错了')).toBe(false)
  const errors = await acde.console({ level: 'errors' })
  expect(errors.entries.map((entry: { text?: string; message?: string }) => entry.text ?? entry.message ?? '')).toEqual([])

  // The toolbar still answers afterwards.
  await app!.getByRole('button', { name: '撤销' }).click()
  await until(app!, (now) => !now.canUndo, 'the undo')
  expect(await app!.evaluate<string>('JSON.stringify(window.wordApp.paragraphs())')).toBe('[""]')
})

test('the toolbar takes clicks while the document is scrolled', async ({ app }) => {
  await until(app!, (now) => now.pages === 1 && now.caret !== null, 'the welcome document')
  // In acdert the part of the page that has scrolled out above the document area still takes
  // the pointer there, unless the toolbar is stacked over it.
  const scrolled = await app!.evaluate<number>(
    '(() => { const s = document.getElementById("scroller"); s.scrollTop = 200; return s.scrollTop })()',
  )
  expect(scrolled).toBeGreaterThan(100)
  const [x, y, under] = JSON.parse(
    await app!.evaluate<string>(
      '(() => { const b = document.getElementById("zoom-in").getBoundingClientRect(); const x = b.left + b.width / 2, y = b.top + b.height / 2; return JSON.stringify([x, y, document.elementFromPoint(x, y)?.id ?? null]) })()',
    ),
  ) as [number, number, string | null]
  expect(under).toBe('zoom-in')
  await app!.mouse.click(x, y)
  for (let attempt = 0; attempt < 40; attempt++) {
    if ((await app!.evaluate<number>('window.wordApp.state().zoom')) > 1) break
    await new Promise((resolve) => setTimeout(resolve, 50))
  }
  expect(await app!.evaluate<number>('window.wordApp.state().zoom')).toBeGreaterThan(1)
})
