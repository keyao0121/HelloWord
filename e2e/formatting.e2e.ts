// End-to-end checks of the formatting controls, find and replace, and keyboard navigation.
// Run with: acde test

import { expect, test } from 'acde/test'

interface Format {
  bold: boolean
  italic: boolean
  underline: boolean
  size: number | null
  font: string | null
  color: string | null
  list: string | null
  indent: number
  spacing: number | null
  strike: boolean
  script: string | null
  firstLine: number
  before: number | null
  after: number | null
  painting: boolean
}

interface State {
  pages: number
  caret: { para: number; offset: number } | null
  selection: string
  format: Format | null
  canUndo: boolean
}

type App = { evaluate<T>(code: string): Promise<T> }

async function state(app: App): Promise<State> {
  return JSON.parse(await app.evaluate<string>('JSON.stringify(window.wordApp?.state() ?? null)')) as State
}

async function paragraphs(app: App): Promise<string[]> {
  return JSON.parse(await app.evaluate<string>('JSON.stringify(window.wordApp.paragraphs())')) as string[]
}

async function until(app: App, check: (state: State) => boolean, what: string): Promise<State> {
  for (let attempt = 0; attempt < 100; attempt++) {
    const now = await state(app)
    if (now && check(now)) return now
    await new Promise((resolve) => setTimeout(resolve, 50))
  }
  throw new Error(`timed out waiting for ${what}: ${JSON.stringify(await state(app))}`)
}

const ready = (app: App): Promise<State> => until(app, (now) => now.pages === 1 && now.caret !== null, 'the welcome document')

/** Picks a value in a drop-down list the way a person does: the value changes and the list says so. */
function choose(app: App, id: string, value: string): Promise<void> {
  return app.evaluate<void>(
    `(() => { const list = document.getElementById(${JSON.stringify(id)}); list.value = ${JSON.stringify(value)}; list.dispatchEvent(new Event('change', { bubbles: true })) })()`,
  )
}

test('a format chosen with nothing selected goes to the text typed next', async ({ app }) => {
  await ready(app!)
  // The end of the document: the caret is after a full stop, inside no word.
  await app!.keyboard.press('Control+End')
  const start = await until(app!, (now) => (now.caret?.para ?? 0) > 5 && now.caret!.offset > 10, 'the end of the document')
  expect(start.format?.bold).toBe(false)

  await app!.keyboard.press('Control+b')
  await until(app!, (now) => now.format?.bold === true, 'bold waiting for text')
  expect(start.canUndo).toBe(false)
  await app!.keyboard.type('Strong')
  const typed = await until(app!, (now) => now.caret?.offset === start.caret!.offset + 6, 'the bold text')
  expect(typed.format?.bold).toBe(true)

  await app!.keyboard.press('Control+b')
  await app!.keyboard.type('plain')
  const plain = await until(app!, (now) => now.caret?.offset === start.caret!.offset + 11, 'the plain text')
  expect(plain.format?.bold).toBe(false)
  // Back inside the bold word, the toolbar shows bold again.
  for (let count = 0; count < 6; count++) await app!.keyboard.press('ArrowLeft')
  await until(app!, (now) => now.format?.bold === true, 'bold under the caret')
})

test('sets font, size, colour and highlight on a selection', async ({ app }) => {
  await ready(app!)
  await app!.keyboard.press('ArrowDown')
  await app!.keyboard.press('Home')
  for (let count = 0; count < 4; count++) await app!.keyboard.press('Shift+ArrowRight')
  await until(app!, (now) => now.selection === 'This', 'a selected word')

  await choose(app!, 'font', 'Times New Roman')
  await until(app!, (now) => now.format?.font === 'Times New Roman', 'the font')
  await choose(app!, 'size', '20')
  await until(app!, (now) => now.format?.size === 20, 'the size')

  await app!.getByRole('button', { name: '文字颜色' }).click()
  await app!.getByRole('button', { name: '红色' }).first().click()
  const red = await until(app!, (now) => now.format?.color === 'FF0000', 'the colour')
  expect(red.selection).toBe('This')
  await app!.getByRole('button', { name: '文字颜色' }).click()
  await app!.getByRole('button', { name: '自动' }).click()
  await until(app!, (now) => now.format?.color === '000000', 'the automatic colour')

  const rectangles = (): Promise<number> =>
    app!.evaluate<number>('window.wordApp.elements(0, "rect")')
  const before = await rectangles()
  await app!.getByRole('button', { name: '高亮颜色' }).click()
  await app!.getByRole('button', { name: '黄色' }).click()
  for (let attempt = 0; attempt < 100 && (await rectangles()) === before; attempt++) {
    await new Promise((resolve) => setTimeout(resolve, 50))
  }
  expect(await rectangles()).toBe(before + 1)
})

test('makes lists, indents, spaces lines and breaks pages', async ({ app }) => {
  await ready(app!)
  await app!.keyboard.press('Control+End')
  const end = await until(app!, (now) => (now.caret?.para ?? 0) > 5, 'the end of the document')
  const last = end.caret!.para

  await app!.keyboard.press('Enter')
  await app!.getByRole('button', { name: '编号' }).click()
  await until(app!, (now) => now.format?.list === 'number', 'a numbered item')
  await app!.keyboard.type('one')
  await app!.keyboard.press('Enter')
  await app!.keyboard.type('two')
  const second = await until(app!, (now) => now.caret?.para === last + 2 && now.caret.offset === 3, 'the second item')
  expect(second.format?.list).toBe('number')
  // Enter in an empty item ends the list, and the paragraph stays.
  await app!.keyboard.press('Enter')
  await until(app!, (now) => now.caret?.para === last + 3 && now.format?.list === 'number', 'an empty third item')
  await app!.keyboard.press('Enter')
  const out = await until(app!, (now) => now.format?.list === null, 'the list ended')
  expect(out.caret?.para).toBe(last + 3)

  await app!.getByRole('button', { name: '项目符号' }).click()
  await until(app!, (now) => now.format?.list === 'bullet', 'a bullet item')
  await app!.getByRole('button', { name: '项目符号' }).click()
  await until(app!, (now) => now.format?.list === null, 'the bullet removed')

  await app!.getByRole('button', { name: '增加缩进' }).click()
  await until(app!, (now) => now.format?.indent === 36, 'the indent')
  await app!.getByRole('button', { name: '减少缩进' }).click()
  await until(app!, (now) => now.format?.indent === 0, 'the indent removed')

  await choose(app!, 'spacing', '1.5')
  await until(app!, (now) => now.format?.spacing === 1.5, 'the line spacing')

  await app!.keyboard.type('after the break')
  await app!.keyboard.press('Home')
  await app!.keyboard.press('Control+Enter')
  const broken = await until(app!, (now) => now.pages === 2, 'a second page')
  expect(broken.caret).toEqual({ para: last + 4, offset: 0 })
  // A paragraph typed after the break stays on the page: the break belongs to the first one only.
  await app!.keyboard.press('End')
  await app!.keyboard.press('Enter')
  await app!.keyboard.type('still page two')
  expect((await until(app!, (now) => now.caret?.para === last + 5, 'a paragraph after the break')).pages).toBe(2)
  // Undo goes back through the typed words and the new paragraph to the break itself.
  for (let count = 0; count < 12 && (await state(app!)).pages === 2; count++) {
    await app!.keyboard.press('Control+z')
    await new Promise((resolve) => setTimeout(resolve, 150))
  }
  await until(app!, (now) => now.pages === 1, 'the break undone')
})

test('finds and replaces', async ({ app }) => {
  await ready(app!)
  const before = await paragraphs(app!)
  const occurrences = before.join('\n').toLowerCase().split('the').length - 1
  expect(occurrences).toBeGreaterThan(3)

  await app!.keyboard.press('Control+f')
  await app!.getByLabel('查找内容').fill('the')
  const shown = (): Promise<string> => app!.evaluate<string>('document.getElementById("find-count").textContent')
  for (let attempt = 0; attempt < 100 && !(await shown()).includes(String(occurrences)); attempt++) {
    await new Promise((resolve) => setTimeout(resolve, 50))
  }
  expect(await shown()).toBe(`– / ${occurrences}`)

  // Enter steps through the occurrences from the caret on, and Shift+Enter steps back.
  await app!.keyboard.press('Enter')
  const first = await until(app!, (now) => now.selection.toLowerCase() === 'the', 'the first occurrence')
  expect(await shown()).toBe(`1 / ${occurrences}`)
  await app!.keyboard.press('Enter')
  await until(app!, (now) => now.selection.toLowerCase() === 'the' && now.caret?.offset !== first.caret?.offset, 'the second occurrence')
  expect(await shown()).toBe(`2 / ${occurrences}`)
  await app!.keyboard.press('Shift+Enter')
  await until(app!, (now) => now.caret?.offset === first.caret?.offset && now.caret?.para === first.caret?.para, 'the first again')

  // Replace changes the selected occurrence and moves on to the next.
  await app!.getByLabel('替换为').fill('THAT')
  await app!.getByRole('button', { name: '替换', exact: true }).click()
  await until(app!, (now) => now.canUndo && now.selection.toLowerCase() === 'the', 'one replacement')
  expect((await paragraphs(app!)).join('\n').split('THAT').length - 1).toBe(1)

  await app!.getByRole('button', { name: '全部替换' }).click()
  for (let attempt = 0; attempt < 100 && !(await shown()).startsWith('已替换'); attempt++) {
    await new Promise((resolve) => setTimeout(resolve, 50))
  }
  expect(await shown()).toBe(`已替换 ${occurrences - 1} 处`)
  const after = (await paragraphs(app!)).join('\n')
  expect(after.split('THAT').length - 1).toBe(occurrences)
  expect(after.toLowerCase().includes('the')).toBe(false)

  // The bar closes with Escape and the document takes the keyboard again; one undo per replace.
  await app!.keyboard.press('Escape')
  await app!.keyboard.press('Control+z')
  await app!.keyboard.press('Control+z')
  await until(app!, (now) => !now.canUndo, 'the replacements undone')
  expect(await paragraphs(app!)).toEqual(before)
})

test('moves by word, by screen and to the ends, and types a tab', async ({ app }) => {
  await ready(app!)
  // "Word for Dopame OS": the end of the first word, then of the second.
  await app!.keyboard.press('Control+ArrowRight')
  await until(app!, (now) => now.caret?.para === 0 && now.caret.offset === 4, 'the end of the first word')
  await app!.keyboard.press('Control+ArrowRight')
  await until(app!, (now) => now.caret?.offset === 8, 'the end of the second word')
  await app!.keyboard.press('Control+ArrowLeft')
  await until(app!, (now) => now.caret?.offset === 5, 'the start of the second word')
  await app!.keyboard.press('Control+Shift+ArrowRight')
  await until(app!, (now) => now.selection === 'for', 'a word selected by keyboard')

  await app!.keyboard.press('Control+End')
  const end = await until(app!, (now) => (now.caret?.para ?? 0) > 5, 'the end of the document')
  await app!.keyboard.press('PageUp')
  await until(app!, (now) => (now.caret?.para ?? 99) < end.caret!.para, 'a screen up')
  await app!.keyboard.press('Control+Home')
  await until(app!, (now) => now.caret?.para === 0 && now.caret.offset === 0, 'the start of the document')

  await app!.keyboard.press('Tab')
  await until(app!, (now) => now.caret?.offset === 1, 'the tab')
  expect((await paragraphs(app!))[0]).toBe('\tWord for Dopame OS')

  // The status bar counts words: Chinese characters one each, Latin words one each.
  const count = (): Promise<string> => app!.evaluate<string>('document.getElementById("count").textContent')
  for (let attempt = 0; attempt < 100 && !(await count()).startsWith('字数'); attempt++) {
    await new Promise((resolve) => setTimeout(resolve, 50))
  }
  const words = Number((await count()).replace(/\D/g, ''))
  expect(words).toBeGreaterThan(80)
  await app!.keyboard.insertText('中文字数')
  for (let attempt = 0; attempt < 100 && Number((await count()).replace(/\D/g, '')) !== words + 4; attempt++) {
    await new Promise((resolve) => setTimeout(resolve, 50))
  }
  expect(Number((await count()).replace(/\D/g, ''))).toBe(words + 4)
})

/** Window coordinates, in CSS pixels, of a point given in points on the first page. */
async function onPage(app: App, x: number, y: number): Promise<[number, number]> {
  const [left, top] = JSON.parse(
    await app.evaluate<string>(
      '(() => { const b = document.querySelector(".page").getBoundingClientRect(); return JSON.stringify([b.left, b.top]) })()',
    ),
  ) as [number, number]
  return [left + (x * 96) / 72, top + (y * 96) / 72]
}

type Pointer = App & {
  mouse: { click(x: number, y: number): Promise<unknown>; dblclick(x: number, y: number): Promise<unknown> }
}

// The welcome document: the heading on the first line, then a paragraph whose first line is
// about 100 points down the page, starting at the 72 point margin with the word "This".
const BODY_LINE_Y = 100
const HEADING_Y = 80
const MARGIN_X = 72

test('strikes text through, raises and lowers it, and clears formats', async ({ app }) => {
  await ready(app!)
  const lines = (): Promise<number> => app!.evaluate<number>('window.wordApp.elements(0, "line")')
  const plain = await lines()
  await (app as unknown as Pointer).mouse.dblclick(...(await onPage(app!, MARGIN_X + 8, BODY_LINE_Y)))
  await until(app!, (now) => now.selection === 'This', 'a selected word')

  await app!.getByRole('button', { name: '删除线' }).click()
  await until(app!, (now) => now.format?.strike === true, 'the strikethrough')
  // The line through the word is drawn.
  expect(await lines()).toBe(plain + 1)

  await app!.getByRole('button', { name: '上标' }).click()
  const raised = await until(app!, (now) => now.format?.script === 'super', 'the superscript')
  // A raised character is drawn smaller than the 11 point text around it.
  expect(raised.format!.size!).toBeLessThan(9)
  await app!.getByRole('button', { name: '下标' }).click()
  await until(app!, (now) => now.format?.script === 'sub', 'the subscript')
  await app!.getByRole('button', { name: '下标' }).click()
  await until(app!, (now) => now.format?.script === null && now.format.size === 11, 'the text back on its line')

  await app!.getByRole('button', { name: '加粗' }).click()
  await until(app!, (now) => now.format?.bold === true, 'the bold')
  await app!.getByRole('button', { name: '清除格式' }).click()
  const cleared = await until(app!, (now) => now.format?.strike === false && now.format.bold === false, 'the cleared formats')
  expect(cleared.selection).toBe('This')
  expect(await lines()).toBe(plain)
})

test('the format painter gives clicked text the format it picked up', async ({ app }) => {
  await ready(app!)
  const pointer = app as unknown as Pointer
  // Pick the heading's look up: bold, 16 points, blue.
  await pointer.mouse.dblclick(...(await onPage(app!, MARGIN_X + 8, HEADING_Y)))
  const source = await until(app!, (now) => now.selection === 'Word', 'a word of the heading')
  expect(source.format).toMatchObject({ bold: true, size: 16 })
  await app!.getByRole('button', { name: '格式刷' }).click()
  await until(app!, (now) => now.format?.painting === true, 'the painter holding a format')

  // A click in a word of the body paints that word, and the painter lets go.
  await pointer.mouse.click(...(await onPage(app!, MARGIN_X + 8, BODY_LINE_Y)))
  const painted = await until(app!, (now) => now.format?.painting === false && now.format.bold === true, 'the painted word')
  expect(painted.format).toMatchObject({ size: 16, color: source.format!.color })
  expect(painted.caret!.para).toBe(1)
  // Only that word: the next one is as it was.
  await app!.keyboard.press('Control+ArrowRight')
  await app!.keyboard.press('Control+ArrowRight')
  const next = await until(app!, (now) => (now.caret?.offset ?? 0) > 5, 'the next word')
  expect(next.format).toMatchObject({ bold: false, size: 11 })
  // One undo takes the whole look off again.
  await app!.keyboard.press('Control+z')
  await until(app!, (now) => !now.canUndo, 'the paint undone')
})

test('sets the first-line indent and the space around a paragraph', async ({ app }) => {
  await ready(app!)
  await app!.keyboard.press('Control+End')
  const start = await until(app!, (now) => (now.caret?.para ?? 0) > 5, 'the last paragraph')
  expect(start.format).toMatchObject({ firstLine: 0, before: null, after: null })

  await app!.getByRole('button', { name: '段落设置' }).click()
  // The dialog opens with what the paragraph has now.
  expect(await app!.evaluate<string>('document.getElementById("first-line").value')).toBe('0')
  await app!.getByLabel('首行缩进（字符）').fill('2')
  await app!.getByLabel('段前（磅）').fill('6')
  await app!.getByLabel('段后（磅）').fill('12')
  await app!.getByRole('button', { name: '确定' }).click()
  // Two characters of 11 point text.
  const set = await until(app!, (now) => now.format?.firstLine === 22, 'the first-line indent')
  expect(set.format).toMatchObject({ before: 6, after: 12 })

  // An empty spacing field gives that side back to the style; the indent stays.
  await app!.getByRole('button', { name: '段落设置' }).click()
  expect(await app!.evaluate<string>('document.getElementById("first-line").value')).toBe('2')
  await app!.getByLabel('段前（磅）').fill('')
  await app!.getByRole('button', { name: '确定' }).click()
  const back = await until(app!, (now) => now.format?.before === null, 'the space before given back')
  expect(back.format).toMatchObject({ firstLine: 22, after: 12 })
})
