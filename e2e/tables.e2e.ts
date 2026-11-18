// End-to-end checks of tables, pictures and Chinese text. Run with: acde test

import { existsSync } from 'node:fs'
import { join } from 'node:path'
import { expect, test } from 'acde/test'

interface State {
  name: string
  dirty: boolean
  pages: number
  caret: { para: number; offset: number } | null
  selection: string
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

test('inserts a table and edits its cells', async ({ app, acde, home }) => {
  await ready(app!)
  const before = await paragraphs(app!)

  await app!.getByRole('button', { name: '插入表格' }).click()
  await app!.getByLabel('行数').fill('2')
  await app!.getByLabel('列数').fill('2')
  await app!.getByRole('button', { name: '插入', exact: true }).click()
  // The caret lands in the first cell, which is the paragraph after the one it was in.
  await until(app!, (now) => now.caret?.para === 1, 'the caret in the first cell')
  expect((await paragraphs(app!)).length).toBe(before.length + 4)

  // Tab goes from cell to cell.
  for (const [index, cell] of ['A1', 'B1', 'A2', 'B2'].entries()) {
    if (index > 0) await app!.keyboard.press('Tab')
    await app!.keyboard.type(cell)
    await until(app!, (now) => now.caret?.para === index + 1 && now.caret.offset === 2, `the text of cell ${cell}`)
  }
  expect((await paragraphs(app!)).slice(1, 5)).toEqual(['A1', 'B1', 'A2', 'B2'])

  // A cell is reached by a click on its text, here the way an agent clicks: by ref.
  const tree = (await acde.snapshot()).text
  const cell = /text "B1"[^\n]*\[(w\d+\.\d+)\]/.exec(tree)![1]
  await acde.act(cell, 'click')
  await until(app!, (now) => now.caret?.para === 2, 'the caret in cell B1')

  // Enter adds a paragraph inside the cell and Backspace joins it back; the other cells stay.
  await app!.keyboard.press('End')
  await app!.keyboard.press('Enter')
  await app!.keyboard.type('more')
  await until(app!, (now) => now.caret?.para === 3 && now.caret.offset === 4, 'a second paragraph in the cell')
  expect((await paragraphs(app!)).slice(1, 6)).toEqual(['A1', 'B1', 'more', 'A2', 'B2'])
  await app!.keyboard.press('Home')
  await app!.keyboard.press('Backspace')
  await until(app!, (now) => now.caret?.para === 2 && now.caret.offset === 2, 'the paragraphs joined')
  expect((await paragraphs(app!)).slice(1, 5)).toEqual(['A1', 'B1more', 'A2', 'B2'])

  // Backspace at the start of a cell does not pull the cell before it in.
  await app!.keyboard.press('Home')
  await app!.keyboard.press('Backspace')
  expect((await paragraphs(app!)).slice(1, 5)).toEqual(['A1', 'B1more', 'A2', 'B2'])

  // The table survives a save and a reopen.
  const docx = join(home, 'table.docx')
  await acde.chooseSaveFile(docx)
  await app!.keyboard.press('Control+s')
  await until(app!, (now) => !now.dirty, 'the save')
  expect(existsSync(docx)).toBe(true)
  const saved = await paragraphs(app!)
  // Closing the only tab leaves an empty document; the file then opens in its place.
  await app!.getByRole('button', { name: '关闭 table.docx' }).click()
  await until(app!, (now) => now.name === '未命名文档', 'an empty document')
  await acde.chooseFiles(docx)
  await app!.getByRole('button', { name: '打开' }).click()
  await until(app!, (now) => now.name === 'table.docx', 'the reopened file')
  expect(await paragraphs(app!)).toEqual(saved)
})

test('undoes a table step by step', async ({ app }) => {
  await ready(app!)
  const before = await paragraphs(app!)
  await app!.getByRole('button', { name: '插入表格' }).click()
  await app!.getByRole('button', { name: '插入', exact: true }).click()
  await until(app!, (now) => now.caret?.para === 1, 'the caret in the first cell')
  // The dialog offers three rows of three columns.
  expect((await paragraphs(app!)).length).toBe(before.length + 9)
  await app!.keyboard.type('cell')
  await until(app!, (now) => now.caret?.offset === 4, 'the typed text')
  await app!.keyboard.press('Control+z')
  await until(app!, (now) => now.caret?.para === 1 && now.caret.offset === 0, 'the text undone')
  await app!.keyboard.press('Control+z')
  await until(app!, (now) => !now.canUndo, 'the table undone')
  expect(await paragraphs(app!)).toEqual(before)
})

test('inserts a picture', async ({ app, acde }) => {
  await ready(app!)
  const before = await paragraphs(app!)
  expect(await app!.evaluate<number>('window.wordApp.images(0)')).toBe(0)
  await acde.chooseFiles(join(import.meta.dir, '..', 'docs', 'screenshots', 'format-selection.png'))
  await app!.getByRole('button', { name: '插入图片' }).click()
  await until(app!, (now) => now.dirty, 'the picture')
  expect(await app!.evaluate<number>('window.wordApp.images(0)')).toBe(1)
  // The picture has a paragraph of its own, after the caret's.
  expect((await paragraphs(app!)).length).toBe(before.length + 1)
  await app!.keyboard.press('Control+z')
  await until(app!, (now) => !now.canUndo, 'the picture undone')
  expect(await app!.evaluate<number>('window.wordApp.images(0)')).toBe(0)
})

test('lays Chinese text out at full width', async ({ app }) => {
  await ready(app!)
  const caretLeft = (): Promise<number> => app!.evaluate<number>('parseFloat(document.querySelector(".caret").style.left)')
  const start = await caretLeft()
  // Characters the old subset font did not have: each one took half its width and overlapped.
  const text = '排版引擎需要重新计算每一行'
  await app!.keyboard.insertText(text)
  await until(app!, (now) => now.caret?.offset === text.length, 'the Chinese text')
  // The heading is 16 points: twelve characters are twelve ems wide, at 96 pixels to 72 points.
  const width = ((await caretLeft()) - start) * (72 / 96)
  expect(width).toBeGreaterThan(text.length * 16 * 0.95)
  expect(width).toBeLessThan(text.length * 16 * 1.05)
})

test('adds and removes rows and columns, and merges and splits cells', async ({ app }) => {
  await ready(app!)
  const before = await paragraphs(app!)
  const operate = async (command: string): Promise<void> => {
    await app!.getByRole('button', { name: '表格操作' }).click()
    await app!.getByRole('menuitem', { name: command }).click()
  }

  await app!.getByRole('button', { name: '插入表格' }).click()
  await app!.getByLabel('行数').fill('2')
  await app!.getByLabel('列数').fill('2')
  await app!.getByRole('button', { name: '插入', exact: true }).click()
  await until(app!, (now) => now.caret?.para === 1, 'the caret in the first cell')
  for (const [index, cell] of ['A1', 'B1', 'A2', 'B2'].entries()) {
    if (index > 0) await app!.keyboard.press('Tab')
    await app!.keyboard.type(cell)
    await until(app!, (now) => now.caret?.para === index + 1 && now.caret.offset === 2, `the text of cell ${cell}`)
  }
  const cells = async (): Promise<string[]> => (await paragraphs(app!)).slice(1, -before.length + 1)

  // The caret is in B2. A row below it: the caret goes to the new row, in the same column.
  await operate('在下方插入行')
  await until(app!, (now) => now.caret?.para === 6 && now.caret.offset === 0, 'the new row')
  expect(await cells()).toEqual(['A1', 'B1', 'A2', 'B2', '', ''])
  await app!.keyboard.type('B3')
  await until(app!, (now) => now.caret?.offset === 2, 'text in the new row')

  // A column to its right, in every row.
  await operate('在右侧插入列')
  await until(app!, (now) => now.caret?.para === 9 && now.caret.offset === 0, 'the new column')
  expect(await cells()).toEqual(['A1', 'B1', '', 'A2', 'B2', '', '', 'B3', ''])

  // Taking the column and the row out again leaves the first four cells.
  await operate('删除列')
  await until(app!, (now) => now.caret?.para === 6, 'the column gone')
  expect(await cells()).toEqual(['A1', 'B1', 'A2', 'B2', '', 'B3'])
  await operate('删除行')
  await until(app!, (now) => now.caret?.para === 4, 'the row gone')
  expect(await cells()).toEqual(['A1', 'B1', 'A2', 'B2'])

  // Merging wants more than one cell selected, and says so.
  await operate('合并单元格')
  for (let attempt = 0; attempt < 40; attempt++) {
    if ((await app!.evaluate<string>('document.getElementById("message").textContent')).includes('选中')) break
    await new Promise((resolve) => setTimeout(resolve, 50))
  }
  expect(await app!.evaluate<string>('document.getElementById("message").textContent')).toContain('选中')

  // Select from A2 into B2 and merge: one cell holds what both held.
  await app!.keyboard.press('Shift+Tab')
  await until(app!, (now) => now.caret?.para === 3 && now.caret.offset === 0, 'the caret in A2')
  for (let step = 0; step < 4; step++) await app!.keyboard.press('Shift+ArrowRight')
  await until(app!, (now) => now.caret?.para === 4 && now.caret.offset === 1 && now.selection.length > 0, 'a selection into B2')
  await operate('合并单元格')
  await until(app!, (now) => now.caret?.para === 3 && now.selection === '', 'the merged cell')
  expect(await cells()).toEqual(['A1', 'B1', 'A2', 'B2'])
  // A2 and B2 are two paragraphs of one cell now: the second follows the first as its sibling.
  const siblings = (): Promise<boolean> => app!.evaluate<boolean>('window.wordApp.sibling(3)')
  expect(await siblings()).toBe(true)

  // Splitting gives the row its second cell back, empty; the content stays in the first.
  await operate('拆分单元格')
  await until(app!, (now) => now.caret?.para === 3 && now.canUndo, 'the split cell')
  for (let attempt = 0; attempt < 40 && (await cells()).length < 5; attempt++) {
    await new Promise((resolve) => setTimeout(resolve, 50))
  }
  expect(await cells()).toEqual(['A1', 'B1', 'A2', 'B2', ''])
  expect(await siblings()).toBe(true)

  // Every step is its own undo step, back to the document before the table.
  for (let step = 0; step < 40 && (await state(app!)).canUndo; step++) await app!.keyboard.press('Control+z')
  await until(app!, (now) => !now.canUndo, 'everything undone')
  expect(await paragraphs(app!)).toEqual(before)
})
