// End-to-end checks of pictures in the text: typing beside one, and selecting, resizing,
// dragging and deleting one. Run with: acde test

import { join } from 'node:path'
import { expect, test } from 'acde/test'

interface State {
  dirty: boolean
  pages: number
  caret: { para: number; offset: number } | null
  picture: { para: number; index: number } | null
  canUndo: boolean
}

type App = {
  evaluate<T>(code: string): Promise<T>
  mouse: {
    click(x: number, y: number): Promise<unknown>
    move(x: number, y: number, options?: { steps: number }): Promise<unknown>
    down(): Promise<unknown>
    up(): Promise<unknown>
  }
}

async function state(app: App): Promise<State> {
  return JSON.parse(await app.evaluate<string>('JSON.stringify(window.wordApp?.state() ?? null)')) as State
}

async function read<T>(app: App, call: string): Promise<T> {
  return JSON.parse(await app.evaluate<string>(`JSON.stringify(window.wordApp.${call})`)) as T
}

async function until(app: App, check: (state: State) => boolean, what: string): Promise<State> {
  for (let attempt = 0; attempt < 100; attempt++) {
    const now = await state(app)
    if (now && check(now)) return now
    await new Promise((resolve) => setTimeout(resolve, 50))
  }
  throw new Error(`timed out waiting for ${what}: ${JSON.stringify(await state(app))}`)
}

/** Window coordinates, in CSS pixels, of a point given in points on the first page. */
async function onPage(app: App, x: number, y: number): Promise<[number, number]> {
  const [left, top] = JSON.parse(
    await app.evaluate<string>(
      '(() => { const b = document.querySelector(".page").getBoundingClientRect(); return JSON.stringify([b.left, b.top]) })()',
    ),
  ) as [number, number]
  return [left + (x * 96) / 72, top + (y * 96) / 72]
}

type Box = [x: number, y: number, width: number, height: number]

/**
 * Puts the test picture after the heading of the welcome document, as paragraph 1, and
 * makes it small enough for the whole document to stay on the first page.
 */
async function withPicture(
  app: App,
  acde: { chooseFiles(path: string): Promise<unknown> },
  press: () => Promise<unknown>,
): Promise<Box> {
  await until(app, (now) => now.pages === 1 && now.caret !== null, 'the welcome document')
  await acde.chooseFiles(join(import.meta.dir, '..', 'docs', 'screenshots', 'format-selection.png'))
  await press()
  await until(app, (now) => now.dirty, 'the picture')
  const [[x, y, width, height]] = await read<Box[]>(app, 'pictureBoxes(0)')
  // Drag the bottom right corner in: a press on the picture selects it and shows the handle.
  await app.mouse.click(...(await onPage(app, x + width / 2, y + height / 2)))
  await until(app, (now) => now.picture?.para === 1, 'the picture selected')
  await app.mouse.move(...(await onPage(app, x + width, y + height)))
  await app.mouse.down()
  await app.mouse.move(...(await onPage(app, x + 120, y + (height * 120) / width)), { steps: 6 })
  await app.mouse.up()
  for (let attempt = 0; attempt < 100; attempt++) {
    const [box] = await read<Box[]>(app, 'pictureBoxes(0)')
    if (Math.abs(box[2] - 120) < 2) return box
    await new Promise((resolve) => setTimeout(resolve, 50))
  }
  throw new Error(`the picture was not resized: ${JSON.stringify(await read(app, 'pictureBoxes(0)'))}`)
}

test('resizes a picture by a corner and types after it', async ({ app, acde }) => {
  const insert = (): Promise<unknown> => app!.getByRole('button', { name: '插入图片' }).click()
  const [x, y, width, height] = await withPicture(app!, acde, insert)
  // The picture kept its proportions, and stays selected in its new size.
  const [[para, drawnWidth, drawnHeight]] = await read<[number, number, number][]>(app!, 'pictures()')
  expect(para).toBe(1)
  expect(Math.abs(drawnWidth / drawnHeight - width / height)).toBeLessThan(0.01)
  expect((await state(app!)).picture).toEqual({ para: 1, index: 0 })

  // A click beside the picture, at half its height, puts the caret after it.
  await app!.mouse.click(...(await onPage(app!, x + width + 40, y + height / 2)))
  await until(app!, (now) => now.picture === null && now.caret?.para === 1 && now.caret.offset === 0, 'the caret after the picture')
  await app!.keyboard.type('caption')
  await until(app!, (now) => now.caret?.offset === 7, 'the typed text')
  expect((await read<string[]>(app!, 'paragraphs()'))[1]).toBe('caption')
  // The text is drawn after the picture, on its line, and the picture is still there.
  expect(await read<Box[]>(app!, 'pictureBoxes(0)')).toEqual([[x, y, width, height]])

  // With the picture selected, typing goes right after it too.
  await app!.mouse.click(...(await onPage(app!, x + width / 2, y + height / 2)))
  await until(app!, (now) => now.picture?.para === 1, 'the picture selected again')
  await app!.keyboard.type('A ')
  await until(app!, (now) => now.caret?.offset === 2 && now.picture === null, 'the text after the picture')
  expect((await read<string[]>(app!, 'paragraphs()'))[1]).toBe('A caption')
})

test('Enter after a picture starts a paragraph below it, and Backspace takes the picture out', async ({ app, acde }) => {
  const insert = (): Promise<unknown> => app!.getByRole('button', { name: '插入图片' }).click()
  const [x, y, width, height] = await withPicture(app!, acde, insert)
  const before = await read<string[]>(app!, 'paragraphs()')

  await app!.mouse.click(...(await onPage(app!, x + width + 40, y + height / 2)))
  await until(app!, (now) => now.picture === null && now.caret?.para === 1 && now.caret.offset === 0, 'the caret after the picture')
  await app!.keyboard.press('Enter')
  await until(app!, (now) => now.caret?.para === 2 && now.caret.offset === 0, 'the new paragraph')
  // The picture stays where it was, in its own paragraph; the new one is below it.
  expect(await read<Box[]>(app!, 'pictureBoxes(0)')).toEqual([[x, y, width, height]])
  expect((await read<[number, number, number][]>(app!, 'pictures()')).map(([para]) => para)).toEqual([1])
  await app!.keyboard.type('below')
  await until(app!, (now) => now.caret?.para === 2 && now.caret.offset === 5, 'the text below the picture')
  const split = await read<string[]>(app!, 'paragraphs()')
  expect(split.length).toBe(before.length + 1)
  expect(split.slice(1, 3)).toEqual(['', 'below'])

  // Backspace with the caret right after the picture deletes the picture, not a paragraph mark.
  await app!.mouse.click(...(await onPage(app!, x + width + 40, y + height / 2)))
  await until(app!, (now) => now.caret?.para === 1 && now.caret.offset === 0, 'the caret after the picture again')
  await app!.keyboard.press('Backspace')
  for (let attempt = 0; attempt < 100 && (await read<number>(app!, 'images(0)')) > 0; attempt++) {
    await new Promise((resolve) => setTimeout(resolve, 50))
  }
  expect(await read<number>(app!, 'images(0)')).toBe(0)
  expect(await read<string[]>(app!, 'paragraphs()')).toEqual(split)
})

test('drags a picture to another place and deletes it', async ({ app, acde }) => {
  const insert = (): Promise<unknown> => app!.getByRole('button', { name: '插入图片' }).click()
  const [x, y, width, height] = await withPicture(app!, acde, insert)
  const before = await read<string[]>(app!, 'paragraphs()')
  const target = before.findIndex((text) => text.startsWith('Save with'))
  expect(target).toBeGreaterThan(1)

  // Drop it on the lower half of a later paragraph: it lands after that paragraph.
  const [dropX, dropY] = JSON.parse(
    await app!.evaluate<string>(
      '(() => { const p = [...document.querySelectorAll(".page-text p")].find((p) => p.textContent.startsWith("Save with")); const b = p.getBoundingClientRect(); return JSON.stringify([b.left + 60, b.bottom - 2]) })()',
    ),
  ) as [number, number]
  await app!.mouse.move(...(await onPage(app!, x + width / 2, y + height / 2)))
  await app!.mouse.down()
  await app!.mouse.move(dropX, dropY, { steps: 8 })
  await app!.mouse.up()
  // The paragraph it left empty went with it, so the one it was dropped on is now one earlier.
  await until(app!, (now) => now.picture?.para === target, 'the picture moved')
  const after = await read<string[]>(app!, 'paragraphs()')
  expect(after.length).toBe(before.length)
  expect(after[target - 1]).toBe(before[target])
  expect((await read<[number, number, number][]>(app!, 'pictures()')).map(([para]) => para)).toEqual([target])

  // One undo puts it back where it was.
  await app!.keyboard.press('Control+z')
  await until(app!, (now) => now.picture === null, 'the move undone')
  expect(await read<string[]>(app!, 'paragraphs()')).toEqual(before)
  expect((await read<[number, number, number][]>(app!, 'pictures()')).map(([para]) => para)).toEqual([1])

  // Delete takes a selected picture out and leaves its paragraph.
  await app!.mouse.click(...(await onPage(app!, x + width / 2, y + height / 2)))
  await until(app!, (now) => now.picture?.para === 1, 'the picture selected')
  await app!.keyboard.press('Delete')
  await until(app!, (now) => now.picture === null, 'the picture deleted')
  expect(await read<number>(app!, 'images(0)')).toBe(0)
  expect((await read<string[]>(app!, 'paragraphs()')).length).toBe(before.length)
})
