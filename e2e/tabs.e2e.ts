// End-to-end checks of the tab bar: several documents open in one window. Run with: acde test

import { expect, test } from 'acde/test'

interface State {
  name: string
  dirty: boolean
  pages: number
  caret: { para: number; offset: number } | null
  canUndo: boolean
}

interface TabState {
  name: string
  dirty: boolean
  active: boolean
}

type App = { evaluate<T>(code: string): Promise<T> }

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

test('keeps several documents open, each with its own caret and undo history', async ({ app }) => {
  await until(app!, (now) => now.pages === 1 && now.caret !== null, 'the welcome document')
  await app!.keyboard.type('one')
  await until(app!, (now) => now.caret?.offset === 3, 'text in the first document')

  // A new document opens in a tab of its own, next to the first, and nothing is asked.
  await app!.getByRole('button', { name: '新建', exact: true }).click()
  await until(app!, (now) => now.name === '未命名文档 2' && now.caret !== null, 'the second document')
  expect(await read<TabState[]>(app!, 'tabs()')).toEqual([
    { name: '未命名文档', dirty: true, active: false },
    { name: '未命名文档 2', dirty: false, active: true },
  ])
  expect(await read<string[]>(app!, 'paragraphs()')).toEqual([''])
  await app!.keyboard.type('two')
  await until(app!, (now) => now.caret?.offset === 3 && now.dirty, 'text in the second document')

  // Back in the first document the caret is where it was left, and typing goes on from there.
  await app!.getByRole('tab', { name: '未命名文档', exact: true }).click()
  const first = await until(app!, (now) => now.name === '未命名文档', 'the first document again')
  expect(first.caret).toEqual({ para: 0, offset: 3 })
  expect((await read<string[]>(app!, 'text()'))[0]).toBe('oneWord for Dopame OS')
  await app!.keyboard.type('!')
  await until(app!, (now) => now.caret?.offset === 4, 'more text in the first document')
  expect((await read<string[]>(app!, 'text()'))[0]).toBe('one!Word for Dopame OS')

  // Undo in the second document takes back its own text and leaves the first alone.
  await app!.getByRole('tab', { name: '未命名文档 2' }).click()
  await until(app!, (now) => now.name === '未命名文档 2' && now.caret?.offset === 3, 'the second document again')
  await app!.keyboard.press('Control+z')
  await until(app!, (now) => !now.canUndo, 'the undo in the second document')
  expect(await read<string[]>(app!, 'paragraphs()')).toEqual([''])
  await app!.getByRole('tab', { name: '未命名文档', exact: true }).click()
  await until(app!, (now) => now.name === '未命名文档', 'the first document once more')
  expect((await read<string[]>(app!, 'text()'))[0]).toBe('one!Word for Dopame OS')
})

test('closes a tab and shows the one beside it', async ({ app }) => {
  await until(app!, (now) => now.pages === 1 && now.caret !== null, 'the welcome document')
  await app!.getByRole('button', { name: '新标签页' }).click()
  await until(app!, (now) => now.name === '未命名文档 2' && now.caret !== null, 'the second document')
  await app!.keyboard.type('draft')
  await until(app!, (now) => now.dirty, 'an unsaved change')

  // Closing a tab that is not on screen shows it first, then asks about its changes.
  await app!.getByRole('tab', { name: '未命名文档', exact: true }).click()
  await until(app!, (now) => now.name === '未命名文档', 'the first document')
  await app!.getByRole('button', { name: '关闭 未命名文档 2' }).click()
  await until(app!, (now) => now.name === '未命名文档 2', 'the tab that is being closed')
  expect(await app!.evaluate<boolean>('document.getElementById("unsaved").hidden')).toBe(false)
  await app!.getByRole('button', { name: '不保存' }).click()
  await until(app!, (now) => now.name === '未命名文档', 'the first document after the close')
  expect(await read<TabState[]>(app!, 'tabs()')).toEqual([{ name: '未命名文档', dirty: false, active: true }])
  expect((await read<string[]>(app!, 'text()'))[0]).toBe('Word for Dopame OS')
})
