// Requests as another app or an agent sends them: through acde:open, in a desktop session.
// The sender here is the desktop page of the small desktop in e2e/fixtures/desk.
// Run with: acde test

import { join } from 'node:path'
import { Acde } from 'acde/driver'
import { expect, test } from 'acde/test'

const WORD = join(import.meta.dir, '..')
const DESK = join(import.meta.dir, 'fixtures', 'desk')

interface State {
  name: string
  dirty: boolean
}

/** Sends a request to Word from the desktop page and resolves to what open() answered. */
function send(acde: Acde, request: Record<string, unknown>): Promise<string> {
  const code = `import('acde:open').then(({ open }) => open('word', ${JSON.stringify(request)})).then(
    (opened) => 'opened ' + typeof opened.window,
    (error) => 'rejected ' + error.status + ' ' + error.code,
  )`
  return acde.evaluate<string>('desktop', code)
}

async function until(acde: Acde, check: (state: State, text: string[]) => boolean, what: string): Promise<string[]> {
  let last = ''
  for (let attempt = 0; attempt < 200; attempt++) {
    try {
      last = await acde.evaluate<string>(
        'word',
        'JSON.stringify(window.wordApp ? [window.wordApp.state(), window.wordApp.text()] : null)',
      )
      const parsed = JSON.parse(last) as [State, string[]] | null
      if (parsed && check(parsed[0], parsed[1])) return parsed[1]
    } catch {
      // The window is not open yet.
    }
    await new Promise((resolve) => setTimeout(resolve, 100))
  }
  throw new Error(`timed out waiting for ${what}: ${last}`)
}

test('another app drives Word with requests', async () => {
  const acde = await Acde.launch({ dirs: [WORD, DESK], desktop: DESK, headless: true })
  try {
    // The first request opens the window; it is held until the app is ready for it.
    expect(await send(acde, { kind: 'new-document', title: '周报', text: '# Weekly report\nDone: the editor.\n## Next\nMore to do.' })).toBe(
      'opened number',
    )
    const created = await until(acde, (state) => state.name === '周报', 'the requested document')
    expect(created.slice(0, 4)).toEqual(['Weekly report', 'Done: the editor.', 'Next', 'More to do.'])

    // Later requests reach the same window.
    await send(acde, { kind: 'insert-text', text: ' Shipped.', at: 'end' })
    await until(acde, (_, text) => text[3] === 'More to do. Shipped.', 'the inserted text')
    await send(acde, { kind: 'replace-text', find: 'editor', replace: 'word processor' })
    const replaced = await until(acde, (_, text) => text[1].includes('word processor'), 'the replacement')
    expect(replaced[1]).toBe('Done: the word processor.')

    // What Word shows is in the element tree, which is how an agent reads the result.
    const tree = (await acde.snapshot()).text
    expect(tree).toContain('heading "Weekly report" [level=1]')
    expect(tree).toContain('Done: the word processor.')
    expect(tree).toContain('已按请求替换 1 处')

    // The runtime turns away a kind the app does not list.
    expect(await send(acde, { kind: 'format-disk' })).toBe('rejected 409 not_accepted')
  } finally {
    await acde.close()
  }
})

test('an agent reads the document from the element tree and edits it by ref', async ({ app, acde }) => {
  const ready = async (): Promise<string[]> => JSON.parse(await app!.evaluate<string>('JSON.stringify(window.wordApp?.state().caret ? window.wordApp.text() : null)')) as string[]
  for (let attempt = 0; attempt < 100 && !(await ready()); attempt++) await new Promise((resolve) => setTimeout(resolve, 50))

  // The canvas has no text of its own: headings and paragraphs are in the tree all the same.
  const tree = (await acde.snapshot()).text
  expect(tree).toContain('heading "Word for Dopame OS" [level=1]')
  expect(tree).toContain('heading "Try it" [level=2]')
  expect(tree).toContain('button "加粗"')

  // Clicking a paragraph by its ref puts the caret in it; the text box takes what is typed.
  const paragraph = /text "Click anywhere[^\n]*\[(w\d+\.\d+)\]/.exec(tree)![1]
  expect((await acde.act(paragraph, 'click')).ok).toBe(true)
  const box = /textbox "文档输入"[^\n]*\[(w\d+\.\d+)\]/.exec((await acde.snapshot()).text)![1]
  await acde.act(box, 'key', 'End')
  await acde.act(box, 'type', ' [agent]')
  let text: string[] = []
  for (let attempt = 0; attempt < 100; attempt++) {
    text = await ready()
    if (text[3].endsWith('[agent]')) break
    await new Promise((resolve) => setTimeout(resolve, 50))
  }
  expect(text[3]).toBe('Click anywhere in the text and type. Enter splits a paragraph, Backspace joins two. [agent]')
  expect((await acde.snapshot()).text).toContain('Backspace joins two. [agent]')
})
