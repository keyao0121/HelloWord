/// <reference types="@dopame/acde-api/shell" />
import { apps } from 'acde:apps'
import { t, translate } from './i18n.js'

translate()
const dock = /** @type {HTMLElement} */ (document.querySelector('#dock'))
dock.setAttribute('aria-label', t('dock'))
const list = await apps.list()
if (list.length === 0) dock.textContent = t('empty')
for (const app of list) {
  const button = document.createElement('button')
  button.type = 'button'
  button.textContent = app.label
  button.addEventListener('click', () => void apps.open(app.id))
  dock.append(button)
}
