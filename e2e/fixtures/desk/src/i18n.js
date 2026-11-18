import en from './locales/en.js'
import zh from './locales/zh.js'

const catalogs = { en, zh }

function pick(tag) {
  if (tag in catalogs) return tag
  try {
    const { language } = new Intl.Locale(tag)
    if (language in catalogs) return language
  } catch {}
  return 'en'
}

export const locale = pick(navigator.language)
const messages = { ...en, ...catalogs[locale] }

export function t(key, values = {}) {
  const text = messages[key] ?? key
  return text.replace(/\{(\w+)\}/g, (all, name) => (name in values ? String(values[name]) : all))
}

export function translate(root = document) {
  document.documentElement.lang = locale
  for (const element of root.querySelectorAll('[data-i18n]')) element.textContent = t(element.dataset.i18n)
}
