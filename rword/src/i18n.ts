/* UI language for the React office apps, using the suite's existing dictionaries.
 * public/common/lang.js resolves the same choice before Quire paints; this mirrors it
 * without loading the original script into the React bundle. */
const LANGUAGES = ['en', 'fr', 'de', 'zh-CN', 'zh-Hant', 'ja', 'ko', 'qps'] as const
type Language = (typeof LANGUAGES)[number]
type Dictionary = Readonly<Record<string, string>>

const sources = import.meta.glob<string>('../../public/common/i18n/*.json', {
  eager: true,
  query: '?raw',
  import: 'default',
})

const dictionaries = new Map<string, Dictionary>()
for (const [path, source] of Object.entries(sources)) {
  const name = path.split('/').pop()?.replace(/\.json$/, '')
  if (!name) continue
  try {
    dictionaries.set(name, JSON.parse(source) as Dictionary)
  } catch {
    // An unreadable catalog must never keep the interface from painting.
  }
}

const isLanguage = (value: string): value is Language =>
  (LANGUAGES as readonly string[]).includes(value)

export function normalizeLanguage(tag: string | null | undefined): Language | null {
  const value = String(tag || '').trim().toLowerCase()
  if (/^zh-(?:tw|hk|mo|hant)(?:-|$)/.test(value)) return 'zh-Hant'
  if (value === 'zh' || value.startsWith('zh-')) return 'zh-CN'
  if (value === 'qps' || value === 'qps-ploc') return 'qps'
  const base = value.split('-')[0]
  return isLanguage(base) ? base : null
}

function storedLanguage(): string {
  const global = window as unknown as {
    VO?: { settings?: { get?: (key: string, fallback?: string | null) => string | null } }
    __VO_SETTINGS__?: Record<string, string>
  }
  const desk = global.__VO_SETTINGS__
  if (desk && 'vibeoffice.lang' in desk) return desk['vibeoffice.lang']
  const get = global.VO?.settings?.get
  if (typeof get === 'function') return get('vibeoffice.lang') || 'auto'
  try {
    return localStorage.getItem('vibeoffice.lang') || 'auto'
  } catch {
    return 'auto'
  }
}

function resolveLanguage(): Language {
  const root = document.documentElement
  const current = normalizeLanguage(root.getAttribute('lang'))
  if (current) return current
  const choice = normalizeLanguage(new URLSearchParams(location.search).get('lang') || storedLanguage())
  if (choice && choice !== 'qps' && choice !== 'en') return choice
  if (choice === 'qps') return 'qps'
  for (const candidate of navigator.languages || [navigator.language]) {
    const normalized = normalizeLanguage(candidate)
    if (normalized && normalized !== 'qps') return normalized
  }
  return 'en'
}

let language: Language | null = null

export function initLanguage(): Language {
  if (language) return language
  language = resolveLanguage()
  document.documentElement.lang = language === 'qps' ? 'qps-ploc' : language
  return language
}

if (typeof document !== 'undefined') initLanguage()

/** With a context suffix ("Line|chart"), the translated text is the part before it. */
function withoutContext(text: string): string {
  return text.replace(/(\S)\|[A-Za-z][A-Za-z -]*$/, '$1')
}

function placeholderMarks(text: string): string {
  return (text.match(/\{\d+\}/g) || []).sort().join()
}

/* Strings added by the command workbench are not in the shared catalog yet. */
const EXTRA_ZH: Readonly<Record<string, string>> = {
  'View Record': '查看记录', 'No data source loaded.': '未加载数据源。',
  'Insert Merge Field': '插入合并域', 'records': '条记录', 'Open Data Source…': '打开数据源…',
  'Merge to New Document': '合并到新文档', 'Change Case': '更改大小写', 'Case': '大小写',
  'Effects': '效果', 'Hidden': '隐藏', 'All caps': '全部大写', 'Small caps': '小型大写',
  'Double underline': '双下划线', 'Formula and Sorting': '公式与排序', 'Operation': '运算',
  'Sum': '求和', 'Average': '平均值', 'Sort ascending': '升序排列', 'Sort descending': '降序排列',
  'Borders and Shading': '边框和底纹', 'Colour': '颜色', 'Table Properties': '表格属性',
  'Field': '域', 'Field name': '域名', 'Bookmark name': '书签名', 'Address': '地址',
  'Insert Hyperlink': '插入超链接', 'Symbol': '符号', 'Number of cells': '单元格数',
  'Choose a CSV file with a header row.': '请选择带标题行的 CSV 文件。',
  '{0} records loaded.': '已加载 {0} 条记录。',
  'Formatting copied.': '已复制格式。', 'Formatting applied.': '已应用格式。',
  'Copy formatting first.': '请先复制格式。', 'Numbering restarted at 1.': '编号已从 1 重新开始。',
  'The previous paragraph is not numbered.': '上一段未编号。',
  'Place the cursor in a table first.': '请先将光标置于表格中。',
  'Place the cursor in a table cell first.': '请先将光标置于表格单元格中。',
  'Place the cursor in a paragraph first.': '请先将光标置于段落中。',
  'Place the cursor on the page first.': '请先将光标置于页面中。',
  'Place the cursor in a heading first.': '请先将光标置于标题中。',
  'Select a picture first.': '请先选择图片。', 'Enter a bookmark name.': '请输入书签名。',
  'Choose a bookmark to reference.': '请选择要引用的书签。', 'Enter a hyperlink address.': '请输入超链接地址。',
  'Load a CSV data source first.': '请先加载 CSV 数据源。', 'No headings in this document.': '本文档没有标题。',
  'No previous location.': '没有上一个位置。', 'No tracked changes.': '没有修订记录。',
  'Already at the start of the document.': '已在文档开头。', 'Already at the end of the document.': '已在文档结尾。',
  'The cursor must be above the last row.': '光标必须位于最后一行之前。',
  'Select one or more paragraphs first.': '请先选择一个或多个段落。',
  'Print Preview': '打印预览', 'Magnifier': '放大镜', 'One Page': '单页', 'Multiple Pages': '多页',
  'Page Width': '页宽', 'Text Width': '文字宽度', 'Whole Page': '整页', 'Two Pages': '双页',
  'Shrink to Fit': '缩小字体填充', 'Close Preview': '关闭预览',
  'This document is already one page long.': '本文档只有一页。',
  'Load a CSV data source before sending mail.': '发送邮件前请先加载 CSV 数据源。',
  'Document Map': '文档结构图', 'Thumbnails': '缩略图', 'Find': '查找', 'Research': '资料检索',
  'Close': '关闭', 'Increase Text Size': '增大文字', 'Decrease Text Size': '减小文字',
  'Actual Page': '实际大小的页', 'Allow Multiple Pages': '允许多页', '{0} x {1} Pages': '{0} × {1} 页',
  'Reading Layout': '阅读版式', 'Print': '打印',
  'Ruler': '标尺', 'Full Screen': '全屏显示', 'Help': '帮助',
}

/** Same contract as VO.t / L.t: a missing translation returns the English source. */
export function t(text: string, ...args: (string | number)[]): string {
  if (!text) return text
  const lang = initLanguage()
  const dict = dictionaries.get(lang === 'qps' ? 'en' : lang)
  const plain = withoutContext(text)
  let out = plain
  if (lang === 'qps') {
    const acc: Record<string, string> = {
      a: 'á', c: 'ç', e: 'é', i: 'í', n: 'ñ', o: 'ó', u: 'ú', y: 'ý',
      A: 'Á', C: 'Ç', E: 'É', I: 'Í', N: 'Ñ', O: 'Ó', U: 'Ú', Y: 'Ý',
    }
    const pseudo = plain.replace(/(\{\d+\}|&.)|[aceinouyACEINOUY]/g, (m, keep: string | undefined) => keep || acc[m])
    const letters = plain.replace(/\{\d+\}|[^A-Za-z]/g, '').length
    out = `[${pseudo}${letters > 2 ? ` ${'·'.repeat(Math.ceil(letters * 0.4))}` : ''}]`
  } else if (dict) {
    for (const key of [text, plain]) {
      const value = dict[key]
      if (typeof value === 'string' && value && placeholderMarks(value) === placeholderMarks(plain)) {
        out = value
        break
      }
    }
    if (out === plain && lang === 'zh-CN') {
      const extra = EXTRA_ZH[plain]
      if (extra) out = extra
    }
  }
  return args.length ? out.replace(/\{(\d+)\}/g, (m, i) => (args[Number(i)] != null ? String(args[Number(i)]) : m)) : out
}

/** Strip an access-key ampersand from translated text, as common/ui.js does for tooltips. */
export function plainAmp(text: string): string {
  return t(text).replace(/&(.)/, '$1')
}

/** A shortcut as the original shows it (common/ui.js → keyText): each key name in the
 *  UI's language, the modifiers in Apple's symbols and order on a Mac. Only the
 *  keys the original names are translated; the arrow keys keep their English. */
const MAC_MODS: [string, string][] = [['Alt', '⌥'], ['Shift', '⇧'], ['Ctrl', '⌘']]
const KEY_NAMES = ['Ctrl', 'Shift', 'Alt', 'Del', 'Ins', 'Esc', 'Home', 'End', 'Enter', 'Tab', 'Space', 'Backspace', 'Page Up', 'Page Down', 'PgUp', 'PgDn']

function keyName(part: string): string {
  return KEY_NAMES.includes(part) ? t(`${part}|key`) : part
}

export function keyText(key: string | undefined): string {
  if (!key) return ''
  const parts = String(key).split('+')
  const isMac = /Mac|iPhone|iPad/.test(navigator.platform || navigator.userAgent)
  if (!isMac) return parts.map(keyName).join('+')
  const mods = MAC_MODS.filter(([m]) => parts.includes(m)).map(([, s]) => s).join('')
  const rest = parts.filter((p) => !MAC_MODS.some(([m]) => m === p)).map(keyName).join('+')
  return mods + rest
}

/** A label as the original's menus show it: its access key underlined. */
export function ampHTML(text: string): { __html: string } {
  const escaped = t(text).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
  const marked = escaped.replace(/&amp;&amp;/g, '\u0000').replace(/&amp;(\S)/g, '<u>$1</u>').replace(/\u0000/g, '&amp;')
  return { __html: marked }
}
