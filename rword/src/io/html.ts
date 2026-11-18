export function htmlToText(source: string): string {
  if (typeof DOMParser === 'undefined') return source.replace(/<[^>]+>/g, '')
  const doc = new DOMParser().parseFromString(source, 'text/html')
  return doc.body?.innerText ?? doc.body?.textContent ?? ''
}
export function textToHtml(text: string, title = 'Document'): string {
  const esc = (value: string) => value.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
  const body = text.split(/\r?\n/).map((line) => `<p>${esc(line) || '&nbsp;'}</p>`).join('\n')
  return `<!doctype html><html lang="en"><head><meta charset="utf-8"><title>${esc(title)}</title></head><body>${body}</body></html>`
}
