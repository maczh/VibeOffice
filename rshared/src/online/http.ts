/* Plain HTTP GET and multipart HTTP POST upload. */
import type { HttpSource, HttpPostTarget, OnlineFile } from './types'
import { OnlineError, nameFromUrl } from './types'

export async function getHttp(s: HttpSource): Promise<OnlineFile> {
  const res = await fetch(s.url, { headers: s.headers })
  if (!res.ok) throw new OnlineError(`Download failed: ${res.status} ${res.statusText}`, res.status)
  const cd = res.headers.get('Content-Disposition') || ''
  const m = /filename\*?=(?:UTF-8''|")?([^";]+)/i.exec(cd)
  return { blob: await res.blob(), name: m ? decodeURIComponent(m[1].replace(/"$/, '')) : nameFromUrl(s.url) }
}

export async function postFile(t: HttpPostTarget, blob: Blob, name: string): Promise<void> {
  const form = new FormData()
  for (const [k, v] of Object.entries(t.fields || {})) form.append(k, v)
  form.append(t.field || 'file', blob, name)
  const res = await fetch(t.url, { method: 'POST', headers: t.headers, body: form })
  if (!res.ok) throw new OnlineError(`Upload failed: ${res.status} ${res.statusText}`, res.status)
}
