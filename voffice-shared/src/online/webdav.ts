/* WebDAV access: GET / PUT a file, PROPFIND a collection. Basic or Bearer auth; CORS required. */
import type { WebDavSource, OnlineFile, WebDavEntry } from './types'
import { OnlineError, authHeaders, nameFromUrl } from './types'
import { parseXml } from '../core/xml'

export async function getWebDav(s: WebDavSource): Promise<OnlineFile> {
  const res = await fetch(s.url, { headers: authHeaders(s) })
  if (!res.ok) throw new OnlineError(`WebDAV GET failed: ${res.status} ${res.statusText}`, res.status)
  return { blob: await res.blob(), name: nameFromUrl(s.url) }
}

export async function putWebDav(s: WebDavSource, blob: Blob): Promise<void> {
  const res = await fetch(s.url, {
    method: 'PUT',
    headers: { ...authHeaders(s), 'Content-Type': blob.type || 'application/octet-stream' },
    body: blob,
  })
  if (!res.ok) throw new OnlineError(`WebDAV PUT failed: ${res.status} ${res.statusText}`, res.status)
}

const PROPFIND_BODY = '<?xml version="1.0" encoding="utf-8"?><propfind xmlns="DAV:"><prop><displayname/><getcontentlength/><getlastmodified/><resourcetype/></prop></propfind>'

export async function listWebDav(s: WebDavSource): Promise<WebDavEntry[]> {
  const res = await fetch(s.url, {
    method: 'PROPFIND',
    headers: { ...authHeaders(s), Depth: '1', 'Content-Type': 'application/xml' },
    body: PROPFIND_BODY,
  })
  if (!res.ok && res.status !== 207) throw new OnlineError(`WebDAV PROPFIND failed: ${res.status} ${res.statusText}`, res.status)
  const root = parseXml(await res.text())
  const base = new URL(s.url)
  const basePath = base.pathname.replace(/\/$/, '')
  const out: WebDavEntry[] = []
  for (const resp of root.getElementsByTagName('response')) {
    const href = resp.getElementsByTagName('href')[0]?.textContent || ''
    let path: string
    try { path = new URL(href, base).pathname } catch { path = href }
    const clean = path.replace(/\/$/, '')
    if (clean === basePath || clean === base.pathname) continue // the collection itself
    const prop = resp.getElementsByTagName('propstat')[0]?.getElementsByTagName('prop')[0]
    const isDir = !!prop?.getElementsByTagName('resourcetype')[0]?.getElementsByTagName('collection').length
    const name = prop?.getElementsByTagName('displayname')[0]?.textContent || decodeURIComponent(clean.split('/').pop() || clean)
    out.push({
      name,
      href: path,
      isDir,
      size: +(prop?.getElementsByTagName('getcontentlength')[0]?.textContent || 0) || undefined,
      modified: prop?.getElementsByTagName('getlastmodified')[0]?.textContent,
    })
  }
  return out
}
