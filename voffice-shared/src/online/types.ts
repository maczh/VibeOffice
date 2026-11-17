/* Online file sources and save targets. All requests run in the browser; the server
 * (WebDAV share, S3-compatible store, upload endpoint) must allow this origin via CORS. */

export interface WebDavSource {
  kind: 'webdav'
  /** full URL of the file, or of a collection when listing */
  url: string
  username?: string
  password?: string
  /** a Bearer token instead of Basic credentials */
  token?: string
}

export interface S3Source {
  kind: 's3'
  /** e.g. https://s3.amazonaws.com or https://minio.example.com:9000 */
  endpoint: string
  bucket: string
  key: string
  region?: string
  accessKeyId: string
  secretAccessKey: string
  sessionToken?: string
  /** path-style (bucket in the path) instead of virtual-host style; MinIO usually needs this */
  forcePathStyle?: boolean
}

export interface PresignedSource {
  kind: 'presigned'
  /** presigned GET url */
  url: string
  /** presigned PUT url for saving; without it the source is read-only */
  putUrl?: string
}

export interface HttpSource {
  kind: 'http'
  url: string
  headers?: Record<string, string>
}

export interface HttpPostTarget {
  kind: 'post'
  url: string
  /** multipart form field carrying the file; default 'file' */
  field?: string
  /** extra form fields sent alongside */
  fields?: Record<string, string>
  headers?: Record<string, string>
}

export type OnlineSource = WebDavSource | S3Source | PresignedSource | HttpSource
export type SaveTarget = WebDavSource | S3Source | PresignedSource | HttpPostTarget

export interface OnlineFile { blob: Blob; name: string }

export interface WebDavEntry {
  name: string
  href: string
  isDir: boolean
  size?: number
  modified?: string
}

export class OnlineError extends Error {
  status?: number
  constructor(message: string, status?: number) { super(message); this.status = status }
}

export function nameFromUrl(url: string): string {
  try {
    const p = new URL(url).pathname
    const last = p.split('/').filter(Boolean).pop() || 'document'
    return decodeURIComponent(last)
  } catch { return 'document' }
}

export function authHeaders(s: WebDavSource): Record<string, string> {
  const h: Record<string, string> = {}
  if (s.token) h.Authorization = 'Bearer ' + s.token
  else if (s.username != null) h.Authorization = 'Basic ' + btoa(unescape(encodeURIComponent(`${s.username}:${s.password ?? ''}`)))
  return h
}
