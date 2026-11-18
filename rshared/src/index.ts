import { Component, createElement, type ReactNode, type ErrorInfo } from 'react'
export * from './office-core'

export interface WebDavSource { kind: 'webdav'; url: string; username?: string; password?: string; token?: string }
export interface S3Source { kind: 's3'; endpoint: string; bucket: string; key: string; region?: string; accessKeyId: string; secretAccessKey: string; sessionToken?: string; forcePathStyle?: boolean }
export interface PresignedSource { kind: 'presigned'; url: string; putUrl?: string }
export interface HttpSource { kind: 'http'; url: string; headers?: Record<string, string> }
export interface HttpPostTarget { kind: 'post'; url: string; field?: string; fields?: Record<string, string>; headers?: Record<string, string> }
export type OnlineSource = WebDavSource | S3Source | PresignedSource | HttpSource
export type SaveTarget = WebDavSource | S3Source | PresignedSource | HttpPostTarget
export type OpenInput = (File & { kind?: undefined }) | (Blob & { kind?: undefined }) | string | OnlineSource
export interface OnlineFile { blob: Blob; name: string }
export interface OfficeApi { open(input: OpenInput, name?: string): Promise<void>; newDocument(): void; getBlob(): Promise<Blob>; save(): Promise<void>; saveAs(name?: string): Promise<Blob>; saveTo(target: SaveTarget, name?: string): Promise<void>; readonly name: string; readonly dirty: boolean }
export interface OfficeProps { src?: OpenInput; fileName?: string; readonly?: boolean }
export function nameFromUrl(url: string): string { try { return decodeURIComponent(new URL(url).pathname.split('/').filter(Boolean).pop() || 'document') } catch { return 'document' } }
function headers(s: WebDavSource): Record<string,string> { if (s.token) return { Authorization: `Bearer ${s.token}` }; if (s.username != null) return { Authorization: `Basic ${btoa(unescape(encodeURIComponent(`${s.username}:${s.password ?? ''}`)))}` }; return {} }
export async function fetchOnline(source: OnlineSource): Promise<OnlineFile> {
  const h = source.kind === 'webdav' ? headers(source) : source.kind === 'http' ? source.headers : undefined
  const url = source.kind === 's3' ? `${source.endpoint.replace(/\/$/, '')}/${source.forcePathStyle ? `${source.bucket}/` : ''}${source.key}` : source.url
  const res = await fetch(url, { headers: h }); if (!res.ok) throw new Error(`Download failed: ${res.status} ${res.statusText}`)
  return { blob: await res.blob(), name: nameFromUrl(url) }
}
export async function saveOnline(target: SaveTarget, blob: Blob | Promise<Blob>, name: string): Promise<void> {
  blob = await blob
  if (target.kind === 'post') { const form = new FormData(); Object.entries(target.fields ?? {}).forEach(([k,v]) => form.append(k,v)); form.append(target.field ?? 'file', blob, name); const r = await fetch(target.url, { method:'POST', headers:target.headers, body:form }); if (!r.ok) throw new Error(`Upload failed: ${r.status} ${r.statusText}`); return }
  const url = target.kind === 's3' ? `${target.endpoint.replace(/\/$/, '')}/${target.forcePathStyle ? `${target.bucket}/` : ''}${target.key}` : target.kind === 'presigned' ? target.putUrl : target.url
  if (!url) throw new Error('The online source is read-only')
  const r = await fetch(url, { method: target.kind === 'webdav' ? 'PUT' : 'PUT', headers: target.kind === 'webdav' ? headers(target) : undefined, body: blob }); if (!r.ok) throw new Error(`Upload failed: ${r.status} ${r.statusText}`)
}
export function downloadBlob(blob: Blob, name: string): void { const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = name; a.click(); setTimeout(() => URL.revokeObjectURL(a.href), 0) }
export class ErrorBoundary extends Component<{children: ReactNode}, {error: Error | null}> { state = { error: null as Error | null }; static getDerivedStateFromError(error: Error) { return { error } }; componentDidCatch(error: Error, info: ErrorInfo) { console.error(error, info) }; render() { return this.state.error ? createElement('div', { role:'alert', className:'vo-error' }, this.state.error.message) : this.props.children } }
