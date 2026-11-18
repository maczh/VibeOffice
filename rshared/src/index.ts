import { Component, createElement, type ReactNode, type ErrorInfo } from 'react'
import { getS3, putS3 } from './online/s3'
import { getWebDav, putWebDav } from './online/webdav'
import { getHttp, postFile } from './online/http'
import type { S3Source as CoreS3Source, WebDavSource as CoreWebDavSource, HttpSource as CoreHttpSource, HttpPostTarget as CorePostTarget } from './online/types'
export * from './office-core'
export * from './chrome'
export * from './common'

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
  if (source.kind === 's3') return getS3(source as CoreS3Source)
  if (source.kind === 'webdav') return getWebDav(source as CoreWebDavSource)
  if (source.kind === 'http') return getHttp(source as CoreHttpSource)
  const res = await fetch(source.url); if (!res.ok) throw new Error(`Download failed: ${res.status} ${res.statusText}`)
  return { blob: await res.blob(), name: nameFromUrl(source.url) }
}
export async function saveOnline(target: SaveTarget, blob: Blob | Promise<Blob>, name: string): Promise<void> {
  blob = await blob
  if (target.kind === 'post') return postFile(target as CorePostTarget, blob, name)
  if (target.kind === 's3') return putS3(target as CoreS3Source, blob)
  if (target.kind === 'webdav') return putWebDav(target as CoreWebDavSource, blob)
  if (!target.putUrl) throw new Error('The online source is read-only')
  const r = await fetch(target.putUrl, { method: 'PUT', headers: { 'Content-Type': blob.type || 'application/octet-stream' }, body: blob }); if (!r.ok) throw new Error(`Upload failed: ${r.status} ${r.statusText}`)
}
export function downloadBlob(blob: Blob, name: string): void { const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = name; a.click(); setTimeout(() => URL.revokeObjectURL(a.href), 0) }
export class ErrorBoundary extends Component<{children: ReactNode}, {error: Error | null}> { state = { error: null as Error | null }; static getDerivedStateFromError(error: Error) { return { error } }; componentDidCatch(error: Error, info: ErrorInfo) { console.error(error, info) }; render() { return this.state.error ? createElement('div', { role:'alert', className:'vo-error' }, this.state.error.message) : this.props.children } }
