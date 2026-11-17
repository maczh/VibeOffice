import type { OnlineSource, SaveTarget, OnlineFile, WebDavEntry, WebDavSource, S3Source } from './types'
import { OnlineError } from './types'
import { getWebDav, putWebDav, listWebDav } from './webdav'
import { getS3, putS3, listS3 } from './s3'
import { getHttp, postFile } from './http'

export * from './types'
export { listWebDav } from './webdav'
export { listS3 } from './s3'

/** open a file from any online source */
export async function fetchOnline(source: OnlineSource): Promise<OnlineFile> {
  switch (source.kind) {
    case 'webdav': return getWebDav(source)
    case 's3': return getS3(source)
    case 'presigned': {
      const res = await fetch(source.url)
      if (!res.ok) throw new OnlineError(`Download failed: ${res.status} ${res.statusText}`, res.status)
      return { blob: await res.blob(), name: new URL(source.url).pathname.split('/').filter(Boolean).pop() || 'document' }
    }
    case 'http': return getHttp(source)
  }
}

/** save a file to any online target */
export async function saveOnline(target: SaveTarget, blob: Blob, name: string): Promise<void> {
  switch (target.kind) {
    case 'webdav': return putWebDav(target, blob)
    case 's3': return putS3(target, blob)
    case 'presigned': {
      if (!target.putUrl) throw new OnlineError('This presigned source has no PUT url; it is read-only.')
      const res = await fetch(target.putUrl, { method: 'PUT', headers: { 'Content-Type': blob.type || 'application/octet-stream' }, body: blob })
      if (!res.ok) throw new OnlineError(`Upload failed: ${res.status} ${res.statusText}`, res.status)
      return
    }
    case 'post': return postFile(target, blob, name)
  }
}

/** browse a WebDAV collection or an S3 prefix with the same entry shape */
export async function listOnline(source: WebDavSource | S3Source, pathOrPrefix = ''): Promise<WebDavEntry[]> {
  if (source.kind === 'webdav') {
    const url = pathOrPrefix ? new URL(pathOrPrefix, source.url).toString() : source.url
    return listWebDav({ ...source, url })
  }
  return listS3(source, pathOrPrefix)
}
