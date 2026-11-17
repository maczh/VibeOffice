/* S3-compatible storage access with AWS Signature Version 4 (Web Crypto), plus list.
 * Works with AWS S3, MinIO and other S3-compatible endpoints; the bucket must allow CORS. */
import type { S3Source, OnlineFile, WebDavEntry } from './types'
import { OnlineError, nameFromUrl } from './types'
import { parseXml } from '../core/xml'

const enc = new TextEncoder()

const hex = (b: ArrayBuffer): string => Array.from(new Uint8Array(b), (x) => x.toString(16).padStart(2, '0')).join('')
const sha256 = async (data: Uint8Array | string): Promise<string> =>
  hex(await crypto.subtle.digest('SHA-256', (typeof data === 'string' ? enc.encode(data) : data) as BufferSource))
const hmac = async (key: Uint8Array | ArrayBuffer, data: string): Promise<ArrayBuffer> => {
  const k = await crypto.subtle.importKey('raw', key as BufferSource, { name: 'HMAC', hash: 'SHA-256' }, false, ['sign'])
  return crypto.subtle.sign('HMAC', k, enc.encode(data))
}
/** RFC 3986 encoding as SigV4 requires (unreserved characters stay, '/' kept in paths) */
const uriEncode = (s: string, keepSlash = false): string =>
  encodeURIComponent(s).replace(/[!'()*]/g, (c) => '%' + c.charCodeAt(0).toString(16).toUpperCase()).replace(/%2F/g, keepSlash ? '/' : '%2F')

function objectUrl(s: S3Source): URL {
  const ep = new URL(s.endpoint)
  const key = s.key.replace(/^\/+/, '')
  if (s.forcePathStyle) {
    ep.pathname = '/' + s.bucket + '/' + key
  } else {
    ep.hostname = s.bucket + '.' + ep.hostname
    ep.pathname = '/' + key
  }
  return ep
}

function bucketUrl(s: S3Source): URL {
  const ep = new URL(s.endpoint)
  if (s.forcePathStyle) ep.pathname = '/' + s.bucket
  else ep.hostname = s.bucket + '.' + ep.hostname
  return ep
}

async function signedFetch(s: S3Source, url: URL, method: string, body: Uint8Array | null, extraQuery: Record<string, string> = {}): Promise<Response> {
  const region = s.region || 'us-east-1'
  const now = new Date()
  const amzDate = now.toISOString().replace(/[:-]|\.\d{3}/g, '')
  const dateStamp = amzDate.slice(0, 8)
  for (const [k, v] of Object.entries(extraQuery)) url.searchParams.set(k, v)
  const payloadHash = await sha256(body ?? '')
  const headers: Record<string, string> = {
    host: url.host,
    'x-amz-content-sha256': payloadHash,
    'x-amz-date': amzDate,
  }
  if (s.sessionToken) headers['x-amz-security-token'] = s.sessionToken
  const signedHeaders = Object.keys(headers).sort().join(';')
  const canonicalHeaders = Object.keys(headers).sort().map((k) => `${k}:${headers[k].trim()}\n`).join('')
  const canonicalQuery = [...url.searchParams.entries()].sort(([a], [b]) => a < b ? -1 : a > b ? 1 : 0)
    .map(([k, v]) => `${uriEncode(k)}=${uriEncode(v)}`).join('&')
  const canonicalRequest = [method, uriEncode(url.pathname, true), canonicalQuery, canonicalHeaders, signedHeaders, payloadHash].join('\n')
  const scope = `${dateStamp}/${region}/s3/aws4_request`
  const stringToSign = ['AWS4-HMAC-SHA256', amzDate, scope, await sha256(canonicalRequest)].join('\n')
  const kDate = await hmac(enc.encode('AWS4' + s.secretAccessKey), dateStamp)
  const kRegion = await hmac(kDate, region)
  const kService = await hmac(kRegion, 's3')
  const kSigning = await hmac(kService, 'aws4_request')
  const signature = hex(await hmac(kSigning, stringToSign))
  const authorization = `AWS4-HMAC-SHA256 Credential=${s.accessKeyId}/${scope}, SignedHeaders=${signedHeaders}, Signature=${signature}`
  const reqHeaders: Record<string, string> = { ...headers, Authorization: authorization }
  delete reqHeaders.host // the browser sets Host itself
  return fetch(url.toString(), { method, headers: reqHeaders, body: body as BodyInit | null })
}

export async function getS3(s: S3Source): Promise<OnlineFile> {
  const res = await signedFetch(s, objectUrl(s), 'GET', null)
  if (!res.ok) throw new OnlineError(`S3 GET failed: ${res.status} ${res.statusText}`, res.status)
  return { blob: await res.blob(), name: nameFromUrl(objectUrl(s).toString()) }
}

export async function putS3(s: S3Source, blob: Blob): Promise<void> {
  const body = new Uint8Array(await blob.arrayBuffer())
  const res = await signedFetch(s, objectUrl(s), 'PUT', body)
  if (!res.ok) throw new OnlineError(`S3 PUT failed: ${res.status} ${res.statusText}`, res.status)
}

/** list objects under a prefix (ListObjectsV2), mapped to the shared entry shape */
export async function listS3(s: S3Source, prefix = ''): Promise<WebDavEntry[]> {
  const url = bucketUrl(s)
  const res = await signedFetch(s, url, 'GET', null, { 'list-type': '2', prefix, 'max-keys': '1000' })
  if (!res.ok) throw new OnlineError(`S3 list failed: ${res.status} ${res.statusText}`, res.status)
  const root = parseXml(await res.text())
  const out: WebDavEntry[] = []
  for (const c of root.getElementsByTagName('Contents')) {
    const key = c.getElementsByTagName('Key')[0]?.textContent || ''
    out.push({
      name: key.split('/').pop() || key,
      href: key,
      isDir: false,
      size: +(c.getElementsByTagName('Size')[0]?.textContent || 0),
      modified: c.getElementsByTagName('LastModified')[0]?.textContent,
    })
  }
  for (const c of root.getElementsByTagName('CommonPrefixes')) {
    const p = c.getElementsByTagName('Prefix')[0]?.textContent || ''
    out.push({ name: p.replace(/\/$/, '').split('/').pop() || p, href: p, isDir: true })
  }
  return out
}
