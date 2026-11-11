export interface RemoteRevision { value: string; kind: 'etag' | 'version' | 'custom' }
export interface RemoteFile { id: string; name: string; mimeType?: string; isDirectory?: boolean; revision?: RemoteRevision | undefined; updatedAt?: number }
export interface RemoteSaveOptions { ifRevision?: RemoteRevision; signal?: AbortSignal }
export interface RemoteStorageProvider {
  list(parentId?: string, signal?: AbortSignal): Promise<RemoteFile[]>;
  read(id: string, signal?: AbortSignal): Promise<{ file: RemoteFile; blob: Blob }>;
  create(parentId: string | undefined, name: string, blob: Blob, signal?: AbortSignal): Promise<RemoteFile>;
  save(id: string, blob: Blob, options?: RemoteSaveOptions): Promise<RemoteFile>;
}
export type RemoteConflictError = Error & { code: 'revision-conflict'; current?: RemoteRevision | undefined };
export function remoteConflict(current?: RemoteRevision): RemoteConflictError {
  return Object.assign(new Error('The remote file changed since it was opened.'), { code: 'revision-conflict' as const, current });
}

export interface WebDavOptions {
  baseUrl: string | URL;
  credentials?: () => Promise<{ username: string; password: string } | undefined>;
  allowedOrigins: readonly string[];
  fetcher?: typeof fetch;
}

/** WebDAV uses only explicitly allowed origins; credentials are requested for each operation and never persisted. */
export function createWebDavProvider(options: WebDavOptions): RemoteStorageProvider {
  const base = new URL(options.baseUrl);
  const allowed = new Set(options.allowedOrigins.map((origin) => new URL(origin).origin));
  if (!['http:', 'https:'].includes(base.protocol) || !allowed.has(base.origin)) throw new TypeError('WebDAV origin is not allowed');
  const fetcher = options.fetcher ?? fetch;
  const urlFor = (id = '') => {
    const url = new URL(id, base);
    if (url.origin !== base.origin || !url.pathname.startsWith(base.pathname)) throw new TypeError('WebDAV path escaped its configured base URL');
    return url;
  };
  const request = async (url: URL, method: string, headers: Record<string, string> = {}, body?: BodyInit, signal?: AbortSignal) => {
    const auth = await options.credentials?.();
    const authorization = auth ? `Basic ${btoa(`${auth.username}:${auth.password}`)}` : undefined;
    const response = await fetcher(url, { method, headers: { ...(authorization ? { Authorization: authorization } : {}), ...headers }, ...(body ? { body } : {}), ...(signal ? { signal } : {}) });
    if (response.status === 412 || response.status === 409) throw remoteConflict();
    if (!response.ok && response.status !== 207) throw new Error(`WebDAV ${method} failed (${response.status})`);
    return response;
  };
  const revision = (response: Response): RemoteRevision | undefined => {
    const etag = response.headers.get('ETag');
    return etag ? { value: etag, kind: 'etag' } : undefined;
  };
  return {
    async list(parentId, signal) {
      const response = await request(urlFor(parentId), 'PROPFIND', { Depth: '1', 'Content-Type': 'application/xml; charset=utf-8' }, '<?xml version="1.0"?><d:propfind xmlns:d="DAV:"><d:prop><d:displayname/><d:getetag/><d:getcontenttype/><d:getlastmodified/><d:resourcetype/></d:prop></d:propfind>', signal);
      const xml = new DOMParser().parseFromString(await response.text(), 'application/xml');
      if (xml.querySelector('parsererror')) throw new Error('WebDAV returned invalid XML');
      return [...xml.getElementsByTagNameNS('DAV:', 'response')].map((node) => {
        const href = node.getElementsByTagNameNS('DAV:', 'href')[0]?.textContent || '';
        const name = node.getElementsByTagNameNS('DAV:', 'displayname')[0]?.textContent || decodeURIComponent(href.split('/').filter(Boolean).pop() || '');
        const etag = node.getElementsByTagNameNS('DAV:', 'getetag')[0]?.textContent || undefined;
        const mimeType = node.getElementsByTagNameNS('DAV:', 'getcontenttype')[0]?.textContent || undefined;
        const isDirectory = node.getElementsByTagNameNS('DAV:', 'collection').length > 0;
        return { id: href, name, isDirectory, ...(mimeType ? { mimeType } : {}), ...(etag ? { revision: { value: etag, kind: 'etag' as const } } : {}) };
      }).filter((file) => file.id !== urlFor(parentId).pathname);
    },
    async read(id, signal) {
      const url = urlFor(id);
      const response = await request(url, 'GET', {}, undefined, signal);
      return { file: { id: url.pathname, name: decodeURIComponent(url.pathname.split('/').filter(Boolean).pop() || 'Document'), ...(response.headers.get('Content-Type') ? { mimeType: response.headers.get('Content-Type')! } : {}), ...(revision(response) ? { revision: revision(response) } : {}) }, blob: await response.blob() };
    },
    async create(parentId, name, blob, signal) {
      const url = urlFor((parentId ? parentId.replace(/\/$/, '') + '/' : '') + encodeURIComponent(name));
      const response = await request(url, 'PUT', { 'If-None-Match': '*' }, blob, signal);
      return { id: url.pathname, name, ...(revision(response) ? { revision: revision(response) } : {}) };
    },
    async save(id, blob, options) {
      const url = urlFor(id);
      const headers: Record<string, string> = {};
      if (options?.ifRevision?.kind === 'etag') headers['If-Match'] = options.ifRevision.value;
      const response = await request(url, 'PUT', headers, blob, options?.signal);
      return { id: url.pathname, name: decodeURIComponent(url.pathname.split('/').filter(Boolean).pop() || 'Document'), ...(revision(response) ? { revision: revision(response) } : {}) };
    },
  };
}

export function createUrlReadProvider(allowedOrigins: readonly string[], fetcher: typeof fetch = fetch): Pick<RemoteStorageProvider, 'read'> {
  const allowed = new Set(allowedOrigins.map((origin) => new URL(origin).origin));
  return { async read(id, signal) {
    const url = new URL(id);
    if (!['http:', 'https:'].includes(url.protocol) || !allowed.has(url.origin)) throw new TypeError('URL origin is not allowed');
    const response = await fetcher(url, { ...(signal ? { signal } : {}) });
    if (!response.ok) throw new Error(`URL open failed (${response.status})`);
    return { file: { id: url.href, name: decodeURIComponent(url.pathname.split('/').pop() || 'Document'), ...(response.headers.get('ETag') ? { revision: { value: response.headers.get('ETag')!, kind: 'etag' as const } } : {}) }, blob: await response.blob() };
  } };
}

export function createRemoteFileCallbacks(
  provider: Pick<RemoteStorageProvider, 'read'> & Partial<Pick<RemoteStorageProvider, 'save' | 'create'>>,
  chooseOpen: () => Promise<RemoteFile | null>,
  chooseSave?: (suggestedName: string) => Promise<{ id?: string; parentId?: string; name: string } | null>,
) {
  return {
    async onOpenFile() {
      const selected = await chooseOpen();
      if (!selected) return null;
      const { blob, file } = await provider.read(selected.id);
      const name = file.name || selected.name;
      return { file: new File([blob], name, { type: file.mimeType || blob.type }), remote: { provider, file } };
    },
    async onSaveFile(name: string, blob: Blob, current?: import('./local-files.js').LocalDocument) {
      if (current?.remote) {
        const { provider: source, file } = current.remote;
        if (!source.save) throw new Error('This remote source is read-only. Choose a writable destination.');
        const saved = await source.save(file.id, blob, file.revision ? { ifRevision: file.revision } : {});
        return { file: new File([blob], saved.name || name, { type: blob.type }), remote: { provider: source, file: saved } };
      }
      if (!chooseSave) throw new Error('No remote destination chooser is configured.');
      const destination = await chooseSave(name);
      if (!destination) return null;
      let file: RemoteFile;
      if (destination.id && provider.save) file = await provider.save(destination.id, blob);
      else if (provider.create) file = await provider.create(destination.parentId, destination.name, blob);
      else throw new Error('This provider cannot create files.');
      return { file: new File([blob], file.name || destination.name, { type: blob.type }), remote: { provider, file } };
    },
  };
}
