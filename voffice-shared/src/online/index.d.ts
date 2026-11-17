import type { OnlineSource, SaveTarget, OnlineFile, WebDavEntry, WebDavSource, S3Source } from './types';
export * from './types';
export { listWebDav } from './webdav';
export { listS3 } from './s3';
/** open a file from any online source */
export declare function fetchOnline(source: OnlineSource): Promise<OnlineFile>;
/** save a file to any online target */
export declare function saveOnline(target: SaveTarget, blob: Blob, name: string): Promise<void>;
/** browse a WebDAV collection or an S3 prefix with the same entry shape */
export declare function listOnline(source: WebDavSource | S3Source, pathOrPrefix?: string): Promise<WebDavEntry[]>;
