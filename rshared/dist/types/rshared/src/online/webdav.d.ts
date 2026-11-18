import type { WebDavSource, OnlineFile, WebDavEntry } from './types';
export declare function getWebDav(s: WebDavSource): Promise<OnlineFile>;
export declare function putWebDav(s: WebDavSource, blob: Blob): Promise<void>;
export declare function listWebDav(s: WebDavSource): Promise<WebDavEntry[]>;
