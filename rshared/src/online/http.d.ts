import type { HttpSource, HttpPostTarget, OnlineFile } from './types';
export declare function getHttp(s: HttpSource): Promise<OnlineFile>;
export declare function postFile(t: HttpPostTarget, blob: Blob, name: string): Promise<void>;
