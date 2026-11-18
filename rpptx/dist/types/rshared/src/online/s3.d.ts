import type { S3Source, OnlineFile, WebDavEntry } from './types';
export declare function getS3(s: S3Source): Promise<OnlineFile>;
export declare function putS3(s: S3Source, blob: Blob): Promise<void>;
/** list objects under a prefix (ListObjectsV2), mapped to the shared entry shape */
export declare function listS3(s: S3Source, prefix?: string): Promise<WebDavEntry[]>;
