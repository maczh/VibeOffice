import { Package } from '@voffice/shared';
import { type Doc } from './model';
export interface WriteOpts {
    pkg?: Package;
    part?: string;
    format?: string;
}
export declare function writeDocx(doc: Doc, opts?: WriteOpts): Promise<Blob>;
