import { Package } from '@voffice/react-shared';
import { type Doc } from './model';
export interface ReadResult {
    doc: Doc;
    pkg: Package;
    part: string;
}
export declare function readDocx(bytes: Uint8Array | ArrayBuffer): Promise<ReadResult>;
