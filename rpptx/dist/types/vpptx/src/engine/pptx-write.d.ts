import { Package } from '@voffice/shared';
import { type Presentation } from './model';
export interface PptxWriteOpts {
    pkg?: Package;
    slideParts?: string[];
    format?: string;
    newImages?: {
        shapeId: string;
        part: string;
        rId: string;
        bytes: Uint8Array;
    }[];
}
export declare function writePptx(pres: Presentation, opts?: PptxWriteOpts): Promise<Blob>;
