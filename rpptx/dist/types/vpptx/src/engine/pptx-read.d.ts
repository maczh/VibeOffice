import { Package } from '@voffice/shared';
import { type Presentation } from './model';
export interface PptxReadResult {
    pres: Presentation;
    pkg: Package;
    slideParts: string[];
}
export declare function readPptx(bytes: Uint8Array | ArrayBuffer): Promise<PptxReadResult>;
