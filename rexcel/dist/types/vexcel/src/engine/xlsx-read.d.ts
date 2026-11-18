import { Package } from '@voffice/shared';
import { type Workbook } from './model';
export interface XlsxReadResult {
    wb: Workbook;
    pkg: Package;
    sheetParts: string[];
}
export declare function readXlsx(bytes: Uint8Array | ArrayBuffer): Promise<XlsxReadResult>;
