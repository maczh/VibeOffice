import { Package } from '@voffice/shared';
import { type Workbook } from './model';
export interface XlsxWriteOpts {
    pkg?: Package;
    sheetParts?: string[];
    format?: string;
}
export declare function writeXlsx(wb: Workbook, opts?: XlsxWriteOpts): Promise<Blob>;
