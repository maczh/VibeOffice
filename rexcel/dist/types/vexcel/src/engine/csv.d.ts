import { type Workbook, type Sheet } from './model';
export declare function parseCsv(text: string, delim?: string): Workbook;
export declare function writeCsv(wb: Workbook, sheet: Sheet, delim?: string): string;
