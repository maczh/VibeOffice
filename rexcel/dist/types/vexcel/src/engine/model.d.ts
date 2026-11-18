export interface CellStyle {
    font?: {
        name?: string;
        size?: number;
        bold?: boolean;
        italic?: boolean;
        underline?: boolean;
        color?: string;
    };
    fill?: string;
    border?: {
        top?: string;
        right?: string;
        bottom?: string;
        left?: string;
    };
    align?: {
        h?: 'left' | 'center' | 'right';
        v?: 'top' | 'middle' | 'bottom';
        wrap?: boolean;
    };
    numFmt?: string;
}
export interface Cell {
    /** computed value */
    v?: string | number | boolean | null;
    /** formula text without '=' */
    f?: string;
    /** cached type: n number, s string, b bool, e error */
    t?: 'n' | 's' | 'b' | 'e';
    s?: number;
}
export interface Merge {
    r1: number;
    c1: number;
    r2: number;
    c2: number;
}
export interface Sheet {
    name: string;
    cells: Map<string, Cell>;
    merges: Merge[];
    colWidths: Map<number, number>;
    rowHeights: Map<number, number>;
    defaultColWidth: number;
    defaultRowHeight: number;
    frozen?: {
        rows: number;
        cols: number;
    };
}
export interface Workbook {
    sheets: Sheet[];
    styles: CellStyle[];
    active: number;
}
export declare const colName: (c: number) => string;
export declare const colIndex: (s: string) => number;
export declare const addr: (r: number, c: number) => string;
export declare function parseAddr(a: string): {
    r: number;
    c: number;
} | null;
export declare const newSheet: (name: string) => Sheet;
export declare const newWorkbook: () => Workbook;
export declare function styleIndex(wb: Workbook, s: CellStyle): number;
export declare function getCell(sheet: Sheet, r: number, c: number): Cell | undefined;
export declare function setCell(sheet: Sheet, r: number, c: number, cell: Cell | undefined): void;
/** display text of a cell (formula results already computed) */
export declare function cellText(cell: Cell | undefined, wb: Workbook, fmt: (v: Cell['v'], f?: string) => string): string;
export declare const cloneWorkbook: (wb: Workbook) => Workbook;
