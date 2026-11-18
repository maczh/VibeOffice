import { type Sheet, type Workbook } from './model';
export declare const HDR_W = 46;
export declare const HDR_H = 22;
export interface Layout {
    xs: number[];
    ys: number[];
    nw: number;
    nh: number;
}
export interface Sel {
    r: number;
    c: number;
    r2: number;
    c2: number;
}
export declare function computeLayout(sheet: Sheet, maxC?: number, maxR?: number): Layout;
export declare function cellRect(layout: Layout, sheet: Sheet, r: number, c: number): {
    x: number;
    y: number;
    w: number;
    h: number;
};
export declare function pointToCell(layout: Layout, x: number, y: number): {
    r: number;
    c: number;
} | null;
export interface DrawState {
    sheet: Sheet;
    wb: Workbook;
    layout: Layout;
    sel: Sel;
    scrollX: number;
    scrollY: number;
    zoom: number;
    showGrid: boolean;
    copyRange?: Sel | null;
}
export declare function drawGrid(canvas: HTMLCanvasElement, st: DrawState): void;
export declare function normSel(s: Sel): {
    r1: number;
    c1: number;
    r2: number;
    c2: number;
};
