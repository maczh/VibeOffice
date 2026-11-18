export declare class Err {
    code: string;
    constructor(code: string);
    toString(): string;
}
export declare const DIV0: () => Err;
export declare const VALUE: () => Err;
export declare const REF: () => Err;
export declare const NAME: () => Err;
export declare const NA: () => Err;
export declare const NUM: () => Err;
export type Val = number | string | boolean | Err | Val[];
export interface RefNode {
    r: number;
    c: number;
    sheet?: string;
}
export interface RangeNode {
    r1: number;
    c1: number;
    r2: number;
    c2: number;
    sheet?: string;
}
export type Node = {
    k: 'num';
    v: number;
} | {
    k: 'str';
    v: string;
} | {
    k: 'bool';
    v: boolean;
} | {
    k: 'err';
    v: string;
} | {
    k: 'ref';
    ref: RefNode;
} | {
    k: 'range';
    range: RangeNode;
} | {
    k: 'un';
    op: string;
    a: Node;
} | {
    k: 'bin';
    op: string;
    a: Node;
    b: Node;
} | {
    k: 'pct';
    a: Node;
} | {
    k: 'call';
    name: string;
    args: Node[];
};
export declare function parseFormula(src: string): Node;
export declare function toNum(v: Val): number | Err;
export declare function toStr(v: Val): string | Err;
export declare function toBool(v: Val): boolean | Err;
export interface EvalCtx {
    get(ref: RefNode): Val;
    range(range: RangeNode): Val[];
    current?: {
        r: number;
        c: number;
    };
}
export type RangeVal = Val[] & {
    __dims?: [number, number];
};
export declare function evalNode(n: Node, ctx: EvalCtx): Val;
export declare function excelDate(d: Date): number;
export declare function fromExcel(n: number): Date;
export declare function callFn(name: string, args: Val[], ctx: EvalCtx): Val;
export declare function collectRefs(n: Node, out: {
    refs: RefNode[];
    ranges: RangeNode[];
}): void;
