import { type Cell, type Doc, type Para, type Run, type Table } from '../engine/model';
export interface Location {
    blockIndex: number;
    para: Para | null;
    table: {
        blockIndex: number;
        row: number;
        col: number;
        cell: Cell;
        table: Table;
    } | null;
}
export interface CommandEnv {
    doc: Doc;
    loc: Location | null;
    selection: string;
    arg?: string;
}
export type CommandFn = (env: CommandEnv) => string | undefined;
export declare function locate(doc: Doc, id: string | null | undefined): Location | null;
export declare function fallbackLocation(doc: Doc): Location;
export declare const tableCommands: Record<string, CommandFn>;
export declare const formatCommands: Record<string, CommandFn>;
export declare const outlineCommands: Record<string, CommandFn>;
export declare const documentCommands: Record<string, CommandFn>;
export declare function commandText(env: CommandEnv): string;
export declare function allBlockText(doc: Doc): string;
export type { Run };
