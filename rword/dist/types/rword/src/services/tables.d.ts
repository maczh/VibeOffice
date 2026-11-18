import { type Doc, type Table } from '../engine/model';
export declare function createTable(rows: number, columns: number): Table;
export declare function tableCell(table: Table, row: number, column: number): import("../engine").Cell;
export declare function addTableRow(table: Table): Table;
export declare function addTableColumn(table: Table): Table;
export declare function removeTableRow(table: Table, row: number): Table;
export declare function removeTableColumn(table: Table, column: number): Table;
export declare function insertTable(doc: Doc, rows: number, columns: number): Table;
