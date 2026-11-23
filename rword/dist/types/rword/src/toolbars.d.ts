import { type QItem, type QTool } from './quireChrome';
/** The interface the toolbar builder needs from the editor. */
export interface ToolbarEnv {
    /** run a command by id */
    run: (id: string) => void;
    /** run a command with an argument */
    runArg: (id: string, arg: string) => void;
    /** a plain item that opens the command's dialog */
    item: (id: string) => QItem;
    /** the current zoom, as the zoom combo shows it */
    zoom: number;
    /** the current style, font and size */
    style: string;
    font: string;
    size: string;
    /** the current outline level and shown level */
    outlineLevel: number;
    outlineShow: number;
    /** the current mail-merge record */
    mergeRecord: number;
    mergeRecords: number;
    /** the picture colour mode currently applied */
    picMode: string;
    /** the wrap mode currently applied */
    wrapMode: string;
    /** the undo and redo history, newest last */
    undoLabels: string[];
    redoLabels: string[];
}
/** A row of the line-style list panel. */
export interface StyleRow {
    html: string;
    label?: string;
    run: () => void;
}
/** A toolbar button for a command id, carrying the original's icon and tip. */
export declare function toolButton(env: ToolbarEnv, id: string, run?: () => void): QTool;
export declare function buildToolbars(env: ToolbarEnv): {
    standard: QTool[];
    formatting: QTool[];
    tables: QTool[];
    reviewing: QTool[];
    outlining: QTool[];
    headerFooter: QTool[];
    picture: QTool[];
    mailMerge: QTool[];
    wordart: QTool[];
    drawing: QTool[];
};
