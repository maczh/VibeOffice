import type { MenuDef, ToolItem } from '@voffice/react-shared';
export interface QuireCommandActions {
    onNew: () => void;
    onOpen: () => void;
    onSave: () => void;
    onUndo: () => void;
    onRedo: () => void;
    onFind: () => void;
    onInsertTable: () => void;
}
export declare function quireMenus(a: QuireCommandActions): MenuDef[];
export declare function quireTools(a: QuireCommandActions): ToolItem[];
