import { type ReactNode } from 'react';
/** Close the menu a panel sits in. Panels call it before they run their action. */
export declare function useMenuClose(): () => void;
export interface QItem {
    id?: string;
    label?: string;
    key?: string;
    icon?: string;
    sep?: boolean;
    checked?: boolean;
    radio?: boolean;
    disabled?: boolean;
    bold?: boolean;
    sub?: QItem[];
    /** a custom panel rendered in place of a menu row (the original's `custom` items) */
    panel?: ReactNode;
    /** a section heading row */
    heading?: string;
    /** the row's own tooltip (the original's `it.tip`) */
    tip?: string;
    /** the class the menu carrying this row gets (the original's `opts.cls`) */
    cls?: string;
    run?: () => void;
}
export interface QMenu {
    label: string;
    items: QItem[];
}
export interface QTool {
    id?: string;
    icon?: string;
    label?: string;
    tip?: string;
    sep?: boolean;
    disabled?: boolean;
    checked?: boolean;
    /** a dropdown arrow next to the button (the original's `tbDrop`) */
    drop?: QItem[];
    /** a split button: the main action plus a dropdown arrow (the original's `tbSplit`) */
    split?: {
        items: QItem[];
    };
    combo?: {
        id?: string;
        value: string;
        options: string[];
        width?: number;
        tip?: string;
        onChange?: (value: string) => void;
    };
    run?: () => void;
}
export declare const LABELS: Record<string, string>;
export declare function iconFor(id: string): string;
export declare function labelFor(id: string): string;
export declare function QIcon({ name, size }: {
    name: string;
    size?: number;
}): ReactNode;
export declare function QMenuBar({ menus, onSearch }: {
    menus: QMenu[];
    onSearch?: (query: string) => void;
}): ReactNode;
export declare function QToolButton({ item }: {
    item: QTool;
}): ReactNode;
export declare function QToolbar({ id, label, items, hidden }: {
    id: string;
    label: string;
    items: QTool[];
    hidden?: boolean;
}): ReactNode;
export declare function QViewBar({ view, onView, onBrowse, zoom }: {
    view: string;
    onView: (v: string) => void;
    onBrowse?: (dir: number) => void;
    zoom: number;
}): ReactNode;
export declare function QTaskPane({ title, onNav, children, hidden }: {
    title: string;
    onNav: (action: string) => void;
    children: ReactNode;
    hidden?: boolean;
}): ReactNode;
export declare function QDialog({ title, children, onClose, wide }: {
    title: string;
    children: ReactNode;
    onClose: () => void;
    wide?: boolean;
}): ReactNode;
