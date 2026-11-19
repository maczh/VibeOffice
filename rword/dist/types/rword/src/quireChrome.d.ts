import { type ReactNode } from 'react';
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
    drop?: QItem[];
    combo?: {
        value: string;
        options: string[];
        width?: number;
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
