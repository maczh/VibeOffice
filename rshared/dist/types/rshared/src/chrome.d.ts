import { type ReactNode } from 'react';
export interface MenuItem {
    label: string;
    key?: string;
    sep?: boolean;
    disabled?: boolean;
    checked?: boolean;
    sub?: MenuItem[];
    run?: () => void | Promise<void>;
}
export interface MenuDef {
    label: string;
    items: MenuItem[];
}
export interface ToolItem {
    id: string;
    label?: string;
    title?: string;
    icon?: ReactNode;
    kind?: 'button' | 'sep';
    run?: () => void | Promise<void>;
}
export declare function MenuBar({ menus }: {
    menus: MenuDef[];
}): import("react").JSX.Element;
export declare function Toolbar({ items }: {
    items: ToolItem[];
}): import("react").JSX.Element;
export declare function Ribbon({ tabs }: {
    tabs: {
        id: string;
        label: string;
        items: ToolItem[];
    }[];
}): import("react").JSX.Element;
export declare function StatusBar({ children }: {
    children: ReactNode;
}): import("react").JSX.Element;
export declare function Dialog({ title, children, onClose }: {
    title: string;
    children: ReactNode;
    onClose: () => void;
}): import("react").JSX.Element;
export declare function OnlineDialog({ mode, fileName, onSubmit, onClose }: {
    mode: 'open' | 'save';
    fileName?: string;
    onSubmit: (url: string) => void;
    onClose: () => void;
}): import("react").JSX.Element;
export declare function ContextMenu({ x, y, items, onClose }: {
    x: number;
    y: number;
    items: MenuItem[];
    onClose: () => void;
}): import("react").JSX.Element;
