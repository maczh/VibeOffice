import { type MenuDef, type ToolItem } from '@voffice/react-shared';
export interface CommandSurfaceProps {
    menus: MenuDef[];
    tools: ToolItem[];
}
export declare function CommandSurface({ menus, tools }: CommandSurfaceProps): import("react").JSX.Element;
