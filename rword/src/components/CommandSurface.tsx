import { MenuBar, Toolbar, type MenuDef, type ToolItem } from '@voffice/react-shared'
export interface CommandSurfaceProps { menus: MenuDef[]; tools: ToolItem[] }
export function CommandSurface({ menus, tools }: CommandSurfaceProps) { return <><MenuBar menus={menus} /><Toolbar items={tools} /></> }
