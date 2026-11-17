/* Command-bar vocabulary shared by the three apps: menus, toolbars, ribbon, status bar.
 * Labels arrive already translated (t()); icons are names from ui/icons.ts. */

export interface MenuItem {
  id?: string
  label?: string
  icon?: string
  /** shortcut text shown at the right */
  key?: string
  checked?: boolean
  radio?: boolean
  disabled?: boolean
  bold?: boolean
  sep?: boolean
  head?: string
  sub?: MenuItem[]
  run?: () => void
}

export interface MenuDef { label: string; items: MenuItem[] }

export interface ToolbarButton {
  kind: 'button'
  id: string
  icon?: string
  label?: string
  tip?: string
  on?: boolean
  disabled?: boolean
  /** dropdown opened by a split arrow or a large ribbon button */
  menu?: MenuItem[]
  run?: () => void
}
export interface ToolbarSep { kind: 'sep' }
export interface ToolbarCombo {
  kind: 'combo'
  id: string
  value: string
  width?: number
  editable?: boolean
  disabled?: boolean
  items?: MenuItem[]
  /** font names render in their own face in the list */
  fontList?: boolean
  onChange?: (value: string) => void
}
export type ToolbarItem = ToolbarButton | ToolbarSep | ToolbarCombo
export interface ToolbarDef { id: string; items: ToolbarItem[] }

export interface RibbonGroup { label: string; items: ToolbarItem[]; rows?: boolean; launch?: () => void }
export interface RibbonTab { id: string; label: string; groups: RibbonGroup[]; ctx?: string }

export interface StatusCell {
  text?: string
  grow?: boolean
  flag?: boolean
  on?: boolean
  act?: boolean
  tip?: string
  run?: () => void
}
