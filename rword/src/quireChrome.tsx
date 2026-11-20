/* Quire chrome for the React editor. Mirrors the DOM and class names of public/quire
 * so the classic looks.css / ui.css / quire.css styles render the interface identically. */
import { createContext, Fragment, useContext, useEffect, useLayoutEffect, useRef, useState, type ReactNode } from 'react'
import { createPortal } from 'react-dom'
import { getIcon } from './icons'
import { ampHTML, keyText, plainAmp, t } from './i18n'
import { COMMAND_META } from './commandMeta'

/** The menu that holds a panel tells it how to close itself (the original's
 * `it.custom(() => ui.closeMenus())`): a panel's action closes the menu first. */
const MenuCloseContext = createContext<() => void>(() => {})

/** Close the menu a panel sits in. Panels call it before they run their action. */
export function useMenuClose(): () => void {
  return useContext(MenuCloseContext)
}

export interface QItem {
  id?: string
  label?: string
  key?: string
  icon?: string
  sep?: boolean
  checked?: boolean
  radio?: boolean
  disabled?: boolean
  bold?: boolean
  sub?: QItem[]
  /** a custom panel rendered in place of a menu row (the original's `custom` items) */
  panel?: ReactNode
  /** a section heading row */
  heading?: string
  /** the row's own tooltip (the original's `it.tip`) */
  tip?: string
  /** the class the menu carrying this row gets (the original's `opts.cls`) */
  cls?: string
  run?: () => void
}
export interface QMenu { label: string; items: QItem[] }
export interface QTool {
  id?: string
  icon?: string
  label?: string
  tip?: string
  sep?: boolean
  disabled?: boolean
  checked?: boolean
  /** a dropdown arrow next to the button (the original's `tbDrop`) */
  drop?: QItem[]
  /** a split button: the main action plus a dropdown arrow (the original's `tbSplit`) */
  split?: { items: QItem[] }
  combo?: { id?: string; value: string; options: string[]; width?: number; tip?: string; onChange?: (value: string) => void }
  run?: () => void
}

/* Command ids whose icon name differs from the id, and human labels. */
const ICON_ALIAS: Record<string, string> = {
  printPreview: 'preview', spelling: 'spell', research: 'research', officeClipboard: 'clipboard',
  pasteSpecial: 'paste', pasteHyperlink: 'hyperlink', clearFormats: 'resetGraphic', clearContents: 'delete',
  goTo: 'find', tb_tables: 'tablesBorders', tb_drawing: 'drawing', insertExcel: 'excel',
  columns: 'columns', viewReading: 'read', viewNormal: 'normalView', viewWeb: 'webView',
  viewPrint: 'printView', viewOutline: 'outlineView', browsePrev: 'browseUp', browseNext: 'browseDown',
  browseObj: 'browseObj', changeCase: 'changeCase', lineSpacing: 'lineSpacing', borderBtn: 'bordersAll',
  highlightBtn: 'highlight', fontColorBtn: 'fontColor', stylesPane: 'stylesPane', revealFormatting: 'reveal',
  focusFont: 'font', focusSize: 'font', dateTime: 'date', pageNumbers: 'pageNum',
  drawLine: 'line', drawArrow: 'arrowStyle', drawRect: 'rect', drawOval: 'oval',
  textBox: 'textbox', insertWordArt: 'wordart', insertDiagram: 'diagram',
  insertClipArt: 'clipart', selectObjects: 'select',
  formatObject: 'formatPicture', wordartEdit: 'wordart',
}
export const LABELS: Record<string, string> = {
  new: 'New', open: 'Open', save: 'Save', saveAs: 'Save As', saveWeb: 'Save to Web', print: 'Print',
  printPreview: 'Print Preview', spelling: 'Spelling and Grammar', research: 'Research', cut: 'Cut',
  copy: 'Copy', paste: 'Paste', painter: 'Format Painter', undo: 'Undo', redo: 'Redo',
  hyperlink: 'Insert Hyperlink', docMap: 'Document Map', showMarks: 'Show/Hide ¶', help: 'Help',
  stylesPane: 'Styles and Formatting', bold: 'Bold', italic: 'Italic', underline: 'Underline',
  alignLeft: 'Align Left', alignCenter: 'Center', alignRight: 'Align Right', justify: 'Justify',
  lineSpacing: 'Line Spacing', numbering: 'Numbering', bullets: 'Bullets', decIndent: 'Decrease Indent',
  incIndent: 'Increase Indent', borderBtn: 'Borders', highlightBtn: 'Highlight', fontColorBtn: 'Font Color',
  viewNormal: 'Normal', viewWeb: 'Web Layout', viewPrint: 'Print Layout', viewOutline: 'Outline',
  viewReading: 'Reading Layout', browsePrev: 'Previous Page', browseNext: 'Next Page',
  browseObj: 'Select Browse Object', insertTable: 'Insert Table', drawTable: 'Draw Table',
  selectObjects: 'Select Objects', drawLine: 'Line', drawArrow: 'Arrow', drawRect: 'Rectangle',
  drawOval: 'Oval', textBox: 'Text Box', insertWordArt: 'Insert LettersArt', insertDiagram: 'Insert Diagram',
  insertClipArt: 'Insert Clip Art', insertPicture: 'Insert Picture', fillColor: 'Fill Color',
  lineColor: 'Line Color', lineStyle: 'Line Style', dashStyle: 'Dash Style', arrowStyle: 'Arrow Style',
  shadowStyle: 'Shadow Style', autoShapes: 'AutoShapes', draw: 'Draw',
  autoshapes: 'AutoShapes', eraser: 'Eraser', table: 'Insert Table',
  bordersAll: 'Borders', textWrap: 'Text Wrapping', colorPic: 'Color',
  pvMag: 'Magnifier', onePage: 'One Page', multiPage: 'Multiple Pages', shrink: 'Shrink to Fit',
  closePreview: 'Close Preview', ruler: 'Ruler', fullScreen: 'Full Screen', close: 'Close',
}
export function iconFor(id: string): string { return ICON_ALIAS[id] || id }
export function labelFor(id: string): string {
  if (LABELS[id]) return LABELS[id]
  const bare = id.replace(/^tb_/, '')
  return bare.charAt(0).toUpperCase() + bare.slice(1)
}

/* The original's command registry: each command's own label (its menu name and the
 * name its button announces) and its shortcut. The toolbars read them from here, so
 * a button announces what the original announces. */
const COMMAND_LABELS: Record<string, string> = {}
const COMMAND_KEYS: Record<string, string> = {}
for (const meta of COMMAND_META) {
  COMMAND_LABELS[meta.id] = meta.label || meta.id
  if (meta.key) COMMAND_KEYS[meta.id] = meta.key
}

export function QIcon({ name, size = 16 }: { name: string; size?: number }): ReactNode {
  return <span aria-hidden="true" dangerouslySetInnerHTML={{ __html: getIcon(name, size) }} />
}

const amp = (s: string) => ampHTML(s)

/** The tip a button shows: its command's tip, with the shortcut appended. */
function buttonTip(item: QTool): string {
  const tip = plainAmp(t(item.tip || item.label || (item.id ? labelFor(item.id) : '')))
  const key = COMMAND_KEYS[item.id || '']
  return key ? `${tip} (${keyText(key)})` : tip
}

/** The name a button announces: the command's own label (the original's `stripAmp(c.label)`). */
function buttonLabel(item: QTool): string {
  const id = item.id || ''
  return plainAmp(t(COMMAND_LABELS[id] || item.tip || item.label || labelFor(id)))
}

/** A split button's arrow names the command it belongs to, not the shortcut
 * (the original's `ui.tbSplit`: the arrow's data-tip is the command's label). */
function splitArrowTip(item: QTool): string {
  const id = item.id || ''
  return plainAmp(t(COMMAND_LABELS[id] || item.tip || item.label || labelFor(id)))
}

function QMenuPopup({ items, x, y, onClose, cls }: { items: QItem[]; x: number; y: number; onClose: () => void; cls?: string }): ReactNode {
  const [active, setActive] = useState<number | null>(null)
  const [sub, setSub] = useState<{ x: number; y: number; items: QItem[] } | null>(null)
  const ref = useRef<HTMLDivElement>(null)
  /* The original Quire appends menus to the body and clamps them into the
   * viewport; do the same after mount so a popup never lands off-screen. */
  useLayoutEffect(() => {
    const el = ref.current
    if (!el) return
    const rect = el.getBoundingClientRect()
    let left = x
    let top = y
    if (left + rect.width > window.innerWidth - 2) left = Math.max(2, window.innerWidth - rect.width - 2)
    if (top + rect.height > window.innerHeight - 2) top = Math.max(2, window.innerHeight - rect.height - 2)
    if (left !== x || top !== y) { el.style.left = `${left}px`; el.style.top = `${top}px` }
  }, [x, y])
  useEffect(() => {
    const down = (e: MouseEvent) => { if (!ref.current?.contains(e.target as Node) && !(e.target as HTMLElement).closest('.mb-item')) onClose() }
    const key = (e: KeyboardEvent) => { if (e.key === 'Escape') onClose() }
    document.addEventListener('mousedown', down); document.addEventListener('keydown', key)
    return () => { document.removeEventListener('mousedown', down); document.removeEventListener('keydown', key) }
  }, [onClose])
  const keys = (e: React.KeyboardEvent) => {
    const usable = items.map((it, i) => (!it.sep && !it.disabled && it.label ? i : -1)).filter((i) => i >= 0)
    if (!usable.length) return
    const at = usable.indexOf(active ?? -1)
    if (e.key === 'ArrowDown') { e.preventDefault(); setActive(usable[(at + 1) % usable.length]) }
    else if (e.key === 'ArrowUp') { e.preventDefault(); setActive(usable[(at - 1 + usable.length) % usable.length]) }
    else if (e.key === 'Enter' && active != null) { e.preventDefault(); const it = items[active]; if (it.sub?.length) return; onClose(); void it.run?.() }
    else if (e.key === 'ArrowLeft') onClose()
  }
  /* the original lets a menu carry a class (a grid or colour panel's menu) */
  const menuCls = cls || items.find((it) => it.cls)?.cls || ''
  return createPortal(<MenuCloseContext.Provider value={onClose}>
    <div ref={ref} className={`menu${menuCls ? ' ' + menuCls : ''}`} role="menu" style={{ left: x, top: y }} tabIndex={-1} onKeyDown={keys}>
      {items.map((it, i) => it.sep
        ? <div className="m-sep" key={`s${i}`} />
        : it.panel !== undefined
        /* the original appends a custom panel straight into the menu */
        ? <Fragment key={`p${i}`}>{it.panel}</Fragment>
        : it.heading
        ? <div key={`h${i}`} className="m-head">{it.heading}</div>
        : <div key={`${it.label}-${i}`} className={`m-item${active === i ? ' act' : ''}${it.disabled ? ' dis' : ''}${it.bold ? ' bold' : ''}${it.sub?.length ? ' has-sub' : ''}`}
            role="menuitem" data-cmd={it.id || undefined} aria-disabled={it.disabled || undefined}
            title={it.tip || undefined}
            onMouseEnter={(e) => { setActive(i); if (it.sub?.length) { const r = (e.currentTarget as HTMLElement).getBoundingClientRect(); setSub({ x: r.right - 2, y: r.top - 2, items: it.sub }) } else setSub(null) }}
            onMouseDown={(e) => e.preventDefault()}
            onClick={() => { if (it.disabled || it.sub?.length) return; onClose(); void it.run?.() }}>
            <span className="m-ico">{it.checked ? <span className={it.radio ? 'm-radio' : 'm-check'}>{it.radio ? '●' : '✓'}</span> : it.icon ? <QIcon name={iconFor(it.icon)} /> : null}</span>
            <span className="m-lbl" dangerouslySetInnerHTML={amp(it.label || '')} />
            <span className="m-key">{it.key ? keyText(it.key) : ''}</span>
            {it.sub?.length ? <span className="m-sub">▶</span> : null}
          </div>)}
    </div>
    {sub && <QMenuPopup items={sub.items} x={sub.x} y={sub.y} onClose={onClose} />}
  </MenuCloseContext.Provider>, document.body)
}

export function QMenuBar({ menus, onSearch }: { menus: QMenu[]; onSearch?: (query: string) => void }): ReactNode {
  const [open, setOpen] = useState<number | null>(null)
  const [pos, setPos] = useState({ x: 0, y: 0 })
  const [query, setQuery] = useState('')
  useEffect(() => {
    const down = () => setOpen(null)
    document.addEventListener('mousedown', down)
    return () => document.removeEventListener('mousedown', down)
  }, [])
  return <div className="menurow">
    <div className="menubar" role="menubar">
      {menus.map((m, i) => <div key={m.label} className={`mb-item${open === i ? ' open' : ''}`} role="menuitem" tabIndex={-1}
        onMouseDown={(e) => { e.stopPropagation(); const r = (e.currentTarget as HTMLElement).getBoundingClientRect(); setPos({ x: r.left, y: r.bottom }); setOpen(open === i ? null : i) }}
        onMouseEnter={(e) => { if (open !== null && open !== i) { const r = (e.currentTarget as HTMLElement).getBoundingClientRect(); setPos({ x: r.left, y: r.bottom }); setOpen(i) } }}
        dangerouslySetInnerHTML={amp(m.label)} />)}
    </div>
    {open !== null && <QMenuPopup items={menus[open].items} x={pos.x} y={pos.y} onClose={() => setOpen(null)} />}
    <div className="cmd-search"><span className="cs-glass">⌕</span><input type="search" value={query} placeholder={t('Search commands')}
      onChange={(e) => setQuery(e.target.value)} onKeyDown={(e) => { if (e.key === 'Enter') { onSearch?.(query); setQuery('') } else e.stopPropagation() }} /></div>
  </div>
}

export function QToolButton({ item }: { item: QTool }): ReactNode {
  const [drop, setDrop] = useState(false)
  const [dropPos, setDropPos] = useState({ x: 0, y: 0 })
  const anchor = useRef<HTMLDivElement>(null)
  const tip = buttonTip(item)
  const aria = buttonLabel(item)
  if (item.sep) return <span className="tb-sep" />
  if (item.combo) return <div ref={anchor} className="combo" style={{ width: item.combo.width || 100 }} data-tip={item.combo.tip ? plainAmp(t(item.combo.tip)) : undefined}>
    <input className="combo-in" id={item.combo.id} value={item.combo.value} onChange={(e) => item.combo!.onChange?.(e.target.value)}
      onKeyDown={(e) => { if (e.key === 'Enter') { item.combo!.onChange?.((e.target as HTMLInputElement).value); setDrop(false) } if (e.key === 'Escape') setDrop(false) }} />
    <button className="combo-btn" type="button" aria-label="Open list" onMouseDown={(e) => e.preventDefault()} onClick={() => { const r = anchor.current?.getBoundingClientRect(); if (r) setDropPos({ x: r.left, y: r.bottom }); setDrop((v) => !v) }}><span className="dd-arrow" /></button>
    {drop && <QMenuPopup cls="combo-list" x={dropPos.x} y={dropPos.y} onClose={() => setDrop(false)} items={item.combo.options.map((o) => ({ label: o, checked: o === item.combo!.value, run: () => item.combo!.onChange?.(o) }))} />}
  </div>
  const face = <>{item.icon || item.id ? <QIcon name={item.icon || iconFor(item.id || '')} /> : null}{item.label ? <span className="tb-lbl" dangerouslySetInnerHTML={ampHTML(item.label)} /> : null}</>
  /* split button: the action, then a separate arrow that opens the menu */
  if (item.split) return <div className="tb-split" data-cmd={item.id}>
    <button className="tb-btn" type="button" title={tip} data-tip={tip} data-cmd={item.id} aria-label={aria} disabled={item.disabled} onMouseDown={(e) => e.preventDefault()} onClick={() => void item.run?.()}>{face}</button>
    <button className="tb-dd" type="button" aria-label={t('More options')} data-tip={splitArrowTip(item)} onMouseDown={(e) => e.preventDefault()} onClick={(e) => {
      /* the original opens no menu when the list is empty (the undo history, say) */
      if (!item.split!.items.length) return
      const r = (e.currentTarget as HTMLElement).getBoundingClientRect(); setDropPos({ x: r.left, y: r.bottom }); setDrop((v) => !v)
    }}><span className="dd-arrow" /></button>
    {drop && <QMenuPopup x={dropPos.x} y={dropPos.y} onClose={() => setDrop(false)} items={item.split.items} />}
  </div>
  /* dropdown button: the icon carries its own arrow (the original's tbDrop names
   * the button after its tip, not after a command) */
  if (item.drop) return <button className={`tb-btn tb-drop${item.label ? ' with-label' : ''}`} type="button" title={tip} data-tip={tip} aria-label={tip}
    disabled={item.disabled} onMouseDown={(e) => e.preventDefault()}
    onClick={(e) => { const r = (e.currentTarget as HTMLElement).getBoundingClientRect(); setDropPos({ x: r.left, y: r.bottom }); setDrop((v) => !v) }}>
    {face}<span className="dd-arrow" />
    {drop && <QMenuPopup x={dropPos.x} y={dropPos.y} onClose={() => setDrop(false)} items={item.drop} />}
  </button>
  return <button className={`tb-btn${item.checked ? ' on' : ''}`} type="button" title={tip} data-tip={tip} data-cmd={item.id} aria-label={aria}
    disabled={item.disabled} onMouseDown={(e) => e.preventDefault()} onClick={() => void item.run?.()}>{face}</button>
}

export function QToolbar({ id, label, items, hidden }: { id: string; label: string; items: QTool[]; hidden?: boolean }): ReactNode {
  return <div className="toolbar" id={`tb-${id}`} role="toolbar" aria-label={label} hidden={hidden}>
    <span className="tb-grip" />
    {items.map((it, i) => <QToolButton key={`${it.id || it.label || 'x'}-${i}`} item={it} />)}
    <span className="tb-end" />
  </div>
}

export function QViewBar({ view, onView, onBrowse, zoom }: { view: string; onView: (v: string) => void; onBrowse?: (dir: number) => void; zoom: number }): ReactNode {
  const views: [string, string, string, string][] = [
    ['normal', 'normalView', '&Normal', 'Ctrl+Alt+N'],
    ['web', 'webView', '&Web Layout', ''],
    ['print', 'printView', '&Print Layout', 'Ctrl+Alt+P'],
    ['outline', 'outlineView', '&Outline', 'Ctrl+Alt+O'],
    ['reading', 'read', 'Read', ''],
  ]
  return <div className="viewbar">
    {views.map(([id, icon, label, key]) => {
      const tip = t(label).replace(/&(.)/, '$1') + (key ? ` (${key})` : '')
      return <button key={id} className={`tb-btn${view === id ? ' on' : ''}`} data-tip={tip} title={tip} aria-label={tip} onClick={() => onView(id)}><QIcon name={icon} /></button>
    })}
    <span className="vb-sp" />
    <span className="browse">
      <button className="tb-btn" data-tip={t('Previous Page')} title={t('Previous Page')} aria-label={t('Previous Page')} onClick={() => onBrowse?.(-1)}><QIcon name="browseUp" /></button>
      <button className="tb-btn tb-drop" data-tip={t('Select Browse Object')} title={t('Select Browse Object')} aria-label={t('Select Browse Object')}><QIcon name="browseObj" /><span className="dd-arrow" /></button>
      <button className="tb-btn" data-tip={t('Next Page')} title={t('Next Page')} aria-label={t('Next Page')} onClick={() => onBrowse?.(1)}><QIcon name="browseDown" /></button>
    </span>
  </div>
}

export function QTaskPane({ title, onNav, children, hidden }: { title: string; onNav: (action: string) => void; children: ReactNode; hidden?: boolean }): ReactNode {
  return <div className="taskpane" id="taskpane" hidden={hidden}>
    <div className="tp-head">
      <button className="tp-nav" title={t('Back')} disabled><QIcon name="back" /></button>
      <button className="tp-nav" title={t('Forward')} disabled><QIcon name="forward" /></button>
      <button className="tp-nav" title={t('Home')} aria-label={t('Home')} onClick={() => onNav('home')}><QIcon name="home" /></button>
      <div className="tp-title" role="button" tabIndex={0} onClick={() => onNav('menu')}><span>{title}</span><span className="dd-arrow" /></div>
      <button className="tp-nav" title={t('Close')} aria-label={t('Close')} onClick={() => onNav('close')}>×</button>
    </div>
    <div className="tp-body">{children}</div>
  </div>
}

export function QDialog({ title, children, onClose, wide }: { title: string; children: ReactNode; onClose: () => void; wide?: boolean }): ReactNode {
  const ref = useRef<HTMLDivElement>(null)
  useEffect(() => {
    const key = (e: KeyboardEvent) => { if (e.key === 'Escape') onClose() }
    document.addEventListener('keydown', key); ref.current?.focus()
    return () => document.removeEventListener('keydown', key)
  }, [onClose])
  return <div className="dlg-overlay" role="presentation">
    <div className={`dlg${wide ? ' wide' : ''}`} role="dialog" aria-modal="true" aria-label={title} tabIndex={-1} ref={ref}>
      <div className="dlg-title">{title}<button className="dlg-x" type="button" aria-label={t('Close')} onClick={onClose}>×</button></div>
      <div className="dlg-body">{children}</div>
    </div>
  </div>
}
