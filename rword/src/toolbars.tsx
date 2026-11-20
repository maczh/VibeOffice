/* The toolbars, built to the original Quire's specification.
 *
 * public/quire/js/app.js → buildToolbars() defines ten toolbars. Each entry is
 * either a command id (a plain button), '|' (a separator), a combo box, a split
 * button (an action plus a dropdown arrow) or a dropdown button (an icon with an
 * arrow). This module mirrors that list item for item, so the React editor shows
 * the same buttons, in the same order, with the same menus. */
import type { ReactNode } from 'react'
import { type QItem, type QTool, useMenuClose } from './quireChrome'
import { META_BY_ID } from './commandMeta'
import {
  BORDER_STYLES, DASH_STYLES, HEADER_FOOTER_AUTOTEXT, LINE_WEIGHTS, WORDART_SHAPES, WRAP_MODES,
} from './services/toolbarCommands'
import {
  QBorderGrid, QColorPanel, QColumnPicker, QHighlightPanel, QShadowGrid, QShapeGallery, QStyleList, QTableGridPicker, lineSVG,
} from './panels'
import { plainAmp, t } from './i18n'

/** The interface the toolbar builder needs from the editor. */
export interface ToolbarEnv {
  /** run a command by id */
  run: (id: string) => void
  /** run a command with an argument */
  runArg: (id: string, arg: string) => void
  /** a plain item that opens the command's dialog */
  item: (id: string) => QItem
  /** the current zoom, as the zoom combo shows it */
  zoom: number
  /** the current style, font and size */
  style: string
  font: string
  size: string
  /** the current outline level and shown level */
  outlineLevel: number
  outlineShow: number
  /** the current mail-merge record */
  mergeRecord: number
  mergeRecords: number
  /** the picture colour mode currently applied */
  picMode: string
  /** the wrap mode currently applied */
  wrapMode: string
  /** the undo and redo history, newest last */
  undoLabels: string[]
  redoLabels: string[]
}

/** The label the original's command registry gives an id, for menu rows. */
function cmdLabel(id: string): string {
  return META_BY_ID[id]?.label || id
}

/** The shortcut the original's command registry gives an id, for a menu row's key column. */
function cmdKey(id: string): string | undefined {
  return META_BY_ID[id]?.key
}

/** A row of the line-style list panel. */
export interface StyleRow { html: string; label?: string; run: () => void }

/* The labels the original shows for the tools that are not commands. */
const TOOL_LABELS: Record<string, string> = {
  tb_tables: 'Tables and Borders',
  tb_drawing: 'Drawing',
}

/** A toolbar button for a command id, carrying the original's icon and tip. */
export function toolButton(env: ToolbarEnv, id: string, run?: () => void): QTool {
  const meta = META_BY_ID[id]
  return {
    id,
    icon: meta?.icon,
    tip: meta?.tip || meta?.label || TOOL_LABELS[id],
    label: undefined,
    run: run || (() => env.run(id)),
  }
}

/** A dropdown button (the original's `tbDrop`): an icon with its own arrow. */
function dropButton(env: ToolbarEnv, id: string, items: QItem[], tip?: string, icon?: string, label?: string): QTool {
  const meta = META_BY_ID[id]
  return {
    id,
    icon: icon || meta?.icon,
    label,
    tip: tip || meta?.tip || meta?.label || TOOL_LABELS[id],
    drop: items,
  }
}

/** A split button (the original's `tbSplit`): an action plus a dropdown arrow. */
function splitButton(env: ToolbarEnv, id: string, items: QItem[], tip?: string): QTool {
  const meta = META_BY_ID[id]
  return {
    id,
    icon: meta?.icon,
    tip: tip || meta?.tip || meta?.label || TOOL_LABELS[id],
    run: () => env.run(id),
    split: { items },
  }
}

/* --------------------------------------------------------------- the menus */

/** The undo/redo history list (the original's historyMenu). The original opens no
 *  menu at all when there is nothing to undo, so the arrow does nothing then. */
function historyPanel(env: ToolbarEnv, kind: 'undo' | 'redo'): QItem[] {
  const list = kind === 'undo' ? env.undoLabels : env.redoLabels
  if (!list.length) return []
  return [{
    panel: (
      <HistoryList labels={list.slice(0, 40)} kind={kind}
        onPick={(count) => { for (let i = 0; i < count; i++) env.run(kind) }} />
    ),
  }]
}

function HistoryList({ labels, kind, onPick }: { labels: string[]; kind: 'undo' | 'redo'; onPick: (count: number) => void }): ReactNode {
  const action = (n: number) => n === 1
    ? t(kind === 'undo' ? 'Undo {0} Action' : 'Redo {0} Action', 1)
    : t(kind === 'undo' ? 'Undo {0} Actions' : 'Redo {0} Actions', 1)
  const close = useMenuClose()
  return <div style={{ background: '#fff' }}>
    <div className="style-list" style={{ maxHeight: 260, overflow: 'auto', minWidth: 180 }}>
      {labels.map((label, i) => <button key={i} type="button" title={label}
        onMouseEnter={(e) => {
          const row = e.currentTarget.parentElement
          if (!row) return
          Array.from(row.children).forEach((x, k) => {
            const el = x as HTMLElement
            el.style.background = k <= i ? '#316ac5' : ''
            el.style.color = k <= i ? '#fff' : ''
          })
        }}
        onClick={() => { close(); onPick(i + 1) }}>{label}</button>)}
    </div>
    <div className="tgp-label">{action(1)}</div>
  </div>
}

/** The line-spacing menu: five multiples plus "More...". */
function lineSpacingMenu(env: ToolbarEnv): QItem[] {
  const items: QItem[] = [1, 1.5, 2, 2.5, 3].map((value) => ({
    label: value.toFixed(1),
    checked: false,
    run: () => env.runArg('setLineSpacing', String(value)),
  }))
  items.push({ sep: true, label: '' })
  items.push({ label: '&More...', run: () => env.run('paragraphDlg') })
  return items
}

/** The border grid menu (the original's bordersDD). */
function borderMenu(env: ToolbarEnv): QItem[] {
  return [{ cls: 'grid-menu', panel: <QBorderGrid onPick={(key) => env.runArg('applyBorder', key)} /> }]
}

/** The highlight panel menu (the original's highlightDD). */
function highlightMenu(env: ToolbarEnv): QItem[] {
  return [{ cls: 'color-menu', panel: <QHighlightPanel onPick={(name) => env.runArg('applyHighlight', name)} /> }]
}

/** The font colour menu (the original's fontColorDD). */
function fontColorMenu(env: ToolbarEnv): QItem[] {
  return [{
    cls: 'color-menu',
    panel: <QColorPanel mode="font" onPick={(pick) => {
      if (pick.auto) env.runArg('applyFontColor', 'auto')
      else if (pick.hex) env.runArg('applyFontColor', pick.hex)
    }} />,
  }]
}

/** The border colour menu: automatic plus the standard colours. */
function borderColorMenu(env: ToolbarEnv): QItem[] {
  return [{
    cls: 'color-menu',
    panel: <QColorPanel mode="plain" onPick={(pick) => {
      if (pick.auto) env.runArg('setBorderColor', 'auto')
      else if (pick.hex) env.runArg('setBorderColor', pick.hex.replace('#', '').toUpperCase())
    }} />,
  }]
}

/** The shading colour menu. */
function shadingColorMenu(env: ToolbarEnv): QItem[] {
  return [{
    cls: 'color-menu',
    panel: <QColorPanel mode="fill" onPick={(pick) => {
      if (pick.none) env.runArg('applyShading', 'none')
      else if (pick.hex) env.runArg('applyShading', pick.hex)
    }} />,
  }]
}

/** The line-style list (the original's borderStyleMenu). */
function borderStyleMenu(env: ToolbarEnv): QItem[] {
  const rows: StyleRow[] = BORDER_STYLES.map(([key, dash]) => ({
    html: key === 'double'
      ? '<svg width="90" height="14"><path d="M2 5h86M2 9h86" stroke="#000"/></svg>'
      : key === 'triple'
        ? '<svg width="90" height="14"><path d="M2 3h86M2 7h86M2 11h86" stroke="#000"/></svg>'
        : key === 'wave'
          ? '<svg width="90" height="14"><path d="M2 7q5-5 10 0t10 0t10 0t10 0t10 0t10 0t10 0t10 0t6 0" fill="none" stroke="#000"/></svg>'
          : lineSVG(key === 'thick' ? 3 : 1, dash || undefined),
    run: () => env.runArg('setBorderStyle', key),
  }))
  rows.push({ html: '', label: 'No Border', run: () => env.runArg('setBorderStyle', 'nil') })
  return [{ panel: <QStyleList rows={rows} /> }]
}

/** The line-weight list (the original's borderWeightMenu). */
function borderWeightMenu(env: ToolbarEnv): QItem[] {
  const rows: StyleRow[] = LINE_WEIGHTS.map((w) => ({
    html: lineSVG(Math.max(0.5, w * 1.33)),
    label: t('{0} pt', String(w)),
    run: () => env.runArg('setBorderWeight', String(w)),
  }))
  return [{ panel: <QStyleList rows={rows} /> }]
}

/** The object line-weight list (the original's objLineWeightMenu). */
function objectWeightMenu(env: ToolbarEnv): QItem[] {
  const rows: StyleRow[] = LINE_WEIGHTS.map((w) => ({
    html: lineSVG(Math.max(0.5, w * 1.33)),
    label: t('{0} pt', String(w)),
    run: () => env.runArg('setLineStyle', String(w)),
  }))
  rows.push({ html: '', label: 'More Lines...', run: () => env.run('formatObject') })
  return [{ panel: <QStyleList rows={rows} /> }]
}

/** The object dash-style list (the original's objDashMenu). */
function objectDashMenu(env: ToolbarEnv): QItem[] {
  const rows: StyleRow[] = DASH_STYLES.map(([key, dash]) => ({
    html: lineSVG(2, dash || undefined),
    run: () => env.runArg('setDashStyle', key),
  }))
  return [{ panel: <QStyleList rows={rows} /> }]
}

/** The object arrow-style list (the original's objArrowMenu). */
function objectArrowMenu(env: ToolbarEnv): QItem[] {
  const rows: StyleRow[] = []
  const shapes: [string | null, string | null][] = [
    [null, null], [null, 'triangle'], ['triangle', null], ['triangle', 'triangle'],
    [null, 'arrow'], [null, 'stealth'], [null, 'oval'], [null, 'diamond'], ['oval', 'oval'],
  ]
  const head = (kind: string | null, x: number, dir: number): string => {
    if (!kind) return ''
    if (kind === 'oval') return `<circle cx="${x}" cy="7" r="3"/>`
    if (kind === 'diamond') return `<path d="M${x - 4} 7l4-4 4 4-4 4z"/>`
    if (kind === 'arrow') return `<path d="M${x - dir * 6} 2l${dir * 6} 5-${dir * 6} 5" fill="none" stroke="#000" stroke-width="1.5"/>`
    return `<path d="M${x} 7l${-dir * 8} -4v8z"/>`
  }
  for (const [a, b] of shapes) {
    rows.push({
      html: `<svg width="90" height="14" viewBox="0 0 90 14" fill="#000"><path d="M6 7h78" stroke="#000" stroke-width="1.5"/>${head(a, 4, -1)}${head(b, 86, 1)}</svg>`,
      run: () => env.runArg('setArrowStyle', `${a || ''}|${b || ''}`),
    })
  }
  return [{ panel: <QStyleList rows={rows} /> }]
}

/** The object shadow grid (the original's objShadowMenu). */
function objectShadowMenu(env: ToolbarEnv): QItem[] {
  return [{ panel: <QShadowGrid onPick={(offset) => env.runArg('setShadowStyle', offset ? `${offset[0]},${offset[1]}` : 'none')} /> }]
}

/** The table grid picker (the original's tableGridPicker). */
function tableGridMenu(env: ToolbarEnv): QItem[] {
  return [{ cls: 'grid-menu', panel: <QTableGridPicker onPick={(rows, cols) => env.runArg('insertTableAt', `${rows}x${cols}`)} /> }]
}

/** The column picker (the original's columnsPicker). */
function columnsMenu(env: ToolbarEnv): QItem[] {
  return [{ cls: 'grid-menu', panel: <QColumnPicker onPick={(count) => env.runArg('setColumns', String(count))} /> }]
}

/** The picture colour menu (the original's picture colour dropdown). */
function pictureColorMenu(env: ToolbarEnv): QItem[] {
  return [
    { label: '&Automatic', run: () => env.runArg('setPicMode', '') },
    { label: '&Grayscale', run: () => env.runArg('setPicMode', 'gray') },
    { label: '&Black & White', run: () => env.runArg('setPicMode', 'bw') },
    { label: '&Washout', run: () => env.runArg('setPicMode', 'washout') },
  ]
}

/** The text-wrapping menu (the original's wrapItems). */
function wrapMenu(env: ToolbarEnv): QItem[] {
  const items: QItem[] = WRAP_MODES.map(([key, label]) => ({
    label,
    checked: env.wrapMode === key,
    run: () => env.runArg('setWrap', key),
  }))
  items.push({ sep: true, label: '' })
  items.push({ label: '&Edit Wrap Points', disabled: true })
  return items
}

/** The LettersArt shape menu. */
function wordArtShapeMenu(env: ToolbarEnv): QItem[] {
  return WORDART_SHAPES.map(([key, label]) => ({
    label,
    run: () => env.runArg('setWordArtShape', key),
  }))
}

/** The header/footer AutoText menu. */
function autoTextMenu(env: ToolbarEnv): QItem[] {
  return HEADER_FOOTER_AUTOTEXT.map(([label, text]) => ({
    label,
    run: () => env.runArg('insertAutoText', text),
  }))
}

/** The review display menu (the original's markup dropdown). */
function markupMenu(env: ToolbarEnv): QItem[] {
  return [
    { id: 'markupFinalMarkup', label: cmdLabel('markupFinalMarkup'), key: cmdKey('markupFinalMarkup'), checked: true, run: () => env.run('markupFinalMarkup') },
    { id: 'markupFinal', label: cmdLabel('markupFinal'), key: cmdKey('markupFinal'), run: () => env.run('markupFinal') },
    { id: 'markupOriginalMarkup', label: cmdLabel('markupOriginalMarkup'), key: cmdKey('markupOriginalMarkup'), run: () => env.run('markupOriginalMarkup') },
    { id: 'markupOriginal', label: cmdLabel('markupOriginal'), key: cmdKey('markupOriginal'), run: () => env.run('markupOriginal') },
  ]
}

/** The review "Show" menu. */
function showMenu(env: ToolbarEnv): QItem[] {
  return [
    { id: 'showComments', label: cmdLabel('showComments'), key: cmdKey('showComments'), run: () => env.run('showComments') },
    { id: 'showInsDel', label: cmdLabel('showInsDel'), key: cmdKey('showInsDel'), run: () => env.run('showInsDel') },
    { id: 'showFormatting', label: cmdLabel('showFormatting'), key: cmdKey('showFormatting'), run: () => env.run('showFormatting') },
    { sep: true, label: '' },
    { id: 'reviewersMenu', label: cmdLabel('reviewersMenu'), key: cmdKey('reviewersMenu'), run: () => env.run('reviewersMenu') },
    { sep: true, label: '' },
    { id: 'balloonsToggle', label: cmdLabel('balloonsToggle'), key: cmdKey('balloonsToggle'), run: () => env.run('balloonsToggle') },
    { id: 'reviewingPane', label: cmdLabel('reviewingPane'), key: cmdKey('reviewingPane'), run: () => env.run('reviewingPane') },
  ]
}

/** The accept-change menu. */
function acceptChangeMenu(env: ToolbarEnv): QItem[] {
  return [
    { id: 'acceptChange', label: cmdLabel('acceptChange'), key: cmdKey('acceptChange'), run: () => env.run('acceptChange') },
    { id: 'acceptShown', label: cmdLabel('acceptShown'), key: cmdKey('acceptShown'), run: () => env.run('acceptShown') },
    { id: 'acceptAll', label: cmdLabel('acceptAll'), key: cmdKey('acceptAll'), run: () => env.run('acceptAll') },
  ]
}

/** The reject-change menu. */
function rejectChangeMenu(env: ToolbarEnv): QItem[] {
  return [
    { id: 'rejectChange', label: cmdLabel('rejectChange'), key: cmdKey('rejectChange'), run: () => env.run('rejectChange') },
    { id: 'rejectShown', label: cmdLabel('rejectShown'), key: cmdKey('rejectShown'), run: () => env.run('rejectShown') },
    { id: 'rejectAll', label: cmdLabel('rejectAll'), key: cmdKey('rejectAll'), run: () => env.run('rejectAll') },
    { sep: true, label: '' },
    { id: 'deleteAllComments', label: cmdLabel('deleteAllComments'), key: cmdKey('deleteAllComments'), run: () => env.run('deleteAllComments') },
  ]
}

/** The table dropdown menu (the original's tables toolbar insert menu). */
function tableMenu(env: ToolbarEnv): QItem[] {
  return [
    { id: 'insertTable', label: cmdLabel('insertTable'), key: cmdKey('insertTable'), run: () => env.run('insertTable') },
    { sep: true, label: '' },
    { id: 'colLeft', label: cmdLabel('colLeft'), key: cmdKey('colLeft'), run: () => env.run('colLeft') },
    { id: 'colRight', label: cmdLabel('colRight'), key: cmdKey('colRight'), run: () => env.run('colRight') },
    { id: 'rowAbove', label: cmdLabel('rowAbove'), key: cmdKey('rowAbove'), run: () => env.run('rowAbove') },
    { id: 'rowBelow', label: cmdLabel('rowBelow'), key: cmdKey('rowBelow'), run: () => env.run('rowBelow') },
    { id: 'insertCells', label: cmdLabel('insertCells'), key: cmdKey('insertCells'), run: () => env.run('insertCells') },
    { sep: true, label: '' },
    { id: 'autofitContents', label: cmdLabel('autofitContents'), key: cmdKey('autofitContents'), run: () => env.run('autofitContents') },
    { id: 'autofitWindow', label: cmdLabel('autofitWindow'), key: cmdKey('autofitWindow'), run: () => env.run('autofitWindow') },
    { id: 'fixedWidth', label: cmdLabel('fixedWidth'), key: cmdKey('fixedWidth'), run: () => env.run('fixedWidth') },
  ]
}

/** The cell alignment menu (the original's alignTop dropdown). */
function cellAlignMenu(env: ToolbarEnv): QItem[] {
  return [
    { id: 'cellTop', label: cmdLabel('cellTop'), key: cmdKey('cellTop'), icon: 'alignTop', run: () => env.run('cellTop') },
    { id: 'cellMiddle', label: cmdLabel('cellMiddle'), key: cmdKey('cellMiddle'), icon: 'alignMiddle', run: () => env.run('cellMiddle') },
    { id: 'cellBottom', label: cmdLabel('cellBottom'), key: cmdKey('cellBottom'), icon: 'alignBottom', run: () => env.run('cellBottom') },
  ]
}

/** The draw menu (the original's drawMenu). */
function drawMenu(env: ToolbarEnv): QItem[] {
  return [
    { label: '&Group', run: () => env.run('groupObj') },
    { label: '&Ungroup', run: () => env.run('ungroupObj') },
    { sep: true, label: '' },
    { label: '&Order', sub: [
      { id: 'bringFront', label: cmdLabel('bringFront'), key: cmdKey('bringFront'), run: () => env.run('bringFront') },
      { id: 'sendBack', label: cmdLabel('sendBack'), key: cmdKey('sendBack'), run: () => env.run('sendBack') },
      { id: 'bringForward', label: cmdLabel('bringForward'), key: cmdKey('bringForward'), run: () => env.run('bringForward') },
      { id: 'sendBackward', label: cmdLabel('sendBackward'), key: cmdKey('sendBackward'), run: () => env.run('sendBackward') },
      { id: 'bringFrontText', label: cmdLabel('bringFrontText'), key: cmdKey('bringFrontText'), run: () => env.run('bringFrontText') },
      { id: 'sendBehindText', label: cmdLabel('sendBehindText'), key: cmdKey('sendBehindText'), run: () => env.run('sendBehindText') },
    ] },
    { label: '&Grid...', run: () => env.run('layoutDlg') },
    { label: 'Nud&ge', sub: [
      { id: 'nudgeUp', label: cmdLabel('nudgeUp'), key: cmdKey('nudgeUp'), run: () => env.run('nudgeUp') },
      { id: 'nudgeDown', label: cmdLabel('nudgeDown'), key: cmdKey('nudgeDown'), run: () => env.run('nudgeDown') },
      { id: 'nudgeLeft', label: cmdLabel('nudgeLeft'), key: cmdKey('nudgeLeft'), run: () => env.run('nudgeLeft') },
      { id: 'nudgeRight', label: cmdLabel('nudgeRight'), key: cmdKey('nudgeRight'), run: () => env.run('nudgeRight') },
    ] },
    { label: 'Rotate or Fl&ip', sub: [
      { id: 'rotateLeft', label: cmdLabel('rotateLeft'), key: cmdKey('rotateLeft'), run: () => env.run('rotateLeft') },
      { id: 'rotateRight', label: cmdLabel('rotateRight'), key: cmdKey('rotateRight'), run: () => env.run('rotateRight') },
      { sep: true, label: '' },
      { id: 'flipH', label: cmdLabel('flipH'), key: cmdKey('flipH'), run: () => env.run('flipH') },
      { id: 'flipV', label: cmdLabel('flipV'), key: cmdKey('flipV'), run: () => env.run('flipV') },
    ] },
    { label: '&Text Wrapping', sub: wrapMenu(env) },
    { sep: true, label: '' },
    { id: 'setDefaults', label: cmdLabel('setDefaults'), key: cmdKey('setDefaults'), run: () => env.run('setDefaults') },
  ]
}

/** The AutoShapes menu (the original's openAutoShapes). */
function autoShapesMenu(env: ToolbarEnv): QItem[] {
  /* the original's grid is a custom panel inside each group's submenu */
  const grid = (group: string) => ({
    panel: <QShapeGallery group={group} onPick={(geom) => env.runArg('insertShape', geom)} />,
  })
  return [
    { label: '&Lines', icon: 'line', sub: [
      { id: 'drawLine', label: cmdLabel('drawLine'), key: cmdKey('drawLine'), run: () => env.run('drawLine') },
      { id: 'drawArrow', label: cmdLabel('drawArrow'), key: cmdKey('drawArrow'), run: () => env.run('drawArrow') },
      { id: 'drawDblArrow', label: cmdLabel('drawDblArrow'), key: cmdKey('drawDblArrow'), run: () => env.run('drawDblArrow') },
    ] },
    { label: '&Basic Shapes', icon: 'autoshapes', sub: [grid('Basic Shapes')] },
    { label: '&Block Arrows', icon: 'arrowStyle', sub: [grid('Block Arrows')] },
    { label: '&Flowchart', icon: 'diagram', sub: [grid('Flowchart')] },
    { label: '&Stars and Banners', icon: 'animation', sub: [grid('Stars and Banners')] },
    { label: '&Callouts', icon: 'comment', sub: [grid('Callouts')] },
    { sep: true, label: '' },
    { label: '&More AutoShapes...', icon: 'clipart', run: () => env.run('insertClipArt') },
  ]
}

/* ------------------------------------------------------------ the toolbars */

export function buildToolbars(env: ToolbarEnv): {
  standard: QTool[]
  formatting: QTool[]
  tables: QTool[]
  reviewing: QTool[]
  outlining: QTool[]
  headerFooter: QTool[]
  picture: QTool[]
  mailMerge: QTool[]
  wordart: QTool[]
  drawing: QTool[]
} {
  const sep: QTool = { sep: true }
  const btn = (id: string): QTool => toolButton(env, id)

  const standard: QTool[] = [
    btn('new'), btn('open'), btn('save'), btn('permission'), btn('sendMail'), btn('print'), btn('printPreview'),
    btn('spelling'), btn('research'), sep,
    btn('cut'), btn('copy'), btn('paste'), btn('painter'), sep,
    splitButton(env, 'undo', historyPanel(env, 'undo')),
    splitButton(env, 'redo', historyPanel(env, 'redo')),
    sep,
    btn('hyperlink'),
    { id: 'tb_tables', icon: 'tablesBorders', tip: 'Tables and Borders', run: () => env.run('tb_tables') },
    dropButton(env, 'table', tableGridMenu(env), 'Insert Table', 'table'),
    btn('insertExcel'),
    dropButton(env, 'columns', columnsMenu(env), 'Columns', 'columns'),
    { id: 'tb_drawing', icon: 'drawing', tip: 'Drawing', run: () => env.run('tb_drawing') },
    btn('docMap'), btn('showMarks'),
    { combo: { id: 'tb-zoom', tip: 'Zoom', value: `${Math.round(env.zoom * 100)}%`, width: 74, options: ['500%', '200%', '150%', '100%', '75%', '50%', '25%', '10%', 'Page Width', 'Text Width', 'Whole Page', 'Two Pages'], onChange: (v) => env.runArg('setZoom', v) } },
    btn('help'),
    { id: 'viewReading', icon: 'read', label: 'Read', tip: 'Read Mode', run: () => env.run('viewReading') },
  ]

  const formatting: QTool[] = [
    btn('stylesPane'),
    { combo: { id: 'tb-style', tip: 'Style', value: env.style, width: 118, options: ['Normal', 'Heading 1', 'Heading 2', 'Heading 3', 'Title', 'Subtitle'], onChange: (v) => env.runArg('setStyle', v) } },
    { combo: { id: 'tb-font', tip: 'Font', value: env.font, width: 136, options: ['Arial', 'Calibri', 'Times New Roman', 'Courier New', 'Georgia', 'Tahoma', 'Verdana'], onChange: (v) => env.runArg('setFont', v) } },
    { combo: { id: 'tb-size', tip: 'Font Size', value: env.size, width: 48, options: ['8', '9', '10', '11', '12', '14', '16', '18', '20', '24', '28', '36', '48', '72'], onChange: (v) => env.runArg('setSize', v) } },
    btn('bold'), btn('italic'), btn('underline'), sep,
    btn('alignLeft'), btn('alignCenter'), btn('alignRight'), btn('justify'), sep,
    dropButton(env, 'lineSpacing', lineSpacingMenu(env), 'Line Spacing', 'lineSpacing'),
    sep,
    btn('numbering'), btn('bullets'), btn('decIndent'), btn('incIndent'), sep,
    splitButton(env, 'borderBtn', borderMenu(env)),
    splitButton(env, 'highlightBtn', highlightMenu(env)),
    splitButton(env, 'fontColorBtn', fontColorMenu(env)),
  ]

  const tables: QTool[] = [
    btn('drawTable'), btn('eraser'), sep,
    dropButton(env, 'dashStyle', borderStyleMenu(env), 'Line Style', 'dashStyle'),
    dropButton(env, 'lineStyle', borderWeightMenu(env), 'Line Weight', 'lineStyle'),
    dropButton(env, 'lineColor', borderColorMenu(env), 'Border Color', 'lineColor'),
    splitButton(env, 'borderBtn', borderMenu(env)),
    dropButton(env, 'fillColor', shadingColorMenu(env), 'Shading Color', 'fillColor'),
    sep,
    dropButton(env, 'table', tableMenu(env), 'Insert Table', 'table'),
    btn('mergeCells'), btn('splitCells'), sep,
    dropButton(env, 'alignTop', cellAlignMenu(env), 'Align', 'alignTop'),
    btn('distRows'), btn('distCols'), sep,
    btn('tableAutoFormat'), btn('textDirection'), sep,
    btn('sortAsc'), btn('sortDesc'), btn('autoSum'),
  ]

  const reviewing: QTool[] = [
    dropButton(env, 'markupFinalMarkup', markupMenu(env), 'Display for Review', undefined, 'Final Showing Markup'),
    dropButton(env, 'show', showMenu(env), 'Show', undefined, 'Show'),
    sep,
    btn('prevChange'), btn('nextChange'), sep,
    splitButton(env, 'acceptChange', acceptChangeMenu(env)),
    splitButton(env, 'rejectChange', rejectChangeMenu(env)),
    sep,
    btn('insertComment'), btn('trackChanges'), btn('reviewingPane'),
  ]

  const outlining: QTool[] = [
    btn('promoteH1'), btn('promote'),
    { combo: { id: 'tb-olvl', tip: 'Outline Level', value: env.outlineLevel ? t('Level {0}', env.outlineLevel) : t('Body text'), width: 74, options: [1, 2, 3, 4, 5, 6, 7, 8, 9].map((n) => t('Level {0}', n)).concat(t('Body text')), onChange: (v) => env.runArg('setOutlineLevel', v.replace(/\D/g, '')) } },
    btn('demote'), btn('demoteBody'), sep,
    btn('moveUp'), btn('moveDown'), btn('expand'), btn('collapse'),
    { combo: { id: 'tb-olshow', tip: 'Show Level', value: env.outlineShow ? t('Show Level {0}', env.outlineShow) : t('Show All Levels'), width: 96, options: [t('Show All Levels')].concat([1, 2, 3, 4, 5, 6, 7, 8, 9].map((n) => t('Show Level {0}', n))), onChange: (v) => env.runArg('setOutlineShow', v.replace(/\D/g, '')) } },
    btn('olFirstLine'), btn('olFormatting'), sep,
    btn('updateTOC'), btn('gotoTOC'),
  ]

  const headerFooter: QTool[] = [
    dropButton(env, 'hfAutoText', autoTextMenu(env), 'Insert AutoText', undefined, 'Insert AutoText'),
    sep,
    btn('hfPageNum'), btn('hfNumPages'), btn('hfFormatPageNum'), btn('hfDate'), btn('hfTime'), sep,
    btn('pageSetup'), btn('hfShowText'), btn('hfLinkPrev'), sep,
    btn('hfSwitch'), btn('hfPrev'), btn('hfNext'), sep,
    { id: 'hfClose', label: 'Close', tip: 'Close', run: () => env.run('hfClose') },
  ]

  const picture: QTool[] = [
    btn('insertPicture'),
    dropButton(env, 'colorPic', pictureColorMenu(env), 'Color', 'colorPic'),
    btn('picMoreContrast'), btn('picLessContrast'), btn('picMoreBright'), btn('picLessBright'), sep,
    btn('picCrop'), btn('picRotate'),
    dropButton(env, 'lineStyle', objectWeightMenu(env), 'Line Style', 'lineStyle'),
    sep,
    dropButton(env, 'textWrap', wrapMenu(env), 'Text Wrapping', 'textWrap'),
    btn('formatObject'), btn('picReset'),
  ]

  const mailMerge: QTool[] = [
    btn('mmMainDoc'), btn('mmOpenSource'), btn('mmRecipients'), sep,
    btn('mmAddress'), btn('mmGreeting'), btn('mmInsertField'), sep,
    btn('mmViewData'), btn('mmFirst'), btn('mmPrev'),
    { combo: { id: 'tb-mmrec', tip: 'Go to Record', value: String(env.mergeRecord), width: 34, options: Array.from({ length: env.mergeRecords }, (_, i) => String(i + 1)), onChange: (v) => env.runArg('goToRecord', v) } },
    btn('mmNext'), btn('mmLast'), sep,
    btn('mmMergeNew'), btn('mmMergePrint'),
  ]

  const wordart: QTool[] = [
    btn('insertWordArt'),
    { id: 'wordartEdit', icon: 'wordart', label: 'Edit Text...', tip: 'Edit Text...', run: () => env.run('wordartEdit') },
    btn('wordartGallery'), btn('formatObject'),
    dropButton(env, 'waShape', wordArtShapeMenu(env), 'LettersArt Shape', 'waShape'),
    dropButton(env, 'textWrap', wrapMenu(env), 'Text Wrapping', 'textWrap'),
  ]

  const drawing: QTool[] = [
    { id: 'draw', label: 'D&raw', tip: 'Draw', drop: drawMenu(env) },
    btn('selectObjects'), sep,
    { id: 'autoshapes', label: 'A&utoShapes', tip: 'AutoShapes', drop: autoShapesMenu(env) },
    btn('drawLine'), btn('drawArrow'), btn('drawRect'), btn('drawOval'), btn('textBox'),
    btn('insertWordArt'), btn('insertDiagram'), btn('insertClipArt'), btn('insertPicture'), sep,
    splitButton(env, 'fillColor', [{ cls: 'color-menu', panel: <QColorPanel mode="fill" onPick={(pick) => {
      if (pick.none) env.runArg('setFillColor', 'none')
      else if (pick.hex) env.runArg('setFillColor', pick.hex)
    }} /> }]),
    splitButton(env, 'lineColor', [{ cls: 'color-menu', panel: <QColorPanel mode="line" onPick={(pick) => {
      if (pick.none) env.runArg('setLineColor', 'none')
      else if (pick.hex) env.runArg('setLineColor', pick.hex)
    }} /> }]),
    splitButton(env, 'fontColorBtn', fontColorMenu(env)),
    sep,
    dropButton(env, 'lineStyle', objectWeightMenu(env), 'Line Style', 'lineStyle'),
    dropButton(env, 'dashStyle', objectDashMenu(env), 'Dash Style', 'dashStyle'),
    dropButton(env, 'arrowStyle', objectArrowMenu(env), 'Arrow Style', 'arrowStyle'),
    dropButton(env, 'shadowStyle', objectShadowMenu(env), 'Shadow Style', 'shadowStyle'),
  ]

  return { standard, formatting, tables, reviewing, outlining, headerFooter, picture, mailMerge, wordart, drawing }
}
