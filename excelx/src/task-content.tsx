import { useEffect, useState, type ReactNode } from 'react';
import { createPortal } from 'react-dom';

interface WorkbookTemplate { name: string }
interface HelpData { topics: Array<[string, string]>; keys: Array<[string, string]> }
interface ClipboardEntry { clip: unknown; text: string; time: number }
interface TaskPaneApi {
  current?: string;
  show(id: string): void;
  refresh?(reason: string): void;
  go?(position: number): void;
  stack?: string[];
  pos?: number;
  registry?: Record<string, { id: string; title: string; hidden?: boolean }>;
}
interface ExcelxPanesApi { task?: TaskPaneApi; TEMPLATES?: Record<string, WorkbookTemplate>; newFromTemplate?(id: string): void; helpQuery?: string; reactHelpData?(): HelpData; clips?: ClipboardEntry[] }
interface ExcelxTaskRuntime {
  panes?: ExcelxPanesApi;
  ui?: { exec(command: string): unknown; openMenu(items: unknown[], rect: { left: number; bottom: number }): unknown };
  icons?: { get(name: string, size?: number): string; app?(size?: number): string };
  t?(text: string, ...args: unknown[]): string;
  bus?: { on(event: string, callback: () => void): () => void };
  dlg?: { insertFunction(name?: string): void };
  fninfo?: { all(): Array<{ name: string; desc: string }>; syntax(fn: { name: string }): string };
  clip?: { set(value: unknown): void; pasteInternal(options: Record<string, unknown>): void };
  grid?: { paint(): void };
  app?: { newWorkbook(): void; openDialog?(accept?: string): void; toggleTask?(visible: boolean): void };
}

function runtime(): ExcelxTaskRuntime | undefined {
  return (window as Window & { L?: ExcelxTaskRuntime }).L;
}

function translate(L: ExcelxTaskRuntime, text: string, ...args: unknown[]): string {
  return L.t?.(text, ...args) || text;
}

/** React-owned Excelx task-pane bodies; workbook creation and clipboard operations stay in the engine. */
export function ExcelxTaskContentLayer({ host }: { host: HTMLElement }) {
  const [headMount, setHeadMount] = useState<HTMLElement | null>(null);
  const [mount, setMount] = useState<HTMLElement | null>(null);
  const [, refresh] = useState(0);
  const [helpQuery, setHelpQuery] = useState('');
  const L = runtime();
  const panes = L?.panes;
  const task = panes?.task;
  const current = task?.current;

  useEffect(() => {
    let target: HTMLElement | null = null;
    let slot: HTMLElement | null = null;
    let headTarget: HTMLElement | null = null;
    let headSlot: HTMLElement | null = null;
    let headDisplay = '';
    let hiddenChildren = new Map<HTMLElement, string>();
    let off: (() => void) | undefined;
    const hideLegacyContent = () => {
      if (!target || !slot) return;
      const active = ['getting-started', 'new', 'help', 'keys', 'clipboard'].includes(runtime()?.panes?.task?.current || '');
      if (!active) {
        hiddenChildren.forEach((display, child) => { child.style.display = display; });
        hiddenChildren.clear();
        return;
      }
      for (const child of Array.from(target.children) as HTMLElement[]) {
        if (child === slot || hiddenChildren.has(child)) continue;
        hiddenChildren.set(child, child.style.display);
        child.style.display = 'none';
      }
    };
    const onTaskRender = () => {
      if (target && slot && !slot.isConnected) target.append(slot);
      hideLegacyContent();
      refresh((value) => value + 1);
    };
    const attach = () => {
      const taskpane = host.querySelector<HTMLElement>('#taskpane');
      const header = taskpane?.querySelector<HTMLElement>(':scope > .tp-head');
      if (taskpane && header && header !== headTarget) {
        headSlot?.remove();
        if (headTarget) headTarget.style.display = headDisplay;
        headTarget = header;
        headDisplay = header.style.display;
        header.style.display = 'none';
        headSlot = host.ownerDocument.createElement('div');
        headSlot.id = 'react-excelx-task-head-root';
        headSlot.className = 'react-task-head-root vibeoffice-react-owned';
        headSlot.style.display = 'contents';
        taskpane.insertBefore(headSlot, header);
        setHeadMount(headSlot);
      }
      const pane = host.querySelector<HTMLElement>('#taskpane > .tp-body');
      if (!pane) return;
      if (pane === target) {
        if (slot && !slot.isConnected) pane.append(slot);
        hideLegacyContent();
        return;
      }
      slot?.remove();
      hiddenChildren.forEach((display, child) => { child.style.display = display; });
      hiddenChildren = new Map();
      target = pane;
      slot = host.ownerDocument.createElement('div');
      slot.id = 'react-excelx-task-body-root';
      slot.className = 'react-task-body-root vibeoffice-react-owned';
      slot.style.display = 'contents';
      pane.append(slot);
      hiddenChildren = new Map();
      hideLegacyContent();
      setMount(slot);
      if (!off) off = runtime()?.bus?.on('excelx-task-render', onTaskRender);
    };
    const observer = new MutationObserver(attach);
    observer.observe(host, { childList: true, subtree: true });
    attach();
    return () => {
      observer.disconnect();
      off?.();
      hiddenChildren.forEach((display, child) => { child.style.display = display; });
      if (headTarget) headTarget.style.display = headDisplay;
      headSlot?.remove();
      slot?.remove();
      setHeadMount(null);
      setMount(null);
    };
  }, [host]);

  useEffect(() => {
    if (current === 'help') {
      setHelpQuery(panes?.helpQuery || '');
      if (panes) panes.helpQuery = '';
    }
  }, [current, panes]);

  if (!L || !task) return null;
  const section = (title: string, ...children: ReactNode[]) => <section className="tp-sec">{title ? <div className="tp-h">{translate(L, title)}</div> : null}{children}</section>;
  const link = (icon: string, label: string, run: () => void) => <button className="tp-link" type="button" onClick={run}>
    {icon && <span aria-hidden="true" dangerouslySetInnerHTML={{ __html: L.icons?.get(icon) || '' }} />}<span>{translate(L, label)}</span>
  </button>;
  const currentEntry = task.registry?.[current || ''];
  const title = currentEntry ? translate(L, currentEntry.title) : '';
  const titleMenu = (button: HTMLButtonElement) => {
    const rect = button.getBoundingClientRect();
    const items = Object.values(task.registry || {}).filter((pane) => !pane.hidden).map((pane) => ({ label: pane.title, run: () => task.show(pane.id), checked: () => task.current === pane.id }));
    L.ui?.openMenu(items, { left: rect.left, bottom: rect.bottom });
  };
  const head = headMount ? createPortal(<div className="tp-head">
    <button className="tp-nav" type="button" aria-label={translate(L, 'Back')} data-tip={translate(L, 'Back')} disabled={!task.pos || task.pos <= 0} onClick={() => task.pos != null && task.go?.(task.pos - 1)} dangerouslySetInnerHTML={{ __html: L.icons?.get('back') || translate(L, 'Back').slice(0, 1) }} />
    <button className="tp-nav" type="button" aria-label={translate(L, 'Forward')} data-tip={translate(L, 'Forward')} disabled={task.pos == null || task.pos >= (task.stack?.length || 0) - 1} onClick={() => task.pos != null && task.go?.(task.pos + 1)} dangerouslySetInnerHTML={{ __html: L.icons?.get('forward') || translate(L, 'Forward').slice(0, 1) }} />
    <button className="tp-nav" type="button" aria-label={translate(L, 'Home')} data-tip={translate(L, 'Home')} onClick={() => task.show('getting-started')} dangerouslySetInnerHTML={{ __html: L.icons?.get('home') || translate(L, 'Home').slice(0, 1) }} />
    <button className="tp-title" type="button" aria-haspopup="menu" onClick={(event) => titleMenu(event.currentTarget)}><span>{title}</span><span className="dd-arrow" /></button>
    <button className="tp-nav" type="button" aria-label={translate(L, 'Close task pane')} data-tip={translate(L, 'Close')} onClick={() => L.app?.toggleTask?.(false)} dangerouslySetInnerHTML={{ __html: L.icons?.get('close') || '×' }} />
  </div>, headMount) : null;
  const withHeader = (content: ReactNode | null) => <>{head}{content}</>;
  if (!mount || !['getting-started', 'new', 'help', 'keys', 'clipboard'].includes(current || '')) return head;
  const templates = Object.entries(panes?.TEMPLATES || {});

  if (current === 'help') {
    const data = panes?.reactHelpData?.() || { topics: [], keys: [] };
    const words = helpQuery.toLocaleLowerCase().split(/\s+/).filter(Boolean);
    const matches = (text: string) => words.every((word) => text.toLocaleLowerCase().includes(word));
    const topics = data.topics.filter(([title, body]) => !words.length || matches(`${translate(L, title)} ${translate(L, body)}`));
    const functions = words.length ? (L.fninfo?.all() || []).filter((fn) => matches(`${fn.name} ${translate(L, fn.desc)}`)).slice(0, 12) : [];
    const keys = words.length ? data.keys.filter(([key, body]) => matches(`${translate(L, key)} ${translate(L, body)}`)) : [];
    return withHeader(createPortal(<>
      <div className="tp-row"><input id="hp-q" type="text" placeholder={translate(L, 'Search')} value={helpQuery} style={{ flex: 1, minWidth: 0 }} onChange={(event) => setHelpQuery(event.currentTarget.value)} onKeyDown={(event) => event.stopPropagation()} /></div>
      {topics.map(([title, body]) => <article className="help-topic" key={title}><h4>{translate(L, title)}</h4><p>{translate(L, body)}</p></article>)}
      {functions.length > 0 && <article className="help-topic"><h4>{translate(L, 'Functions')}</h4>{functions.map((fn) => <button className="help-fn" type="button" key={fn.name} onClick={() => L.dlg?.insertFunction(fn.name)}><b translate="no">{L.fninfo?.syntax(fn)}</b><span>{translate(L, fn.desc)}</span></button>)}</article>}
      {keys.length > 0 && <article className="help-topic"><h4>{translate(L, 'Keyboard shortcuts')}</h4><table className="kbd-table"><tbody>{keys.map(([key, body]) => <tr key={key}><td>{translate(L, key)}</td><td>{translate(L, body)}</td></tr>)}</tbody></table></article>}
      {!topics.length && !functions.length && !keys.length && <div className="tp-note">{translate(L, 'No help topics match your search.')}</div>}
    </>, mount));
  }

  if (current === 'keys') {
    const keys = panes?.reactHelpData?.().keys || [];
    return withHeader(createPortal(<table className="kbd-table"><tbody>{keys.map(([key, description]) => <tr key={key}><td>{translate(L, key)}</td><td>{translate(L, description)}</td></tr>)}</tbody></table>, mount));
  }

  if (current === 'clipboard') {
    const items = panes?.clips || [];
    const paste = (item: ClipboardEntry) => { L.clip?.set(item.clip); L.clip?.pasteInternal({}); L.grid?.paint(); };
    return withHeader(createPortal(<>
      <div className="tp-btns">
        <button className="btn small" type="button" onClick={() => { for (const item of items.slice().reverse()) paste(item); }}>{translate(L, 'Paste All')}</button>
        <button className="btn small" type="button" onClick={() => { if (panes?.clips) panes.clips.length = 0; task.refresh?.('clip'); }}>{translate(L, 'Clear All')}</button>
      </div>
      <div className="tp-note">{translate(L, '{0} of 24 - Clipboard. Click an item to paste.', items.length)}</div>
      {items.map((item, index) => <button className="cb-item" type="button" key={`${item.time}:${index}`} onClick={() => paste(item)}>
        <span aria-hidden="true" dangerouslySetInnerHTML={{ __html: L.icons?.get('sheetIcon') || '' }} /><span className="cb-t" translate="no">{item.text || '[cells]'}</span>
      </button>)}
    </>, mount));
  }

  if (!mount) return head;
  if (current === 'new') return withHeader(createPortal(<>
    {section('New', link('new', 'Blank workbook', () => L.app?.newWorkbook()), link('open', 'From existing workbook...', () => L.app?.openDialog?.('.xlsx,.xlsm,.xltx,.xltm')))}
    {section('Templates', ...templates.map(([id, template]) => link('sheetIcon', template.name, () => panes?.newFromTemplate?.(id))))}
    <div className="tp-note">{translate(L, 'Templates open as new, unsaved workbooks.')}</div>
  </>, mount));

  const editionName = (window as Window & { VO?: { editionName?(app: string): string } }).VO?.editionName?.('Excelx') || 'Excelx';
  return withHeader(createPortal(<>
    <div className="tp-row"><span aria-hidden="true" dangerouslySetInnerHTML={{ __html: L.icons?.app?.(32) || '' }} /><b translate="no" style={{ color: 'var(--pane-ink)' }}>{editionName}</b></div>
    {section('Search for:', <div className="tp-row">
      <input type="text" placeholder={translate(L, 'Search help')} style={{ flex: 1, minWidth: 0 }} onKeyDown={(event) => {
        if (event.key === 'Enter') { if (panes) panes.helpQuery = event.currentTarget.value; task.show('help'); }
      }} />
      <button className="btn small" type="button" onClick={(event) => {
        const query = event.currentTarget.parentElement?.querySelector('input')?.value || '';
        if (panes) panes.helpQuery = query;
        task.show('help');
      }}>{translate(L, 'Go')}</button>
    </div>)}
    {section('Open', link('open', 'More...', () => L.ui?.exec('open')), link('new', 'Create a new workbook...', () => task.show('new')))}
    {section('Learn', link('help', 'What Excelx can do', () => { if (panes) panes.helpQuery = ''; task.show('help'); }),
      link('fx', 'Functions reference', () => L.dlg?.insertFunction()), link('help', 'Keyboard shortcuts', () => task.show('keys')))}
    {section('Files', <div className="tp-note">{translate(L, 'Opens .xlsx, .xlsm, .xltx, .csv, .txt and XML Spreadsheet 2003 files; drop a file anywhere on the window to open it. Saves .xlsx, .xlsm, .xltx, .csv, .txt, .htm, .xml and PDF.')}</div>)}
  </>, mount));
}
