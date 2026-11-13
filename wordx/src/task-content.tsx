import { useEffect, useLayoutEffect, useRef, useState, type CSSProperties, type KeyboardEvent } from 'react';
import { createPortal } from 'react-dom';

type HelpEntry = [string, string];
interface ClipArtItem { id: string; name: string; svg: string }
interface ClipboardItem { index: number; text: string; table: boolean }
interface WordxRuntime {
  spell?: { reactSpellCatalog?(): { kind: string; label: string; before: string; word: string; after: string; suggestions: string[]; language: string; grammar: boolean; showExplain: boolean; canChangeAll: boolean; canAuto: boolean } | null; reactSpellChoose?(value: string): void; reactSpellAction?(action: string, value?: unknown): void };
  preserve?: { reactControlCatalog?(): { name: string; type: string; value: string; choices: Array<{ value: string; text: string }> } | null; reactControlSetValue?(value: string): void };
  mailmerge?: {
    reactRecipientsCatalog?(): { name: string; fields: string[]; rows: Array<{ index: number; values: string[]; included: boolean }> } | null;
    reactRecipientsAction?(action: 'all' | 'row', index: number, included: boolean): void;
    reactAddressListCatalog?(): { fields: string[]; rows: string[][] };
    reactAddressListSetRows?(rows: string[][]): void;
    reactFieldCatalog?(): string[];
    reactSelectField?(field: string): void;
    reactInsertField?(field: string): void;
  };
  clipart?: { search(query: string): ClipArtItem[]; url(item: ClipArtItem): string };
  panes?: {
    task?: { current?: string };
    reactHelpCatalog?(): { topics: HelpEntry[]; keys: HelpEntry[]; query: string; topic: string | null };
    reactSetHelpState?(query: string, topic: string | null): void;
    reactSearchCatalog?(query: string): Array<{ before: string; match: string; after: string }>;
    reactOpenSearchHit?(index: number): void;
    searchQuery?: string;
    insertClip?(item: ClipArtItem): void;
    reactNewCatalog?(): Array<{ id: string; name: string }>;
    reactNewAction?(action: string): void;
    reactClipboardCatalog?(): ClipboardItem[];
    reactPasteClipboard?(index: number): void;
    reactPasteAllClipboard?(): void;
    reactClearClipboard?(): void;
    reactGettingStartedAction?(action: string): void;
    reactStylesCatalog?(mode: string): Array<{ id: string; name: string; type: string; css: string; description: string; selected: boolean; mark: string }>;
    reactStylesCurrent?(): string;
    reactStyleAction?(action: string, id?: string, point?: { x: number; y: number }): void;
    reactRevealCatalog?(): { sample: string; distinguish: boolean; marks: boolean; groups: Array<{ title: string; action: string; rows: Array<{ label: string; value: string }> }>; styleDescription?: string };
    reactRevealAction?(action: string, value?: boolean): void;
    reactProtectCatalog?(): { formatting: boolean; editing: boolean; kind: string; on: boolean };
    reactProtectSet?(field: string, value: boolean): void;
    reactProtectApply?(state: { editing: boolean; formatting: boolean; kind: string }): Promise<void>;
    reactResearchSearch?(query: string, service: string, fromHistory?: boolean): Promise<{ thesaurus: null | { base?: string; missing?: boolean; near?: string[]; senses?: Array<{ heading: string; words: string[] }>; antonyms?: string[] }; document: Array<{ before: string; match: string; after: string }>; error: string }>;
    reactOpenResearchHit?(index: number): void;
    reactResearchWordMenu?(word: string, rect: DOMRect, lookUp: (word: string) => void): void;
    researchQuery?: string;
    researchService?: string;
    researchHistory?: string[];
    researchAuto?: boolean;
    reactDocMapCatalog?(): { thumbs?: boolean; items: Array<{ id: number; level: number; text: string; title: string; hasChildren: boolean; collapsed: boolean; selected: boolean }>; pages?: Array<{ index: number; width: number; height: number; scale: number; source: HTMLElement }> };
    reactDocMapToggle?(id: number): void;
    reactDocMapSelect?(id: number): void;
    reactThumbSelect?(index: number): void;
    reactMailMergeState?(): { step: number; type: string; hasData: boolean; typed: boolean; name: string; rows: number; included: number; recipient: number; steps: string[] };
    reactMailMergeAction?(action: string, value?: string): Promise<void> | void;
    reactReadingAction?(action: string): void;
    reactPreviewAction?(action: string, value?: string, anchor?: HTMLButtonElement): void;
    reactPreviewClick?(event: MouseEvent): void;
  preview?: { mag?: boolean; reactMultiPageAction?(rows: number, columns: number): void };
    reactTaskShellState?(): { current: string; title: string; backDisabled: boolean; forwardDisabled: boolean; items: Array<{ id: string; title: string; current: boolean }> };
    reactTaskShellAction?(action: string, anchor?: HTMLButtonElement, value?: string): void;
  };
  layout?: { zoom?: number; root?: HTMLElement; scroller?: HTMLElement };
  outline?: { reactPointerDown?(event: MouseEvent): void };
  find?: { reactGoToCatalog?(): Array<{ id: string; label: string; prompt: string }>; reactGoToAction?(kind: string, value: string, dir: number): boolean; reactCloseDialog?(): void; reactFindTab?(): number; reactFindCommand?(action: string, anchor?: HTMLElement): void; reactFindFormatState?(): { find: string; replace: string }; reactFindFormatMenuCatalog?(): Array<{ id: string; label: string }>; reactFindFormatMenuAction?(id: string): void; reactFindSpecialMenuCatalog?(): Array<{ code: string; label: string }>; reactFindSpecialMenuAction?(code: string): void };
  dlg?: { reactTablePropsCatalog?(): TablePropertiesState; reactTablePropsAction?(action: string, value?: Partial<TablePropertiesState>): TablePropertiesState | void; reactBordersCatalog?(): BordersDialogState; reactBordersAction?(action: string, value?: any): BordersDialogState | void; reactParagraphCatalog?(): ParagraphDialogState; reactParagraphAction?(action: string, value?: unknown): ParagraphDialogState | void; reactCaptionNumberingCatalog?(): CaptionNumbering; reactCaptionNumberingAction?(value: Partial<CaptionNumbering>): void; reactCaptionCatalog?(): CaptionDialogState; reactCaptionAction?(action: string, value?: Partial<CaptionDialogState>): Promise<void> | void; reactPageSetupCatalog?(): PageSetupState; reactPageSetupAction?(action: string, value?: Partial<PageSetupState>): void; reactColumnsCatalog?(): ColumnsDialogState; reactColumnsAction?(action: string, value?: Partial<ColumnsDialogState> | { n: number; eq: boolean; fractions?: number[] }): void; reactLineNumbersCatalog?(): { enabled: boolean; start: number; distance: string; countBy: number; restart: string; restarts: Array<[string, string]> }; reactLineNumbersAction?(action: string, value?: Record<string, unknown>): void; reactDropCapCatalog?(): { position: string; positions: Array<[string, string]>; fonts: string[]; font: string; lines: number; distanceText: string }; reactDropCapAction?(action: string, value?: string): void; reactLanguageCatalog?(): { languages: Array<{ value: string; label: string }>; language: string; noProof: boolean; detect: boolean }; reactLanguageAction?(action: string, value?: string): void; reactBookmarkCatalog?(): { name: string; sort: string; hidden: boolean; names: string[] }; reactBookmarkAction?(action: string, value?: string): boolean | void; reactObjectCatalog?(): Array<{ value: string; label: string }>; reactObjectAction?(action: string, value?: string): void; reactEquationAction?(action: string, value?: string): void; reactBreakCatalog?(): Array<{ value: string; label: string }>; reactBreakAction?(value: string): void; reactDateTimeCatalog?(): Array<{ value: string; label: string }>; reactDateTimeState?(): { format: string; autoUpdate: boolean }; reactPageNumberAction?(position: string, alignment: string, showFirst: boolean): void; reactPageNumberFormatState?(): { fmt: string; chapter: boolean; mode: string; start: number }; reactPageNumberFormatCatalog?(): Array<{ value: string; label: string }>; reactPageNumberFormatAction?(format: string, chapter: boolean, mode: string, start: number): void; reactCellDialogCatalog?(kind?: string): Array<[string, string]>; reactCellDialogKind?: string; reactInsertTableState?(): { columns: number; rows: number }; reactInsertTableAction?(action: string, options?: { columns: number; rows: number; fit: string; width: string; remember: boolean }): void; reactSplitCellsState?(): { columns: number; rows: number; merge: boolean }; reactTextToTableState?(): { initialSeparator: string; columns: number; paragraphs: Array<{ text: string; tabs: number }> }; reactTextToTableAction?(options: { sep: string; cols: number; autofit: string }): void; reactFormulaCatalog?(): { guess: string; functions: string[]; bookmarks: string[]; formats: string[] }; reactFormulaAction?(formula: string, format: string): void; reactCrossRefCatalog?(type?: string): { types: Array<{ value: string; label: string }>; references: Array<{ value: string; label: string }>; items: Array<{ value: string; label: string }> }; reactCrossRefAction?(type: string, index: number, reference: string, hyperlink: boolean, aboveBelow: boolean): void; reactMarkEntryState?(): { main: string; sub: string; option: string; xref: string; bold: boolean; italic: boolean }; reactMarkEntryAction?(all: boolean, values: { main: string; sub: string; option: string; xref: string; bold: boolean; italic: boolean }): void };
  rulers?: { draw?(): void; reactTabState?(): { kind: string; label: string }; reactTabNext?(): { kind: string; label: string }; reactMarkers?(): Array<{ className: string; tip: string; left: number; top?: number; width?: number; kind?: string; pos?: number; cid?: string; html: string }>; reactVerticalMarkers?(): Array<{ kind: string; top: number; left: number; width: number }>; reactMarkerDown?(event: PointerEvent): void; reactVerticalMarginDown?(event: PointerEvent, which: string): void; reactRulerDown?(event: PointerEvent): void; reactRulerDoubleClick?(event: MouseEvent): void; reactVerticalDoubleClick?(): void };
  app?: { fileName?: string; view?: string; prevView?: string; toggleDocMap?(): void; toggleThumbs?(): void; setView?(view: string): void; reactMapSplitterDown?(event: PointerEvent): void; reactSplitDown?(event: PointerEvent): void; reactSplitPageClick?(index: number): void; reactResize?(): void; reactScroll?(): void; reactBeforeUnload?(event: globalThis.BeforeUnloadEvent): void; reactAppKeydown?(event: globalThis.KeyboardEvent): void; reactModifierKey?(event: globalThis.KeyboardEvent): void; reactDropHintActive?: boolean; reactDropEvent?(kind: string, event: globalThis.DragEvent): void; painter?: { release(): boolean; start(sticky?: boolean, value?: unknown): boolean; on(): boolean }; highlighter?: { release(): boolean; start(sticky?: boolean, value?: unknown): boolean; on(): boolean; stop(): void }; reactStatusState?(): Array<{ id: string; text: string; className: string; tip: string; html: string }>; reactStatusAction?(action: string): void; reactShapeMenuCatalog?(): Array<{ kind: string; label: string; line: boolean; svg: string }>; reactShapeMenuAction?(kind: string): void; reactShadowMenuCatalog?(): Array<{ index: number; label: string; svg: string }>; reactShadowMenuAction?(index: number): void };
  ed?: { reactEvent?(type: string, event: Event): void };
  drawtool?: { start(options: { geom: string; line: boolean }): void; reactPointerDown?(event: PointerEvent): void; reactKeydown?(event: globalThis.KeyboardEvent): void };
  drawing?: { reactHandleDown?(event: PointerEvent): void };
  review?: { togglePane(on: boolean): void; reactPaneCatalog?(): Array<{ index: number; kind: string; head: string; body: string; color: string }>; reactPaneSelect?(index: number, edit?: boolean): void; reactReviewersCatalog?(): Array<{ author: string; color: string }>; reactCommentCatalog?(): { text: string } | null; reactCommentSetText?(text: string): void; reactBalloonDown?(id: string): void };
  icons?: { get(name: string): string; app?(size: number): string };
  ui?: { keyText(key: string): string; colorMenu?(anchor: HTMLElement, options: { mode: string }, onPick: (value: unknown) => void): void; moreColors?(initial: string | null, onPick: (color: string | null) => void): void };
  color?: { STANDARD: string[] };
  t?(text: string, ...args: unknown[]): string;
  bus?: { on(event: string, listener: () => void): () => void };
}
function runtime(): WordxRuntime | undefined { return (window as Window & { L?: WordxRuntime }).L; }
function tr(L: WordxRuntime, text: string): string { return L.t?.(text) || text; }

function OutlineInteractionBridge({ L }: { L: WordxRuntime }) {
  useEffect(() => {
    const pointerDown = (event: PointerEvent) => {
      const target = event.target as Element | null;
      if (target?.closest('#split2-bar')) L.app?.reactSplitDown?.(event);
      const balloon = target?.closest<HTMLElement>('.balloon[data-cid]');
      if (balloon) L.review?.reactBalloonDown?.(balloon.dataset.cid || '');
      if (target?.closest('.obj-handles .oh')) L.drawing?.reactHandleDown?.(event);
    };
    const click = (event: MouseEvent) => { const page = (event.target as Element | null)?.closest<HTMLElement>('#split2 .pg'); if (page) L.app?.reactSplitPageClick?.(Number(page.dataset.pg)); };
    document.addEventListener('pointerdown', pointerDown);
    document.addEventListener('click', click);
    return () => { document.removeEventListener('pointerdown', pointerDown); document.removeEventListener('click', click); };
  }, [L]);
  useEffect(() => {
    const root = L.layout?.root;
    const pointerDown = L.outline?.reactPointerDown;
    if (!root || !pointerDown) return;
    root.addEventListener('mousedown', pointerDown);
    return () => root.removeEventListener('mousedown', pointerDown);
  }, [L]);
  useEffect(() => {
    const types = ['beforeinput', 'keydown', 'compositionstart', 'compositionend', 'paste', 'copy', 'cut', 'drop', 'dragstart', 'dragover', 'mousedown', 'click', 'dblclick', 'contextmenu', 'focus', 'blur'];
    const forward = (event: Event) => L.ed?.reactEvent?.(event.type, event);
    let root: HTMLElement | null = null;
    const attach = () => {
      const next = L.layout?.root || null;
      if (next === root) return;
      if (root) types.forEach((type) => root?.removeEventListener(type, forward));
      root = next;
      if (root) types.forEach((type) => root?.addEventListener(type, forward));
    };
    const off = L.bus?.on('app-ready', attach);
    attach();
    return () => { off?.(); if (root) types.forEach((type) => root?.removeEventListener(type, forward)); };
  }, [L]);
  useEffect(() => {
    const types = ['dragenter', 'dragleave', 'dragover', 'drop'] as const;
    const handlers: Record<(typeof types)[number], (event: globalThis.DragEvent) => void> = {
      dragenter: (event) => L.app?.reactDropEvent?.('enter', event),
      dragleave: (event) => L.app?.reactDropEvent?.('leave', event),
      dragover: (event) => L.app?.reactDropEvent?.('over', event),
      drop: (event) => L.app?.reactDropEvent?.('drop', event),
    };
    const drawPointerDown = (event: PointerEvent) => {
      if (event.target instanceof Element && event.target.closest('#scroller')) L.drawtool?.reactPointerDown?.(event);
    };
    let app: HTMLElement | null = null;
    let observer: MutationObserver | null = null;
    const attach = () => {
      const next = document.querySelector<HTMLElement>('#app');
      if (next === app) { if (next) observer?.disconnect(); return; }
      if (app) types.forEach((type) => app?.removeEventListener(type, handlers[type]));
      app = next;
      if (app) {
        types.forEach((type) => app?.addEventListener(type, handlers[type]));
        app.addEventListener('pointerdown', drawPointerDown, true);
      }
      if (app) observer?.disconnect();
    };
    observer = new MutationObserver(attach);
    observer.observe(document.documentElement, { childList: true, subtree: true });
    attach();
    return () => { observer.disconnect(); if (app) { types.forEach((type) => app?.removeEventListener(type, handlers[type])); app.removeEventListener('pointerdown', drawPointerDown, true); } };
  }, [L]);
  useEffect(() => {
    const pointerDown = (event: PointerEvent) => L.app?.reactMapSplitterDown?.(event);
    let splitter: HTMLElement | null = null;
    let observer: MutationObserver | null = null;
    const attach = () => {
      const next = document.querySelector<HTMLElement>('#split-map');
      if (next === splitter) { if (next) observer?.disconnect(); return; }
      splitter?.removeEventListener('pointerdown', pointerDown);
      splitter = next;
      splitter?.addEventListener('pointerdown', pointerDown);
      if (splitter) observer?.disconnect();
    };
    observer = new MutationObserver(attach);
    observer.observe(document.documentElement, { childList: true, subtree: true });
    attach();
    return () => { observer.disconnect(); splitter?.removeEventListener('pointerdown', pointerDown); };
  }, [L]);
  useEffect(() => {
    const resize = () => L.app?.reactResize?.();
    window.addEventListener('resize', resize);
    return () => window.removeEventListener('resize', resize);
  }, [L]);
  useEffect(() => {
    const beforeUnload = (event: BeforeUnloadEvent) => L.app?.reactBeforeUnload?.(event);
    window.addEventListener('beforeunload', beforeUnload);
    return () => window.removeEventListener('beforeunload', beforeUnload);
  }, [L]);
  useEffect(() => {
    const scroll = () => L.app?.reactScroll?.();
    let scroller: HTMLElement | null = null;
    let observer: MutationObserver | null = null;
    const attach = () => {
      const next = document.querySelector<HTMLElement>('#scroller');
      if (next === scroller) { if (next) observer?.disconnect(); return; }
      scroller?.removeEventListener('scroll', scroll);
      scroller = next;
      scroller?.addEventListener('scroll', scroll);
      if (scroller) observer?.disconnect();
    };
    observer = new MutationObserver(attach);
    observer.observe(document.documentElement, { childList: true, subtree: true });
    attach();
    return () => { observer.disconnect(); scroller?.removeEventListener('scroll', scroll); };
  }, [L]);
  useEffect(() => {
    const keydown = (event: globalThis.KeyboardEvent) => { L.app?.reactModifierKey?.(event); L.app?.reactAppKeydown?.(event); L.drawtool?.reactKeydown?.(event); };
    const keyup = (event: globalThis.KeyboardEvent) => L.app?.reactModifierKey?.(event);
    let root: HTMLElement | null = null;
    let observer: MutationObserver | null = null;
    const attach = () => {
      const next = (window as Window & { VO?: { reactDirectHost?: { root?: HTMLElement | null } } }).VO?.reactDirectHost?.root || null;
      if (next === root) { if (next) observer?.disconnect(); return; }
      root?.removeEventListener('keydown', keydown);
      root?.removeEventListener('keyup', keyup);
      root = next;
      root?.addEventListener('keydown', keydown);
      root?.addEventListener('keyup', keyup);
      if (root) observer?.disconnect();
    };
    observer = new MutationObserver(attach);
    observer.observe(document.documentElement, { childList: true, subtree: true });
    attach();
    return () => { observer.disconnect(); root?.removeEventListener('keydown', keydown); root?.removeEventListener('keyup', keyup); };
  }, [L]);
  return null;
}

function HelpPane({ L }: { L: WordxRuntime }) {
  const catalog = L.panes?.reactHelpCatalog?.();
  const [query, setQuery] = useState(catalog?.query || '');
  const words = query.toLowerCase().split(/\s+/).filter(Boolean);
  const hit = (text: string) => words.every((word) => text.toLowerCase().includes(word));
  const topics = (catalog?.topics || []).filter(([title, detail]) => catalog?.topic !== 'keys' || words.length > 0).filter(([title, detail]) => !words.length || hit(`${title} ${detail} ${tr(L, title)} ${tr(L, detail)}`));
  const keys = (catalog?.keys || []).filter(([shortcut, description]) => !words.length || hit(`${shortcut} ${description} ${tr(L, description)}`));
  const keyNames = (shortcut: string) => shortcut.replace(/\b(Ctrl|Shift|Alt|Enter|Tab|Space)\b/g, (key) => L.ui?.keyText(key) || key);
  return <div className="tp-body">
    <div className="tp-row"><input aria-label={tr(L, 'Search')} type="text" value={query} placeholder={tr(L, 'Search')} style={{ flex: 1, minWidth: 0 }} onChange={(event) => { const value=event.currentTarget.value; setQuery(value); L.panes?.reactSetHelpState?.(value, catalog?.topic || null); }} /></div>
    {topics.map(([title, detail]) => <div className="help-topic" key={title}><h4>{tr(L, title)}</h4><p>{tr(L, detail)}</p></div>)}
    {!!keys.length && <div className="help-topic"><h4>{tr(L, 'Keyboard shortcuts')}</h4><table className="kbd-table"><tbody>{keys.map(([shortcut, description]) => <tr key={shortcut}><td translate="no">{keyNames(shortcut)}</td><td>{tr(L, description)}</td></tr>)}</tbody></table></div>}
    {!topics.length && !keys.length && <div className="tp-note">{tr(L, 'No help topics match your search.')}</div>}
    {catalog?.topic === 'keys' && !words.length && <button className="tp-link" type="button" onClick={() => L.panes?.reactSetHelpState?.('', null)}>{tr(L, 'Show all help topics')}</button>}
  </div>;
}

function SearchPane({ L }: { L: WordxRuntime }) {
  const [query, setQuery] = useState(L.panes?.searchQuery || '');
  const [hits, setHits] = useState<Array<{ before: string; match: string; after: string }>>([]);
  const search = () => setHits(L.panes?.reactSearchCatalog?.(query) || []);
  return <div className="tp-body">
    <section className="tp-sec"><div className="tp-h">{tr(L, 'Search for')}</div>
      <div className="tp-row"><label htmlFor="react-wordx-search">{tr(L, 'Search text:')}</label></div>
      <div className="tp-row"><input id="react-wordx-search" type="text" value={query} style={{ flex: 1, minWidth: 0 }} onChange={(event) => setQuery(event.currentTarget.value)} onKeyDown={(event) => { if (event.key === 'Enter') search(); event.stopPropagation(); }} /><button className="btn small" type="button" onClick={search}>{tr(L, 'Go')}</button></div>
      <div className="tp-note">{tr(L, 'Searches the text of the open document, including headers, footers, notes and comments.')}</div>
    </section>
    <div className="search-results">
      {hits.length > 0 && <div className="tp-note">{hits.length === 1 ? L.t?.('{0} match in {1}', hits.length, L.app?.fileName || '') : L.t?.('{0} matches in {1}', hits.length, L.app?.fileName || '')}</div>}
      {hits.map((hit, index) => <div className="search-hit" key={`${index}:${hit.before}:${hit.match}`} onClick={() => L.panes?.reactOpenSearchHit?.(index)}>{hit.before}<b>{hit.match}</b>{hit.after}</div>)}
      {!!query.trim() && !hits.length && <div className="tp-note">{tr(L, 'No matches were found in this document.')}</div>}
    </div>
  </div>;
}

function ClipArtPane({ L }: { L: WordxRuntime }) {
  const [draft, setDraft] = useState('');
  const [query, setQuery] = useState('');
  const clips = L.clipart?.search(query) || [];
  return <div className="tp-body">
    <section className="tp-sec"><div className="tp-h">{tr(L, 'Search for:')}</div><div className="tp-row">
      <input id="react-wordx-clipart-search" type="text" value={draft} style={{ flex: 1, minWidth: 0 }} onChange={(event) => setDraft(event.currentTarget.value)} onKeyDown={(event) => { if (event.key === 'Enter') setQuery(draft); event.stopPropagation(); }} />
      <button className="btn small" type="button" onClick={() => setQuery(draft)}>{tr(L, 'Go')}</button>
    </div></section>
    <div className="clip-grid">{clips.map((clip) => <button className="clip-it" key={clip.id} type="button" data-tip={clip.name} aria-label={clip.name} onClick={() => L.panes?.insertClip?.(clip)}><img src={L.clipart?.url(clip)} alt="" /><span>{tr(L, clip.name)}</span></button>)}{!clips.length && <div className="tp-note">{tr(L, 'No clips match. Try "people", "chart", "nature" or "office".')}</div>}</div>
    <div className="tp-note">{tr(L, 'Click a clip to insert it at the insertion point.')}</div>
  </div>;
}

function NewDocumentPane({ L }: { L: WordxRuntime }) {
  const templates = L.panes?.reactNewCatalog?.() || [];
  const action = (key: string, label: string) => <button className="tp-link" type="button" key={key} onClick={() => L.panes?.reactNewAction?.(key)}><span>{tr(L, label)}</span></button>;
  return <div className="tp-body">
    <section className="tp-sec"><div className="tp-h">{tr(L, 'New')}</div>{action('blank', 'Blank document')}{action('web', 'Web page')}{action('email', 'E-mail message')}{action('existing', 'From existing document...')}</section>
    <section className="tp-sec"><div className="tp-h">{tr(L, 'Templates')}</div>{action('template-dialog', 'On my computer...')}{templates.map((template) => action(`template:${template.id}`, template.name))}</section>
  </div>;
}

function ClipboardPane({ L }: { L: WordxRuntime }) {
  const [revision, setRevision] = useState(0);
  useEffect(() => L.bus?.on('wordx-task-refresh', () => setRevision((value) => value + 1)), [L]);
  const items = L.panes?.reactClipboardCatalog?.() || [];
  return <div className="tp-body" data-revision={revision}>
    <div className="tp-btns">
      <button className="btn small" type="button" onClick={() => L.panes?.reactPasteAllClipboard?.()}>{tr(L, 'Paste All')}</button>
      <button className="btn small" type="button" onClick={() => L.panes?.reactClearClipboard?.()}>{tr(L, 'Clear All')}</button>
    </div>
    <div className="tp-note">{L.t?.('{0} of 24 - Clipboard. Click an item to paste.', items.length) || `${items.length} of 24 - Clipboard. Click an item to paste.`}</div>
    {items.map((item) => <button className="cb-item" key={item.index} type="button" onClick={() => L.panes?.reactPasteClipboard?.(item.index)}>
      <span aria-hidden="true" dangerouslySetInnerHTML={{ __html: L.icons?.get(item.table ? 'table' : 'font') || '' }} />
      <span className="cb-t" translate="no">{item.text}</span>
    </button>)}
  </div>;
}

function GettingStartedPane({ L }: { L: WordxRuntime }) {
  const [query, setQuery] = useState('');
  const action = (key: string, label: string, icon?: string) => <button className="tp-link" type="button" key={key} onClick={() => L.panes?.reactGettingStartedAction?.(key)}>{icon && <span aria-hidden="true" dangerouslySetInnerHTML={{ __html: L.icons?.get(icon) || '' }} />}<span>{tr(L, label)}</span></button>;
  return <div className="tp-body">
    <div className="tp-row"><span aria-hidden="true" dangerouslySetInnerHTML={{ __html: L.icons?.app?.(32) || '' }} /><b translate="no" style={{ color: 'var(--pane-ink)' }}>Wordx</b></div>
    <section className="tp-sec"><div className="tp-h">{tr(L, 'Search for:')}</div><div className="tp-row"><input aria-label={tr(L, 'Search help')} placeholder={tr(L, 'Search help')} value={query} style={{ flex: 1, minWidth: 0 }} onChange={(event) => setQuery(event.currentTarget.value)} onKeyDown={(event) => { if (event.key === 'Enter' && query.trim()) L.panes?.reactGettingStartedAction?.(`help:${query.trim()}`); event.stopPropagation(); }} /><button className="btn small" type="button" onClick={() => { if (query.trim()) L.panes?.reactGettingStartedAction?.(`help:${query.trim()}`); }}>{tr(L, 'Go')}</button></div></section>
    <section className="tp-sec"><div className="tp-h">{tr(L, 'Open')}</div>{action('open', 'More...', 'open')}{action('new', 'Create a new document...', 'new')}{action('sample', 'Open the welcome document', 'sample')}</section>
    <section className="tp-sec"><div className="tp-h">{tr(L, 'Learn')}</div>{action('help', 'Keyboard shortcuts and help', 'help')}{action('styles', 'Styles and Formatting', 'stylesPane')}{action('mailmerge', 'Mail Merge wizard', 'mailMerge')}</section>
  </div>;
}

function StylesPane({ L }: { L: WordxRuntime }) {
  const [mode, setMode] = useState('avail');
  const [, refresh] = useState(0);
  useEffect(() => {
    const offs = ['sel', 'doc-changed', 'doc-loaded'].map((event) => L.bus?.on(event, () => refresh((value) => value + 1))).filter(Boolean) as Array<() => void>;
    return () => offs.forEach((off) => off());
  }, [L]);
  const styles = L.panes?.reactStylesCatalog?.(mode) || [];
  const apply = (action: string, id?: string) => L.panes?.reactStyleAction?.(action, id);
  return <div className="tp-body">
    <section className="tp-sec"><div className="tp-h">{tr(L, 'Formatting of selected text')}</div><div className="tp-box" style={{ minHeight: 22, fontSize: 12 }} translate="no">{L.panes?.reactStylesCurrent?.() || ''}</div><div className="tp-btns"><button className="btn small" type="button" onClick={() => apply('select-all')}>{tr(L, 'Select All')}</button><button className="btn small" type="button" onClick={() => apply('new')}>{tr(L, 'New Style...')}</button></div></section>
    <section className="tp-sec"><div className="tp-h">{tr(L, 'Pick formatting to apply')}</div><div className="sf-list">
      <div className="sf-item" role="button" tabIndex={0} onClick={() => apply('clear')} onKeyDown={(event) => { if (event.key === 'Enter' || event.key === ' ') apply('clear'); }}><span className="sf-n">{tr(L, 'Clear Formatting')}</span></div>
      {styles.map((style) => {
        const preview = Object.fromEntries(style.css.split(';').filter(Boolean).map((rule) => { const split = rule.indexOf(':'); const key = rule.slice(0, split).replace(/-([a-z])/g, (_, letter: string) => letter.toUpperCase()); return [key, rule.slice(split + 1)]; })) as CSSProperties;
        return <div className={`sf-item${style.selected ? ' on' : ''}`} key={style.id} data-style={style.id} data-tip={style.description} role="button" tabIndex={0} translate="no" onClick={() => apply('apply', style.id)} onKeyDown={(event) => { if (event.key === 'Enter' || event.key === ' ') apply('apply', style.id); }} onContextMenu={(event) => { event.preventDefault(); L.panes?.reactStyleAction?.('context', style.id, { x: event.clientX, y: event.clientY }); }}><span className="sf-n" style={preview}>{style.name}</span><span className="sf-k">{style.mark}</span></div>;
      })}
    </div></section>
    <div className="tp-row"><label htmlFor="react-wordx-style-show">{tr(L, 'Show:')}</label><select id="react-wordx-style-show" value={mode} style={{ flex: 1 }} onChange={(event) => setMode(event.currentTarget.value)}><option value="avail">{tr(L, 'Available formatting')}</option><option value="inuse">{tr(L, 'Formatting in use')}</option><option value="styles">{tr(L, 'Available styles')}</option><option value="all">{tr(L, 'All styles')}</option></select></div>
  </div>;
}

function RevealFormattingPane({ L }: { L: WordxRuntime }) {
  const [, refresh] = useState(0);
  useEffect(() => {
    const offs = ['sel', 'doc-changed', 'doc-loaded'].map((event) => L.bus?.on(event, () => refresh((value) => value + 1))).filter(Boolean) as Array<() => void>;
    return () => offs.forEach((off) => off());
  }, [L]);
  const catalog = L.panes?.reactRevealCatalog?.();
  return <div className="tp-body">
    <section className="tp-sec"><div className="tp-h">{tr(L, 'Selected text')}</div><div className="tp-box" translate="no">{catalog?.sample || ''}</div></section>
    <section className="tp-sec"><div className="tp-h">{tr(L, 'Formatting of selected text')}</div><div className="rf-tree">
      {catalog?.groups.map((group) => <section key={group.action}><h5><button type="button" className="tp-link" onClick={() => L.panes?.reactRevealAction?.(group.action)}>{tr(L, group.title)}</button></h5>{group.rows.map((row) => <div className="rf-row" key={`${row.label}:${row.value}`}><span>{tr(L, row.label)}</span> <span translate="no">{row.value}</span></div>)}</section>)}
      {catalog?.styleDescription && <div className="tp-note">{catalog.styleDescription}</div>}
    </div></section>
    <label className="tp-row"><input type="checkbox" checked={!!catalog?.distinguish} onChange={(event) => L.panes?.reactRevealAction?.('toggle-source', event.currentTarget.checked)} />{tr(L, 'Distinguish style source')}</label>
    <label className="tp-row"><input type="checkbox" checked={!!catalog?.marks} onChange={() => L.panes?.reactRevealAction?.('marks')} />{tr(L, 'Show all formatting marks')}</label>
  </div>;
}

function ProtectDocumentPane({ L }: { L: WordxRuntime }) {
  const catalog = L.panes?.reactProtectCatalog?.();
  const [formatting, setFormatting] = useState(!!catalog?.formatting);
  const [editing, setEditing] = useState(!!catalog?.editing);
  const [kind, setKind] = useState(catalog?.kind || 'readOnly');
  const kinds: Array<[string, string]> = [['readOnly', 'No changes (Read only)'], ['tracked', 'Tracked changes'], ['comments', 'Comments'], ['forms', 'Filling in forms']];
  return <div className="tp-body">
    <section className="tp-sec"><div className="tp-h">{tr(L, '1. Formatting restrictions')}</div><label className="tp-row"><input type="checkbox" checked={formatting} onChange={(event) => { setFormatting(event.currentTarget.checked); L.panes?.reactProtectSet?.('formatting', event.currentTarget.checked); }} />{tr(L, 'Limit formatting to a selection of styles')}</label></section>
    <section className="tp-sec"><div className="tp-h">{tr(L, '2. Editing restrictions')}</div><label className="tp-row"><input type="checkbox" checked={editing} onChange={(event) => setEditing(event.currentTarget.checked)} />{tr(L, 'Allow only this type of editing in the document:')}</label><select aria-label={tr(L, 'Editing restriction')} value={kind} disabled={!editing} style={{ width: '100%' }} onChange={(event) => setKind(event.currentTarget.value)}>{kinds.map(([value, label]) => <option value={value} key={value}>{tr(L, label)}</option>)}</select></section>
    <section className="tp-sec"><div className="tp-h">{tr(L, '3. Start enforcement')}</div><div className="tp-note">{tr(L, catalog?.on ? 'This document is protected from unintentional editing.' : 'Are you ready to apply these settings? (You can turn them off later)')}</div><button className="btn" type="button" style={{ width: '100%' }} onClick={() => void L.panes?.reactProtectApply?.({ editing, formatting, kind })}>{tr(L, catalog?.on ? 'Stop Protection' : 'Yes, Start Enforcing Protection')}</button></section>
  </div>;
}

function ResearchPane({ L }: { L: WordxRuntime }) {
  const [query, setQuery] = useState(L.panes?.researchQuery || '');
  const [service, setService] = useState(L.panes?.researchService || 'thesaurus');
  const [result, setResult] = useState<{ thesaurus: null | { base?: string; missing?: boolean; near?: string[]; senses?: Array<{ heading: string; words: string[] }>; antonyms?: string[] }; document: Array<{ before: string; match: string; after: string }>; error: string } | null>(null);
  const [loading, setLoading] = useState(false);
  const [revision, refresh] = useState(0);
  const search = async (value = query, fromHistory = false) => {
    setQuery(value);
    if (!value.trim()) return;
    setLoading(true);
    const data = await L.panes?.reactResearchSearch?.(value, service, fromHistory);
    setResult(data || null); setLoading(false); refresh((v) => v + 1);
  };
  useEffect(() => { if (query && L.panes && L.panes.researchAuto !== false) { L.panes.researchAuto = false; void search(query); } }, []);
  const history = L.panes?.researchHistory || [];
  const lookup = (word: string) => { setQuery(word); void search(word); };
  const previous = () => { if (history.length < 2) return; history.pop(); void search(history[history.length - 1], true); };
  const words = (items: string[]) => items.map((word) => <button className="tp-link th-word" type="button" key={word} onClick={(event) => L.panes?.reactResearchWordMenu?.(word, event.currentTarget.getBoundingClientRect(), lookup)} onDoubleClick={() => lookup(word)}>{word}<span className="dd-arrow" aria-hidden="true" /></button>);
  return <div className="tp-body" data-revision={revision}>
    <section className="tp-sec"><div className="tp-h">{tr(L, 'Search for:')}</div><div className="tp-row"><input aria-label={tr(L, 'Search for')} value={query} style={{ flex: 1, minWidth: 0 }} onChange={(event) => setQuery(event.currentTarget.value)} onKeyDown={(event) => { if (event.key === 'Enter') void search(); event.stopPropagation(); }} /><button className="btn small" type="button" onClick={() => void search()}>{tr(L, 'Go')}</button></div><div className="tp-row"><select aria-label={tr(L, 'Reference source')} value={service} style={{ flex: 1, minWidth: 0 }} onChange={(event) => setService(event.currentTarget.value)}><option value="thesaurus">{tr(L, 'Thesaurus: English (U.S.)')}</option><option value="thesaurusUK">{tr(L, 'Thesaurus: English (U.K.)')}</option><option value="document">{tr(L, 'Search This Document')}</option><option value="all">{tr(L, 'All Reference Books')}</option></select></div><div className="tp-row"><button className="btn small" type="button" disabled={history.length < 2} onClick={previous}>◄ {tr(L, 'Back')}</button></div></section>
    {loading && <div className="tp-note">{tr(L, 'Searching...')}</div>}
    {result?.error && <div className="tp-note">{result.error}</div>}
    {result?.thesaurus && <section className="help-topic"><h4>{tr(L, 'Thesaurus: English')}</h4>{result.thesaurus.base && <p className="tp-note">{L.t?.('Showing results for "{0}".', result.thesaurus.base)}</p>}{result.thesaurus.missing && <><p className="tp-note">{L.t?.('No results were found for "{0}". Alphabetical list:', query)}</p>{words(result.thesaurus.near || [])}</>}{result.thesaurus.senses?.map((sense) => <div key={sense.heading}><div className="th-sense"><b>{sense.heading}</b></div>{words(sense.words)}</div>)}{!!result.thesaurus.antonyms?.length && <div><div className="th-sense"><b>{L.t?.('{0} (Antonym)', result.thesaurus.antonyms[0])}</b></div>{words(result.thesaurus.antonyms)}</div>}</section>}
    {result && (service === 'document' || service === 'all') && <section className="help-topic"><h4>{tr(L, 'In this document')}</h4>{result.document.map((hit, index) => <button className="search-hit" type="button" key={`${index}:${hit.before}:${hit.match}`} onClick={() => L.panes?.reactOpenResearchHit?.(index)}>{hit.before}<b>{hit.match}</b>{hit.after}</button>)}{!result.document.length && <p className="tp-note">{tr(L, 'The text was not found in this document.')}</p>}{result.document.length >= 50 && <p className="tp-note">{tr(L, 'Showing the first 50 matches.')}</p>}</section>}
  </div>;
}

function DocumentMap({ L }: { L: WordxRuntime }) {
  const [, refresh] = useState(0);
  const [activePage, setActivePage] = useState(-1);
  useEffect(() => {
    const events = ['wordx-docmap-render', 'sel', 'doc-changed', 'doc-loaded'];
    const offs = events.map((event) => L.bus?.on(event, () => refresh((value) => value + 1))).filter(Boolean) as Array<() => void>;
    const target = document.querySelector('#docmap');
    const observer = target && typeof ResizeObserver !== 'undefined' ? new ResizeObserver(() => refresh((value) => value + 1)) : null;
    if (target && observer) observer.observe(target);
    return () => { offs.forEach((off) => off()); observer?.disconnect(); };
  }, [L]);
  const catalog = L.panes?.reactDocMapCatalog?.();
  if (catalog?.thumbs) {
    const pages = catalog.pages || [];
    if (!pages.length) return <div className="tp-note">{tr(L, 'Thumbnails show pages in Print Layout and Reading Layout.')}</div>;
    return <div className="thumbs">{pages.map((page) => <PageThumbnail key={page.index} page={page} active={activePage === page.index} onClick={() => { setActivePage(page.index); L.panes?.reactThumbSelect?.(page.index); }} />)}</div>;
  }
  const items = catalog?.items || [];
  if (!items.length) return <div className="dm-item" style={{ color: 'var(--ink-2)' }}>{tr(L, '(No headings)')}</div>;
  return <>{items.map((item) => <div className={`dm-item${item.selected ? ' on' : ''}`} key={item.id} data-pid={item.id} title={item.title} translate="no" style={{ paddingLeft: 4 + (item.level - 1) * 12 }} onClick={() => L.panes?.reactDocMapSelect?.(item.id)}>
    <button className="dm-tog" type="button" aria-label={tr(L, item.collapsed ? 'Expand heading' : 'Collapse heading')} onClick={(event) => { event.stopPropagation(); if (item.hasChildren) L.panes?.reactDocMapToggle?.(item.id); }}>{item.hasChildren ? (item.collapsed ? '+' : '−') : ''}</button><span>{item.text}</span>
  </div>)}</>;
}

function PageThumbnail({ page, active, onClick }: { page: { index: number; width: number; height: number; scale: number; source: HTMLElement }; active: boolean; onClick: () => void }) {
  const [target, setTarget] = useState<HTMLDivElement | null>(null);
  useLayoutEffect(() => {
    if (!target) return;
    const clone = page.source.cloneNode(true) as HTMLElement;
    clone.removeAttribute('data-pg');
    clone.querySelectorAll('[contenteditable]').forEach((element) => element.removeAttribute('contenteditable'));
    clone.style.cssText += `;transform:scale(${page.scale});transform-origin:0 0;margin:0;pointer-events:none;box-shadow:none`;
    target.replaceChildren(clone);
    return () => { target.replaceChildren(); };
  }, [page.source, page.scale, target]);
  return <><div className={`thumb${active ? ' on' : ''}`} style={{ width: page.width, height: page.height }} onClick={onClick}><div ref={setTarget} /></div><div className="thumb-n">{page.index + 1}</div></>;
}

function MailMergePane({ L }: { L: WordxRuntime }) {
  const [, refresh] = useState(0);
  useEffect(() => L.bus?.on('wordx-task-render', () => refresh((value) => value + 1)), [L]);
  const state = L.panes?.reactMailMergeState?.();
  if (!state) return null;
  const action = (key: string) => void L.panes?.reactMailMergeAction?.(key);
  const link = (key: string, label: string) => <button className="tp-link" type="button" onClick={() => action(key)}><span>{tr(L, label)}</span></button>;
  const types: Array<[string, string]> = [['letters','Letters'],['envelopes','Envelopes'],['labels','Labels'],['directory','Directory']];
  let body;
  if (state.step === 1) body = <><section className="tp-sec"><div className="tp-h">{tr(L, 'Select document type')}</div><div className="tp-note">{tr(L, 'What type of document are you working on?')}</div>{types.map(([value,label]) => <label className="tp-row" key={value}><input type="radio" name="react-mm-type" checked={state.type===value} onChange={() => L.panes?.reactMailMergeAction?.('type',value)} />{tr(L,label)}</label>)}</section><div className="tp-note">{tr(L, 'Letters: send letters to a group of people. You can personalize the letter that each person receives.')}</div></>;
  else if (state.step === 2) body = <><section className="tp-sec"><div className="tp-h">{tr(L, 'Select starting document')}</div><div className="tp-note">{tr(L, 'How do you want to set up your letters?')}</div><label className="tp-row"><input type="radio" name="react-mm-start" checked readOnly />{tr(L, 'Use the current document')}</label><label className="tp-row"><input type="radio" name="react-mm-start" onChange={() => action('template')} />{tr(L, 'Start from a template')}</label><label className="tp-row"><input type="radio" name="react-mm-start" onChange={() => action('open')} />{tr(L, 'Start from existing document')}</label></section>{state.type==='envelopes'&&link('envelopes','Envelope options...')}{state.type==='labels'&&link('labels','Label options...')}</>;
  else if (state.step === 3) body = <section className="tp-sec"><div className="tp-h">{tr(L, 'Select recipients')}</div><label className="tp-row"><input type="radio" name="react-mm-source" checked={!state.typed} readOnly />{tr(L, 'Use an existing list')}</label><label className="tp-row"><input type="radio" name="react-mm-source" checked={state.typed} readOnly onClick={() => action('type-list')} />{tr(L, 'Type a new list')}</label><div className="tp-note">{state.hasData ? L.t?.('Currently, your recipients are selected from: {0} ({1} of {2})', state.name, state.included, state.rows) : tr(L, 'Use names and addresses from a file or a database.')}</div>{link('browse','Browse...')}{state.hasData&&link('recipients','Edit recipient list...')}{link('type-list','Type a new list...')}</section>;
  else if (state.step === 4) body = <section className="tp-sec"><div className="tp-h">{tr(L, 'Write your letter')}</div><div className="tp-note">{tr(L, 'Write your letter now, then click a location in the document and choose one of the items below.')}</div>{link('address','Address block...')}{link('greeting','Greeting line...')}{link('field','More items...')}</section>;
  else if (state.step === 5) body = <section className="tp-sec"><div className="tp-h">{tr(L, 'Preview your letters')}</div><div className="tp-note">{tr(L, 'One of the merged letters is previewed here. To preview another letter, click one of the following:')}</div><div className="tp-row"><button className="btn small" type="button" onClick={() => action('recipient-back')}>&lt;&lt;</button><span>{L.t?.('Recipient: {0} of {1}', state.recipient, state.included)}</span><button className="btn small" type="button" onClick={() => action('recipient-next')}>&gt;&gt;</button></div>{link('find','Find a recipient...')}{link('recipients','Edit recipient list...')}<button className="btn small" type="button" onClick={() => action('exclude')}>{tr(L, 'Exclude this recipient')}</button></section>;
  else body = <section className="tp-sec"><div className="tp-h">{tr(L, 'Complete the merge')}</div><div className="tp-note">{tr(L, 'Wordx is ready to produce your letters.')}</div>{link('print','Print... (PDF)')}{link('new-letters','Edit individual letters...')}</section>;
  return <div className="tp-body"><div className="tp-h">{state.steps[state.step-1]}</div>{body}<div className="tp-btns"><div className="tp-note">{L.t?.('Step {0} of 6', state.step)}</div>{state.step<6&&<button className="tp-link" type="button" disabled={state.step===3&&!state.hasData} onClick={() => action('next')}><span>{L.t?.('Next: {0}', state.steps[state.step])}</span></button>}{state.step>1&&<button className="tp-link" type="button" onClick={() => action('previous')}><span>{L.t?.('Previous: {0}', state.steps[state.step-2])}</span></button>}</div></div>;
}

function ReviewingPane({ L }: { L: WordxRuntime }) {
  const [, refresh] = useState(0);
  const [selected, setSelected] = useState(-1);
  useEffect(() => {
    const offs = ['wordx-review-render', 'doc-changed', 'doc-loaded', 'sel'].map((event) => L.bus?.on(event, () => refresh((value) => value + 1))).filter(Boolean) as Array<() => void>;
    return () => offs.forEach((off) => off());
  }, [L]);
  const items = L.review?.reactPaneCatalog?.() || [];
  return <><div className="rp-head"><span>{tr(L, 'Main document changes and comments')}</span><button className="tp-nav" type="button" aria-label={tr(L, 'Close Reviewing Pane')} onClick={() => L.review?.togglePane?.(false)}>×</button></div>
    {items.length ? items.map((item) => <div className={`rp-item${selected === item.index ? ' on' : ''}`} key={item.index} onClick={() => { setSelected(item.index); L.review?.reactPaneSelect?.(item.index); }} onDoubleClick={() => { if (item.kind === 'cmt') L.review?.reactPaneSelect?.(item.index, true); }}><div className="rp-h" style={{ color: item.color }} translate="no">{item.head}</div><div className="rp-b" translate="no">{item.body}</div></div>) : <div className="rp-item tp-note">{tr(L, 'There are no tracked changes or comments in this document.')}</div>}
  </>;
}

function ReadingToolbar({ L }: { L: WordxRuntime }) {
  const [, refresh] = useState(0);
  useEffect(() => L.bus?.on('view', () => refresh((value) => value + 1)), [L]);
  if (L.app?.view !== 'reading') return null;
  const button = (icon: string, label: string, action: string, tip?: string) => <button className={`tb-btn${label ? ' with-label' : ''}`} type="button" aria-label={tr(L, tip || label)} data-tip={tr(L, tip || label)} onPointerDown={(event) => event.preventDefault()} onClick={() => L.panes?.reactReadingAction?.(action)}>{icon ? <span dangerouslySetInnerHTML={{ __html: L.icons?.get(icon) || '' }} /> : null}{label ? <span className="tb-lbl">{tr(L, label)}</span> : null}</button>;
  return <div className="readbar" role="toolbar" aria-label={tr(L, 'Reading Layout')}>
    {button('docMap', 'Document Map', 'docmap')}{button('thumbnails', 'Thumbnails', 'thumbs')}{button('find', 'Find', 'find')}{button('research', 'Research', 'research')}<span className="tb-sep" />
    {button('growFont', '', 'grow', 'Increase Text Size')}{button('shrinkFont', '', 'shrink', 'Decrease Text Size')}{button('onePage', '', 'page', 'Actual Page')}{button('multiPage', '', 'pages', 'Allow Multiple Pages')}<span className="tb-sep" />{button('', 'Close', 'close')}
  </div>;
}

function PrintPreviewToolbar({ L }: { L: WordxRuntime }) {
  const [, refresh] = useState(0);
  useEffect(() => {
    const offs = ['view', 'zoom', 'wordx-preview-view'].map((event) => L.bus?.on(event, () => refresh((value) => value + 1))).filter(Boolean) as Array<() => void>;
    return () => offs.forEach((off) => off());
  }, [L]);
  useEffect(() => {
    if (L.app?.view !== 'preview') return;
    const scroller = L.layout?.scroller;
    if (!scroller) return;
    const onClick = (event: MouseEvent) => L.panes?.reactPreviewClick?.(event);
    scroller.addEventListener('click', onClick);
    return () => scroller.removeEventListener('click', onClick);
  }, [L, L.app?.view, L.layout?.scroller]);
  if (L.app?.view !== 'preview') return null;
  const action = (name: string, value?: string, anchor?: HTMLButtonElement) => L.panes?.reactPreviewAction?.(name, value, anchor);
  const button = (icon: string, label: string, name: string, tip?: string, on?: boolean) => <button className={`tb-btn${label ? ' with-label' : ''}${on ? ' on' : ''}`} type="button" aria-label={tr(L, tip || label)} data-tip={tr(L, tip || label)} aria-pressed={on} onPointerDown={(event) => event.preventDefault()} onClick={(event) => action(name, undefined, event.currentTarget)}>{icon ? <span dangerouslySetInnerHTML={{ __html: L.icons?.get(icon) || '' }} /> : null}{label ? <span className="tb-lbl">{tr(L, label)}</span> : null}</button>;
  const zoom = Math.round((L.layout?.zoom || 1) * 100);
  return <div className="toolbar preview-bar" id="tb-preview" role="toolbar" aria-label={tr(L, 'Print Preview')}>
    <span className="tb-grip" />{button('print', '', 'print', 'Print')}{button('zoom', '', 'magnifier', 'Magnifier', !!L.panes?.preview?.mag)}{button('onePage', '', 'page', 'One Page')}{button('multiPage', '', 'multi', 'Multiple Pages')}
    <select aria-label={tr(L, 'Zoom')} value={`${zoom}%`} onChange={(event) => action('zoom', event.currentTarget.value)}>{['500%','200%','150%','100%','75%','50%','25%','10%','Page Width','Text Width','Whole Page','Two Pages'].map((item) => <option key={item}>{item}</option>)}{!['500%','200%','150%','100%','75%','50%','25%','10%'].includes(`${zoom}%`) && <option>{zoom}%</option>}</select>
    {button('ruler', '', 'ruler', 'View Ruler')}{button('shrink', '', 'shrink', 'Shrink to Fit')}{button('fullScreen', '', 'fullscreen', 'Full Screen')}{button('', 'Close Preview', 'close', 'Close Preview')}{button('help', '', 'help', 'Help')}
  </div>;
}

function TaskPaneHeader({ L }: { L: WordxRuntime }) {
  const [, refresh] = useState(0);
  const [menuOpen, setMenuOpen] = useState(false);
  const [menuPosition, setMenuPosition] = useState({ left: 0, top: 0 });
  const trigger = useRef<HTMLButtonElement>(null);
  const menu = useRef<HTMLDivElement>(null);
  useEffect(() => L.bus?.on('wordx-task-render', () => refresh((value) => value + 1)), [L]);
  useEffect(() => {
    if (!menuOpen) return;
    menu.current?.querySelector<HTMLButtonElement>('[role="menuitemradio"]')?.focus();
    const outside = (event: PointerEvent) => {
      if (!menu.current?.contains(event.target as Node) && !trigger.current?.contains(event.target as Node)) setMenuOpen(false);
    };
    document.addEventListener('pointerdown', outside);
    return () => document.removeEventListener('pointerdown', outside);
  }, [menuOpen]);
  const state = L.panes?.reactTaskShellState?.();
  if (!state) return null;
  const button = (action: string, icon: string, label: string, disabled = false) => <button className="tp-nav" type="button" aria-label={tr(L, label)} data-tip={tr(L, label)} disabled={disabled} onClick={(event) => L.panes?.reactTaskShellAction?.(action, event.currentTarget)}><span dangerouslySetInnerHTML={{ __html: L.icons?.get(icon) || '' }} /></button>;
  const openMenu = () => {
    if (!trigger.current) return;
    const { left, bottom } = trigger.current.getBoundingClientRect();
    setMenuOpen((open) => !open);
    setMenuPosition({ left, top: bottom });
  };
  const handleMenuKey = (event: KeyboardEvent<HTMLDivElement>) => {
    if (event.key === 'Escape') { event.preventDefault(); setMenuOpen(false); trigger.current?.focus(); return; }
    if (!['ArrowDown', 'ArrowUp', 'Home', 'End'].includes(event.key)) return;
    const items = [...(menu.current?.querySelectorAll<HTMLButtonElement>('[role="menuitemradio"]') || [])];
    if (!items.length) return;
    event.preventDefault();
    const current = items.indexOf(document.activeElement as HTMLButtonElement);
    const next = event.key === 'Home' ? 0 : event.key === 'End' ? items.length - 1 : (current + (event.key === 'ArrowDown' ? 1 : -1) + items.length) % items.length;
    items[next]?.focus();
  };
  return <div className="tp-head">
    {button('back', 'back', 'Back', state.backDisabled)}{button('forward', 'forward', 'Forward', state.forwardDisabled)}{button('home', 'home', 'Home')}
    <button ref={trigger} className="tp-title" type="button" aria-haspopup="menu" aria-expanded={menuOpen} aria-controls="react-wordx-task-pane-menu" aria-label={tr(L, state.title)} onClick={openMenu}><span>{tr(L, state.title)}</span><span className="dd-arrow" /></button>
    {button('close', 'close', 'Close task pane')}
    {menuOpen ? <div ref={menu} id="react-wordx-task-pane-menu" className="menu" role="menu" aria-label={tr(L, 'Task Pane')} style={{ left: menuPosition.left, top: menuPosition.top }} onKeyDown={handleMenuKey}>
      {state.items.map((item) => <button key={item.id} type="button" role="menuitemradio" aria-checked={item.current} className="m-item" style={{ background: 'none', width: '100%', textAlign: 'left' }} onClick={() => { L.panes?.reactTaskShellAction?.('show', undefined, item.id); setMenuOpen(false); }}>
        <span className="m-ico">{item.current ? <span className="m-check">✓</span> : null}</span><span className="m-lbl">{tr(L, item.title)}</span><span /><span />
      </button>)}
    </div> : null}
  </div>;
}

function WordxStatusBar({ L }: { L: WordxRuntime }) {
  const [, refresh] = useState(0);
  useEffect(() => L.bus?.on('wordx-status-render', () => refresh((value) => value + 1)), [L]);
  const cells = L.app?.reactStatusState?.() || [];
  const actions: Record<string, string> = { page: 'goto', trk: 'track', ext: 'extend', ovr: 'overtype', lang: 'language', spell: 'spell' };
  const activate = (id: string) => { const action = actions[id]; if (action) L.app?.reactStatusAction?.(action); };
  return <div className="react-wordx-statusbar" role="status" style={{ display: 'flex', flex: '1', minWidth: 0 }}>{cells.map((cell) => {
    const action = actions[cell.id];
    const handleKey = (event: KeyboardEvent<HTMLDivElement>) => { if (event.key === 'Enter' || event.key === ' ') { event.preventDefault(); activate(cell.id); } };
    const props = { className: `${cell.className} ${cell.id}`, 'data-tip': cell.tip || undefined, role: action ? 'button' : undefined, tabIndex: action ? 0 : undefined, onClick: cell.id === 'spell' ? () => activate(cell.id) : undefined, onDoubleClick: cell.id !== 'spell' && action ? () => activate(cell.id) : undefined, onKeyDown: action ? handleKey : undefined };
    return cell.html ? <div key={cell.id} {...props} dangerouslySetInnerHTML={{ __html: cell.html }} /> : <div key={cell.id} {...props}>{cell.text}</div>;
  })}</div>;
}

function GoToPane({ L }: { L: WordxRuntime }) {
  const options = L.find?.reactGoToCatalog?.() || [];
  const [kind, setKind] = useState('page');
  const [value, setValue] = useState('');
  const input = useRef<HTMLInputElement>(null);
  useEffect(() => L.bus?.on('wordx-goto-focus', () => input.current?.focus()), [L]);
  const selected = options.find((option) => option.id === kind) || options[0];
  const go = (direction: number) => L.find?.reactGoToAction?.(kind, value, direction);
  return <div className="row" style={{ alignItems: 'flex-start' }}>
    <div className="col" style={{ width: 140 }}><label htmlFor="react-wordx-goto-kind">{tr(L, 'Go to what:')}</label><select id="react-wordx-goto-kind" aria-label={tr(L, 'Go to what:')} size={Math.min(9, Math.max(4, options.length))} value={kind} onChange={(event) => setKind(event.currentTarget.value)}>{options.map((option) => <option key={option.id} value={option.id}>{option.label}</option>)}</select></div>
    <div className="col" style={{ flex: 1, minWidth: 150 }}><label htmlFor="react-wordx-goto-value">{selected?.prompt || ''}</label><input ref={input} id="react-wordx-goto-value" type="text" value={value} onChange={(event) => setValue(event.currentTarget.value)} onKeyDown={(event) => { if (event.key === 'Enter') { event.preventDefault(); event.stopPropagation(); go(1); } }} />
      <div className="tp-note">{tr(L, 'Enter + and − to move relative to the current location. Example: +4 will move forward four items.')}</div>
      <div className="tp-btns"><button className="btn small" type="button" onClick={() => go(-1)}>{tr(L, 'Previous')}</button><button className="btn small primary" type="button" onClick={() => go(1)}>{tr(L, value.trim() ? 'Go To' : 'Next')}</button><button className="btn small" type="button" onClick={() => L.find?.reactCloseDialog?.()}>{tr(L, 'Close')}</button></div>
    </div>
  </div>;
}

type FindInputAdapter = {
  reactFindState?(): { text: string; replace: string; highlight: boolean; scope: string; matchCase: boolean; wholeWord: boolean; wildcards: boolean; soundsLike: boolean; allForms: boolean; dir: string; more: boolean };
  reactFindAction?(action: string, field?: string | Record<string, unknown>): void;
  reactFindCommand?(action: string, anchor?: HTMLElement): void;
  reactFindTab?(): number;
  reactFindFormatState?(): { find: string; replace: string };
  reactFindFormatMenuCatalog?(): Array<{ id: string; label: string }>;
  reactFindFormatMenuAction?(id: string): void;
  reactFindSpecialMenuCatalog?(): Array<{ code: string; label: string }>;
  reactFindSpecialMenuAction?(code: string): void;
};

function FindReplaceInput({ L, replace }: { L: WordxRuntime; replace: boolean }) {
  const adapter = L.find as FindInputAdapter | undefined;
  const initial = adapter?.reactFindState?.() || { text: '', replace: '' };
  const id = replace ? 'fr-repl' : 'fr-find';
  const label = replace ? 'Replace with:' : 'Find what:';
  const [value, setValue] = useState(replace ? initial.replace : initial.text);
  const input = useRef<HTMLInputElement>(null);
  useEffect(() => { input.current?.focus(); return L.bus?.on('wordx-find-focus', () => input.current?.focus()); }, [L]);
  return <div className="row" style={{ alignItems: 'center', flexWrap: 'nowrap' }}>
    <label htmlFor={id} style={{ width: 80 }}>{tr(L, label)}</label>
    <input ref={input} id={id} type="text" list={replace ? 'fr-rhist' : 'fr-fhist'} value={value} style={{ flex: 1, minWidth: 0 }} onFocus={() => adapter?.reactFindAction?.('focus', replace ? 'repl' : 'find')} onChange={(event) => { setValue(event.currentTarget.value); adapter?.reactFindAction?.('input'); }} onKeyDown={(event) => { if (event.key === 'Enter') { event.preventDefault(); event.stopPropagation(); adapter?.reactFindCommand?.('find-next'); } }} />
  </div>;
}

function FindFooterActions({ L }: { L: WordxRuntime }) {
  const adapter = L.find as FindInputAdapter | undefined;
  const [tab, setTab] = useState(() => adapter?.reactFindTab?.() ?? 0);
  useEffect(() => L.bus?.on('wordx-find-tab', () => setTab(adapter?.reactFindTab?.() ?? 0)), [L, adapter]);
  const command = (action: string) => adapter?.reactFindCommand?.(action);
  const button = (label: string, action: string, primary = false) => <button key={action} type="button" className={`btn${primary ? ' primary' : ''}`} onPointerDown={(event) => event.preventDefault()} onClick={() => command(action)}>{tr(L, label)}</button>;
  if (tab === 2) return null;
  return <>{tab === 1 ? button('Replace', 'replace-one') : null}{tab === 1 ? button('Replace All', 'replace-all') : null}{button('Find Next', 'find-next', true)}{button('Cancel', 'close')}</>;
}

function FindAuxiliaryControls({ L }: { L: WordxRuntime }) {
  const adapter = L.find as FindInputAdapter | undefined;
  const initial = adapter?.reactFindState?.() || { text: '', replace: '', highlight: false, scope: 'main', matchCase: false, wholeWord: false, wildcards: false, soundsLike: false, allForms: false, dir: 'all' };
  const [highlight, setHighlight] = useState(initial.highlight);
  const [scope, setScope] = useState(initial.scope);
  return <div className="row" style={{ alignItems: 'center' }}>
    <label><input id="react-wordx-find-highlight" type="checkbox" checked={highlight} onChange={(event) => { setHighlight(event.currentTarget.checked); adapter?.reactFindAction?.('highlight', String(event.currentTarget.checked)); }} />{tr(L, 'Highlight all items found in:')}</label>
    <select id="react-wordx-find-scope" aria-label={tr(L, 'Search scope')} value={scope} onChange={(event) => { setScope(event.currentTarget.value); adapter?.reactFindAction?.('scope', event.currentTarget.value); }}>
      <option value="main">{tr(L, 'Main Document')}</option><option value="all">{tr(L, 'All stories (headers, notes, comments)')}</option>
    </select>
  </div>;
}

function MailMergeRecipientsDialogBody({ L }: { L: WordxRuntime }) {
  const catalog = L.mailmerge?.reactRecipientsCatalog?.();
  const [, refresh] = useState(0);
  const [sort, setSort] = useState({ column: -1, descending: false });
  if (!catalog) return null;
  const allIncluded = catalog.rows.length > 0 && catalog.rows.every((row) => row.included);
  const rows = catalog.rows.slice();
  if (sort.column >= 0) rows.sort((a, b) => {
    const result = (a.values[sort.column] || '').localeCompare(b.values[sort.column] || '', undefined, { numeric: true });
    return sort.descending ? -result : result;
  });
  const toggle = (action: 'all' | 'row', index: number, included: boolean) => {
    L.mailmerge?.reactRecipientsAction?.(action, index, included);
    refresh((version) => version + 1);
  };
  return <div className="datasheet" style={{ maxHeight: 280 }}>
    <table className="mm-table" data-react-owned="mail-merge-recipients" translate="no"><thead><tr>
      <th><input type="checkbox" aria-label={tr(L, 'All')} checked={allIncluded} onChange={(event) => toggle('all', -1, event.currentTarget.checked)} /></th>
      {catalog.fields.map((field, index) => <th key={`${field}-${index}`} role="button" tabIndex={0} onClick={() => setSort((previous) => ({ column: index, descending: previous.column === index ? !previous.descending : false }))} onKeyDown={(event) => { if (event.key === 'Enter' || event.key === ' ') { event.preventDefault(); setSort((previous) => ({ column: index, descending: previous.column === index ? !previous.descending : false })); } }}>{field}{sort.column === index ? sort.descending ? ' ▼' : ' ▲' : ''}</th>)}
    </tr></thead><tbody>{rows.map((row) => <tr key={row.index}>
      <td><input type="checkbox" aria-label={tr(L, 'Include')} checked={row.included} onChange={(event) => toggle('row', row.index, event.currentTarget.checked)} /></td>
      {row.values.map((value, index) => <td key={index}>{value}</td>)}
    </tr>)}</tbody></table>
  </div>;
}

function NewAddressListDialogBody({ L }: { L: WordxRuntime }) {
  const catalog = L.mailmerge?.reactAddressListCatalog?.() || { fields: [], rows: [] };
  const [rows, setRows] = useState(() => catalog.rows.map((row) => row.slice()));
  const [current, setCurrent] = useState(0);
  const firstInput = useRef<HTMLInputElement>(null);
  const updateRows = (next: string[][]) => { setRows(next); L.mailmerge?.reactAddressListSetRows?.(next); };
  const updateCell = (column: number, value: string) => updateRows(rows.map((row, index) => index === current ? row.map((cell, cellIndex) => cellIndex === column ? value : cell) : row));
  const navigate = (index: number) => setCurrent(Math.max(0, Math.min(rows.length - 1, index)));
  const add = () => { const next = [...rows, catalog.fields.map(() => '')]; updateRows(next); setCurrent(next.length - 1); window.setTimeout(() => firstInput.current?.focus(), 0); };
  const remove = () => { if (rows.length <= 1) return; const next = rows.filter((_, index) => index !== current); updateRows(next); setCurrent(Math.min(current, next.length - 1)); };
  return <div className="col">
    <div style={{ display: 'grid', gridTemplateColumns: '120px 1fr', gap: '3px 8px', maxHeight: 300, overflow: 'auto', paddingRight: 4 }} translate="no">
      {catalog.fields.map((field, index) => <div key={field} style={{ display: 'contents' }}><label htmlFor={`react-wordx-address-${index}`}>{field}</label><input ref={index === 0 ? firstInput : undefined} id={`react-wordx-address-${index}`} type="text" style={{ width: '100%' }} value={rows[current]?.[index] || ''} onChange={(event) => updateCell(index, event.currentTarget.value)} /></div>)}
    </div>
    <div className="tp-btns">
      <button type="button" className="btn small" onClick={add}>{tr(L, 'New Entry')}</button>
      <button type="button" className="btn small" disabled={rows.length <= 1} onClick={remove}>{tr(L, 'Delete Entry')}</button>
      <button type="button" className="btn small" disabled={current <= 0} aria-label={tr(L, 'Previous')} onClick={() => navigate(current - 1)}>&lt;</button>
      <button type="button" className="btn small" disabled={current >= rows.length - 1} aria-label={tr(L, 'Next')} onClick={() => navigate(current + 1)}>&gt;</button>
      <span>{L.t?.('Entry {0} of {1}', current + 1, rows.length) || `Entry ${current + 1} of ${rows.length}`}</span>
    </div>
  </div>;
}

function InsertMergeFieldDialogBody({ L }: { L: WordxRuntime }) {
  const fields = L.mailmerge?.reactFieldCatalog?.() || [];
  const [selected, setSelected] = useState(fields[0] || '');
  const select = (field: string) => { setSelected(field); L.mailmerge?.reactSelectField?.(field); };
  return <div className="col">
    <label htmlFor="react-wordx-mail-merge-field-select">{tr(L, 'Fields:')}</label>
    <select id="react-wordx-mail-merge-field-select" size={10} style={{ width: '100%' }} value={selected} onChange={(event) => select(event.currentTarget.value)} onDoubleClick={() => L.mailmerge?.reactInsertField?.(selected)}>
      {fields.map((field, index) => <option key={`${field}-${index}`} value={field}>{field}</option>)}
    </select>
  </div>;
}

function ReviewersDialogBody({ L }: { L: WordxRuntime }) {
  const reviewers = L.review?.reactReviewersCatalog?.() || [];
  return <div className="col" data-react-owned="reviewers-dialog">
    <div>{tr(L, 'Reviewers in this document:')}</div>
    {reviewers.length === 0 ? <div className="tp-note">{tr(L, 'No one has tracked changes or comments yet.')}</div> : reviewers.map(({ author, color }) => <div key={author} data-reviewer={author} style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
      <span aria-hidden="true" style={{ width: 12, height: 12, background: color, border: '1px solid #555', display: 'inline-block' }} />{author}
    </div>)}
  </div>;
}

function CommentEditDialogBody({ L }: { L: WordxRuntime }) {
  const [text, setText] = useState(() => L.review?.reactCommentCatalog?.()?.text || '');
  return <textarea autoFocus rows={6} style={{ width: '100%' }} value={text} onChange={(event) => { const next = event.currentTarget.value; setText(next); L.review?.reactCommentSetText?.(next); }} />;
}

function ContentControlDialogBody({ L }: { L: WordxRuntime }) {
  const catalog = L.preserve?.reactControlCatalog?.();
  const [value, setValue] = useState(catalog?.value || '');
  const update = (next: string) => { setValue(next); L.preserve?.reactControlSetValue?.(next); };
  if (!catalog) return null;
  return <div className="col" data-react-owned="content-control-value">
    <label htmlFor="react-wordx-content-control-value">{tr(L, 'Value:')}</label>
    {catalog.type === 'dropDownList' ? <select id="react-wordx-content-control-value" size={Math.min(8, catalog.choices.length || 1)} value={value} onChange={(event) => update(event.currentTarget.value)}>{catalog.choices.map(({ value: choice, text }, index) => <option key={`${choice}-${index}`} value={choice}>{text}</option>)}</select>
      : <input id="react-wordx-content-control-value" type={catalog.type === 'date' ? 'date' : 'text'} value={value} onChange={(event) => update(event.currentTarget.value)} />}
  </div>;
}

function SpellCheckDialogBody({ L }: { L: WordxRuntime }) {
  const [revision, refresh] = useState(0);
  useEffect(() => L.bus ? L.bus.on('wordx-spell-refresh', () => refresh((value) => value + 1)) : undefined, [L]);
  const catalog = L.spell?.reactSpellCatalog?.();
  const [choice, setChoice] = useState('');
  useEffect(() => { setChoice(catalog?.suggestions[0] || ''); }, [revision]);
  if (!catalog) return null;
  const action = (name: string) => L.spell?.reactSpellAction?.(name);
  const noSuggestions = catalog.suggestions.length === 0;
  return <div className="row" data-react-owned="spell-check" style={{ flexWrap: 'nowrap' }}>
    <div className="col" style={{ flex: 1, minWidth: 0 }}>
      <div className="cd-cap">{tr(L, catalog.label)}</div>
      <div className="sg-context" contentEditable={false} translate="no">{catalog.before}<span className={catalog.kind === 'spelling' ? 'sg-bad' : 'sg-gram'}>{catalog.word}</span>{catalog.after}</div>
      <label htmlFor="react-wordx-spell-suggestions">{tr(L, 'Suggestions:')}</label>
      <select id="react-wordx-spell-suggestions" size={6} style={{ width: '100%' }} value={choice} onChange={(event) => { const next = event.currentTarget.value; setChoice(next); L.spell?.reactSpellChoose?.(next); }} onDoubleClick={() => action('change')}>
        {noSuggestions ? <option value="" disabled>{tr(L, catalog.kind === 'spelling' ? '(No Spelling Suggestions)' : '(No Suggestions)')}</option> : catalog.suggestions.map((suggestion, index) => <option key={`${suggestion}-${index}`} value={suggestion} translate={suggestion === ' ' || suggestion === 'Delete Repeated Word' ? undefined : 'no'}>{suggestion === ' ' ? tr(L, '(single space)') : tr(L, suggestion)}</option>)}
      </select>
      <div className="field"><label>{tr(L, 'Dictionary language:')}</label><span>{tr(L, catalog.language)}</span></div>
      <label><input type="checkbox" checked={catalog.grammar} onChange={(event) => L.spell?.reactSpellAction?.('grammar', event.currentTarget.checked)} /> {tr(L, 'Check grammar')}</label>
    </div>
    <div className="col" style={{ width: 140 }}>
      <button className="btn" onClick={() => action('ignore-once')}>{tr(L, 'Ignore Once')}</button>
      <button className="btn" onClick={() => action('ignore-all')}>{tr(L, catalog.kind === 'spelling' ? 'Ignore All' : 'Ignore Rule')}</button>
      <button className="btn" onClick={() => action('add')}>{tr(L, catalog.kind === 'spelling' ? 'Add to Dictionary' : 'Next Sentence')}</button>
      <div style={{ height: 16 }} />
      <button className="btn" onClick={() => action('change')} disabled={noSuggestions && catalog.kind !== 'repeat'}>{tr(L, 'Change')}</button>
      {catalog.canChangeAll && <button className="btn" onClick={() => action('change-all')} disabled={noSuggestions}>{tr(L, 'Change All')}</button>}
      {catalog.canAuto && <button className="btn" onClick={() => action('auto-correct')}>{tr(L, 'AutoCorrect')}</button>}
      {catalog.showExplain && <button className="btn" onClick={() => action('explain')}>{tr(L, 'Explain...')}</button>}
    </div>
  </div>;
}

function FindOptionsControls({ L }: { L: WordxRuntime }) {
  const adapter = L.find as FindInputAdapter | undefined;
  const initial = adapter?.reactFindState?.() || { text: '', replace: '', highlight: false, scope: 'main', matchCase: false, wholeWord: false, wildcards: false, soundsLike: false, allForms: false, dir: 'all' };
  const [state, setState] = useState(initial);
  const update = (key: 'matchCase' | 'wholeWord' | 'wildcards' | 'soundsLike' | 'allForms' | 'dir', value: boolean | string) => {
    const next = { ...state, [key]: value };
    if (key === 'wildcards' && value) { next.soundsLike = false; next.allForms = false; }
    setState(next);
    adapter?.reactFindAction?.('options', next);
  };
  return <fieldset className="col"><legend>{tr(L, 'Search Options')}</legend>
    <label>{tr(L, 'Search:')} <select id="fr-dir" value={state.dir} onChange={(event) => update('dir', event.currentTarget.value)}><option value="all">{tr(L, 'All')}</option><option value="down">{tr(L, 'Down')}</option><option value="up">{tr(L, 'Up')}</option></select></label>
    <label><input type="checkbox" checked={state.matchCase} onChange={(event) => update('matchCase', event.currentTarget.checked)} />{tr(L, 'Match case')}</label>
    <label><input type="checkbox" checked={state.wholeWord} onChange={(event) => update('wholeWord', event.currentTarget.checked)} />{tr(L, 'Find whole words only')}</label>
    <label><input type="checkbox" checked={state.wildcards} onChange={(event) => update('wildcards', event.currentTarget.checked)} />{tr(L, 'Use wildcards')}</label>
    <label><input type="checkbox" checked={state.soundsLike} disabled={state.wildcards} onChange={(event) => update('soundsLike', event.currentTarget.checked)} />{tr(L, 'Sounds like (English)')}</label>
    <label><input type="checkbox" checked={state.allForms} disabled={state.wildcards} onChange={(event) => update('allForms', event.currentTarget.checked)} />{tr(L, 'Find all word forms (English)')}</label>
  </fieldset>;
}

function FindFormatControls({ L }: { L: WordxRuntime }) {
  const adapter = L.find as FindInputAdapter | undefined;
  return <div className="tp-btns">
    <button type="button" className="btn small" onClick={() => adapter?.reactFindCommand?.('clear-format')}>{tr(L, 'No Formatting')}</button>
    <button type="button" className="btn small" onClick={(event) => adapter?.reactFindCommand?.('format', event.currentTarget)}>{tr(L, 'Format')} ▾</button>
    <button type="button" className="btn small" onClick={(event) => adapter?.reactFindCommand?.('special', event.currentTarget)}>{tr(L, 'Special')} ▾</button>
  </div>;
}

function FindFormatSummary({ L, replace }: { L: WordxRuntime; replace: boolean }) {
  const adapter = L.find as FindInputAdapter | undefined;
  const [summary, setSummary] = useState(() => adapter?.reactFindFormatState?.()[replace ? 'replace' : 'find'] || '');
  useEffect(() => L.bus?.on('wordx-find-format-summary', () => setSummary(adapter?.reactFindFormatState?.()[replace ? 'replace' : 'find'] || '')), [L, adapter, replace]);
  return <>{summary}</>;
}

function FindMoreToggle({ L }: { L: WordxRuntime }) {
  const adapter = L.find as FindInputAdapter | undefined;
  const [expanded, setExpanded] = useState(adapter?.reactFindState?.().more || false);
  const toggle = () => { const next = !expanded; setExpanded(next); adapter?.reactFindAction?.('more', String(next)); };
  return <button type="button" className="btn small" aria-expanded={expanded} onClick={toggle}>{expanded ? '« ' + tr(L, '&Less') : tr(L, '&More') + ' »'}</button>;
}

type SortDialogState = {
  inTable: boolean;
  fields: Array<[string, string]>;
  keys: Array<{ by: string; type: string; dir: 'asc' | 'desc' }>;
  header: 'yes' | 'no';
  separator: '\t' | ',' | 'other';
  caseSensitive: boolean;
};

function SortDialogBody({ L }: { L: WordxRuntime }) {
  const adapter = L.dlg as (typeof L.dlg & {
    reactSortCatalog?(): SortDialogState;
    reactSortAction?(action: string, state?: SortDialogState): void;
  }) | undefined;
  const [state, setState] = useState(adapter?.reactSortCatalog?.() || { inTable: false, fields: [], keys: [{ by: '0', type: 'text', dir: 'asc' as const }, { by: '', type: 'text', dir: 'asc' as const }, { by: '', type: 'text', dir: 'asc' as const }], header: 'no' as const, separator: '\t' as const, caseSensitive: false });
  const update = (next: SortDialogState) => { setState(next); adapter?.reactSortAction?.('change', next); };
  const updateKey = (index: number, key: 'by' | 'type' | 'dir', value: string) => {
    const keys = state.keys.map((item, current) => current === index ? { ...item, [key]: value } : item) as SortDialogState['keys'];
    update({ ...state, keys });
  };
  const keyRow = (index: number, label: string) => {
    const key = state.keys[index] || { by: '', type: 'text', dir: 'asc' as const };
    return <fieldset key={index}><legend>{tr(L, label)}</legend><select id={`react-wordx-sort-by-${index}`} aria-label={tr(L, label)} value={key.by} onChange={(event) => updateKey(index, 'by', event.currentTarget.value)}>{(index ? [['', '(none)'] as [string, string], ...state.fields] : state.fields).map(([value, text]) => <option key={value || 'none'} value={value}>{tr(L, text)}</option>)}</select><label>{tr(L, 'Type:')} <select id={`react-wordx-sort-type-${index}`} value={key.type} onChange={(event) => updateKey(index, 'type', event.currentTarget.value)}><option value="text">{tr(L, 'Text')}</option><option value="number">{tr(L, 'Number')}</option><option value="date">{tr(L, 'Date')}</option></select></label><div role="group" aria-label={tr(L, `${label} direction`)}><label><input type="radio" name={`react-wordx-sort-dir-${index}`} checked={key.dir === 'asc'} onChange={() => updateKey(index, 'dir', 'asc')} />{tr(L, 'Ascending')}</label><label><input type="radio" name={`react-wordx-sort-dir-${index}`} checked={key.dir === 'desc'} onChange={() => updateKey(index, 'dir', 'desc')} />{tr(L, 'Descending')}</label></div></fieldset>;
  };
  return <div className="col">{keyRow(0, 'Sort by')}{keyRow(1, 'Then by')}{keyRow(2, 'Then by')}<fieldset><legend>{tr(L, 'My list has')}</legend><label><input id="react-wordx-sort-header-yes" type="radio" name="react-wordx-sort-header" checked={state.header === 'yes'} onChange={() => update({ ...state, header: 'yes' })} />{tr(L, 'Header row')}</label><label><input id="react-wordx-sort-header-no" type="radio" name="react-wordx-sort-header" checked={state.header === 'no'} onChange={() => update({ ...state, header: 'no' })} />{tr(L, 'No header row')}</label></fieldset><div className="row"><label>{tr(L, 'Separate fields at:')} <select id="react-wordx-sort-separator" value={state.separator} onChange={(event) => update({ ...state, separator: event.currentTarget.value as SortDialogState['separator'] })}><option value="\t">{tr(L, 'Tabs')}</option><option value=",">{tr(L, 'Commas')}</option><option value="other">{tr(L, 'Other')}</option></select></label><label><input id="react-wordx-sort-case" type="checkbox" checked={state.caseSensitive} onChange={(event) => update({ ...state, caseSensitive: event.currentTarget.checked })} />{tr(L, 'Case sensitive')}</label></div></div>;
}

type TableAutoFormatState = { names: string[]; selected: string; look: { firstRow: boolean; firstCol: boolean; lastRow: boolean; lastCol: boolean } };
type TableStylePreview = Array<Array<{ text: string; style: string }>>;

function RulerTabSelector({ L }: { L: WordxRuntime }) {
  const [state, setState] = useState(() => L.rulers?.reactTabState?.() || { kind: 'left', label: 'Left Tab' });
  const next = () => { const value = L.rulers?.reactTabNext?.(); if (value) setState(value); };
  const paths: Record<string, React.ReactNode> = {
    left: <path d="M2 0v7h6" />, center: <path d="M4.5 0v7M0 7h9" />, right: <path d="M7 0v7H1" />,
    decimal: <><path d="M4.5 0v7M0 7h9" /><circle cx="7.3" cy="4" r="1" /></>, bar: <path d="M4.5 0v9" />,
    firstLine: <path d="M.5 .5h8l-4 5z" />, hanging: <path d="M4.5 1l4 5h-8z" />,
  };
  return <button className="tabsel" type="button" aria-label={tr(L, state.label)} title={tr(L, state.label)} data-tip={tr(L, state.label)} onPointerDown={(event) => event.preventDefault()} onClick={next}><svg width="9" height="9" viewBox="0 0 9 9" aria-hidden="true" fill={state.kind === 'firstLine' || state.kind === 'hanging' ? '#fff' : 'none'} stroke="#000" strokeWidth="1.6">{paths[state.kind] || paths.left}</svg></button>;
}

function RulerMarkerLayer({ L }: { L: WordxRuntime }) {
  const [markers, setMarkers] = useState(() => L.rulers?.reactMarkers?.() || []);
  useEffect(() => L.bus?.on('wordx-ruler-marks', () => setMarkers(L.rulers?.reactMarkers?.() || [])), [L]);
  return <>{markers.map((marker, index) => <div key={`${marker.kind || 'mark'}-${marker.pos ?? marker.cid ?? index}`} className={marker.className} data-tip={tr(L, marker.tip)} data-kind={marker.kind} data-pos={marker.pos} data-cid={marker.cid} style={{ left: `${marker.left}px`, top: marker.top == null ? undefined : `${marker.top}px`, width: marker.width == null ? undefined : `${marker.width}px`, zIndex: 2 }} onPointerDown={(event) => { event.stopPropagation(); L.rulers?.reactMarkerDown?.(event.nativeEvent); }} onDoubleClick={(event) => { event.stopPropagation(); L.rulers?.reactRulerDoubleClick?.(event.nativeEvent); }} dangerouslySetInnerHTML={{ __html: marker.html }} />)}</>;
}

function RulerEventSurface({ L }: { L: WordxRuntime }) {
  return <div aria-hidden="true" style={{ position: 'absolute', inset: 0, zIndex: 1, background: 'transparent' }} onPointerDown={(event) => { event.stopPropagation(); L.rulers?.reactRulerDown?.(event.nativeEvent); }} onDoubleClick={(event) => { event.stopPropagation(); L.rulers?.reactRulerDoubleClick?.(event.nativeEvent); }} />;
}

function VerticalRulerEventSurface({ L }: { L: WordxRuntime }) {
  return <div aria-hidden="true" style={{ position: 'absolute', inset: 0, zIndex: 1, background: 'transparent' }} onDoubleClick={(event) => { event.stopPropagation(); L.rulers?.reactVerticalDoubleClick?.(); }} />;
}

function VerticalRulerMarkerLayer({ L }: { L: WordxRuntime }) {
  const [markers, setMarkers] = useState(() => L.rulers?.reactVerticalMarkers?.() || []);
  useEffect(() => L.bus?.on('wordx-ruler-marks', () => setMarkers(L.rulers?.reactVerticalMarkers?.() || [])), [L]);
  return <>{markers.map((marker) => <div key={marker.kind} className="react-wordx-v-ruler-marker" data-kind={marker.kind} aria-label={tr(L, marker.kind === 'mt' ? 'Top margin' : 'Bottom margin')} style={{ position: 'absolute', zIndex: 2, left: `${marker.left}px`, top: `${marker.top}px`, width: `${marker.width}px`, height: '8px', cursor: 'ns-resize', borderTop: '2px solid var(--ruler-line)', boxSizing: 'border-box' }} onPointerDown={(event) => { event.stopPropagation(); L.rulers?.reactVerticalMarginDown?.(event.nativeEvent, marker.kind); }} />)}</>;
}

type TablePropertiesState = {
  tab: number; tableWidthEnabled: boolean; tableWidth: number; tableWidthUnit: string; alignment: string; indent: number;
  wrap: string; margins: { t: number; b: number; l: number; r: number }; spacing: number; autoResize: boolean;
  rowHeightEnabled: boolean; rowHeight: number; rowHeightRule: string; allowRowBreak: boolean; repeatHeader: boolean;
  columnWidthEnabled: boolean; columnWidth: number; cellWidthEnabled: boolean; cellWidth: number; verticalAlignment: string;
  sameCellMargins: boolean; cellMargins: { t: number; b: number; l: number; r: number }; wrapText: boolean; fitText: boolean;
  rowLabel: string; columnLabel: string;
};

function TablePropertiesDialogBody({ L }: { L: WordxRuntime }) {
  const adapter = L.dlg;
  const [state, setState] = useState<TablePropertiesState>(() => adapter?.reactTablePropsCatalog?.() || {
    tab: 0, tableWidthEnabled: false, tableWidth: 0, tableWidthUnit: 'dxa', alignment: 'left', indent: 0, wrap: 'none',
    margins: { t: 0, b: 0, l: 0, r: 0 }, spacing: 0, autoResize: true, rowHeightEnabled: false, rowHeight: 14.4,
    rowHeightRule: 'atLeast', allowRowBreak: true, repeatHeader: false, columnWidthEnabled: true, columnWidth: 0,
    cellWidthEnabled: true, cellWidth: 0, verticalAlignment: 'top', sameCellMargins: true,
    cellMargins: { t: 0, b: 0, l: 0, r: 0 }, wrapText: true, fitText: false, rowLabel: 'Row:', columnLabel: 'Column:',
  });
  const update = (patch: Partial<TablePropertiesState>) => {
    const next = { ...state, ...patch };
    setState(next);
    adapter?.reactTablePropsAction?.('change', next);
  };
  const number = (id: string, value: number, change: (value: number) => void, min = 0) => <input id={id} type="number" min={min} step="any" value={value} onChange={(event) => change(Number(event.currentTarget.value) || 0)} />;
  const marginInputs = (scope: 'margins' | 'cellMargins') => <div className="row">{([['t','Top'],['l','Left'],['b','Bottom'],['r','Right']] as const).map(([key,label]) => <label key={key} data-table-margin={key}>{tr(L, label)}{number(`react-wordx-${scope}-${key}`, state[scope][key], (value) => update({ [scope]: { ...state[scope], [key]: value } }))}</label>)}</div>;
  const tabs = ['Table', 'Row', 'Column', 'Cell'];
  return <div className="col" id="react-wordx-table-properties-body">
    <div className="tabs-head" role="tablist">{tabs.map((label, index) => <button key={label} type="button" role="tab" aria-selected={state.tab === index} className={state.tab === index ? 'on' : ''} onClick={() => update({ tab: index })}>{tr(L, label)}</button>)}</div>
    {state.tab === 0 ? <div role="tabpanel" className="col">
      <fieldset><legend>{tr(L, 'Size')}</legend><label><input id="react-wordx-table-width-enabled" type="checkbox" checked={state.tableWidthEnabled} onChange={(event) => update({ tableWidthEnabled: event.currentTarget.checked })} />{tr(L, 'Preferred width:')}</label><label>{tr(L, 'Width:')}{number('react-wordx-table-width', state.tableWidth, (tableWidth) => update({ tableWidth }))}</label><label>{tr(L, 'Measure in:')}<select id="react-wordx-table-width-unit" value={state.tableWidthUnit} onChange={(event) => update({ tableWidthUnit: event.currentTarget.value })}><option value="dxa">{tr(L, 'Inches/cm')}</option><option value="pct">{tr(L, 'Percent')}</option></select></label></fieldset>
      <fieldset><legend>{tr(L, 'Alignment')}</legend><label>{tr(L, 'Alignment:')}<select id="react-wordx-table-alignment" value={state.alignment} onChange={(event) => update({ alignment: event.currentTarget.value })}><option value="left">{tr(L, 'Left')}</option><option value="center">{tr(L, 'Center')}</option><option value="right">{tr(L, 'Right')}</option></select></label><label>{tr(L, 'Indent from left:')}{number('react-wordx-table-indent', state.indent, (indent) => update({ indent }))}</label></fieldset>
      <fieldset><legend>{tr(L, 'Text wrapping')}</legend>{(['none','around'] as const).map((value) => <label key={value}><input type="radio" name="react-wordx-table-wrap" checked={state.wrap === value} onChange={() => update({ wrap: value })} />{tr(L, value === 'none' ? 'None' : 'Around')}</label>)}</fieldset>
      <fieldset><legend>{tr(L, 'Default cell margins')}</legend>{marginInputs('margins')}</fieldset>
      <fieldset><legend>{tr(L, 'Default cell spacing')}</legend><label>{tr(L, 'Allow spacing between cells')}{number('react-wordx-table-spacing', state.spacing, (spacing) => update({ spacing }))}</label></fieldset>
      <label><input type="checkbox" checked={state.autoResize} onChange={(event) => update({ autoResize: event.currentTarget.checked })} />{tr(L, 'Automatically resize to fit contents')}</label>
      <button type="button" className="btn small" onClick={() => adapter?.reactTablePropsAction?.('borders')}>{tr(L, 'Borders and Shading...')}</button>
    </div> : null}
    {state.tab === 1 ? <div role="tabpanel" className="col"><div className="cd-cap">{tr(L, state.rowLabel)}</div><fieldset><legend>{tr(L, 'Size')}</legend><label><input id="react-wordx-row-height-enabled" type="checkbox" checked={state.rowHeightEnabled} onChange={(event) => update({ rowHeightEnabled: event.currentTarget.checked })} />{tr(L, 'Specify height:')}</label><label>{tr(L, 'Height:')}{number('react-wordx-row-height', state.rowHeight, (rowHeight) => update({ rowHeight }), 1)}</label><label>{tr(L, 'Row height is:')}<select id="react-wordx-row-height-rule" value={state.rowHeightRule} onChange={(event) => update({ rowHeightRule: event.currentTarget.value })}><option value="atLeast">{tr(L, 'At least')}</option><option value="exact">{tr(L, 'Exactly')}</option></select></label></fieldset><fieldset><legend>{tr(L, 'Options')}</legend><label><input type="checkbox" checked={state.allowRowBreak} onChange={(event) => update({ allowRowBreak: event.currentTarget.checked })} />{tr(L, 'Allow row to break across pages')}</label><label><input id="react-wordx-row-header" type="checkbox" checked={state.repeatHeader} onChange={(event) => update({ repeatHeader: event.currentTarget.checked })} />{tr(L, 'Repeat as header row at the top of each page')}</label></fieldset></div> : null}
    {state.tab === 2 ? <div role="tabpanel" className="col"><div className="cd-cap">{tr(L, state.columnLabel)}</div><fieldset><legend>{tr(L, 'Size')}</legend><label><input type="checkbox" checked={state.columnWidthEnabled} onChange={(event) => update({ columnWidthEnabled: event.currentTarget.checked })} />{tr(L, 'Preferred width:')}</label><label>{tr(L, 'Width:')}{number('react-wordx-column-width', state.columnWidth, (columnWidth) => update({ columnWidth }), 6)}</label></fieldset></div> : null}
    {state.tab === 3 ? <div role="tabpanel" className="col"><fieldset><legend>{tr(L, 'Size')}</legend><label><input type="checkbox" checked={state.cellWidthEnabled} onChange={(event) => update({ cellWidthEnabled: event.currentTarget.checked })} />{tr(L, 'Preferred width:')}</label><label>{tr(L, 'Width:')}{number('react-wordx-cell-width', state.cellWidth, (cellWidth) => update({ cellWidth }), 6)}</label></fieldset><fieldset><legend>{tr(L, 'Vertical alignment')}</legend><label>{tr(L, 'Vertical alignment:')}<select id="react-wordx-cell-v-align" value={state.verticalAlignment} onChange={(event) => update({ verticalAlignment: event.currentTarget.value })}><option value="top">{tr(L, 'Top')}</option><option value="center">{tr(L, 'Center')}</option><option value="bottom">{tr(L, 'Bottom')}</option></select></label></fieldset><fieldset><legend>{tr(L, 'Cell margins')}</legend><label><input type="checkbox" checked={state.sameCellMargins} onChange={(event) => update({ sameCellMargins: event.currentTarget.checked })} />{tr(L, 'Same as the whole table')}</label>{marginInputs('cellMargins')}</fieldset><fieldset><legend>{tr(L, 'Options')}</legend><label><input type="checkbox" checked={state.wrapText} onChange={(event) => update({ wrapText: event.currentTarget.checked })} />{tr(L, 'Wrap text')}</label><label><input id="react-wordx-cell-fit-text" type="checkbox" checked={state.fitText} onChange={(event) => update({ fitText: event.currentTarget.checked })} />{tr(L, 'Fit text')}</label></fieldset></div> : null}
  </div>;
}

function TableAutoFormatDialogBody({ L }: { L: WordxRuntime }) {
  const adapter = L.dlg as (typeof L.dlg & {
    reactTableAutoFormatCatalog?(): TableAutoFormatState;
    reactTableAutoFormatPreview?(name: string, look: TableAutoFormatState['look']): TableStylePreview;
    reactTableAutoFormatAction?(action: string, state?: TableAutoFormatState): void;
  }) | undefined;
  const [state, setState] = useState(adapter?.reactTableAutoFormatCatalog?.() || { names: [], selected: '', look: { firstRow: true, firstCol: true, lastRow: false, lastCol: false } });
  const update = (next: TableAutoFormatState) => { setState(next); adapter?.reactTableAutoFormatAction?.('change', next); };
  const preview = adapter?.reactTableAutoFormatPreview?.(state.selected, state.look) || [];
  const previewCellStyle = (cssText: string) => Object.fromEntries(cssText.split(';').filter(Boolean).map((part) => {
    const separator = part.indexOf(':');
    const name = part.slice(0, separator).trim().replace(/-([a-z])/g, (_match, letter: string) => letter.toUpperCase());
    return [name, part.slice(separator + 1).trim()];
  })) as CSSProperties;
  const optionLabels: Array<[keyof TableAutoFormatState['look'], string]> = [['firstRow', 'Heading rows'], ['firstCol', 'First column'], ['lastRow', 'Last row'], ['lastCol', 'Last column']];
  return <div className="col">
    <div>{tr(L, 'Category: All table styles')}</div>
    <div className="row" style={{ alignItems: 'flex-start' }}>
      <div className="col" style={{ width: 200 }}><label htmlFor="react-wordx-table-autoformat-style">{tr(L, 'Table styles:')}</label><select id="react-wordx-table-autoformat-style" size={9} value={state.selected} onChange={(event) => update({ ...state, selected: event.currentTarget.value })}>{state.names.map((name) => <option key={name} value={name} translate="no">{name}</option>)}</select></div>
      <div className="col"><div className="cd-cap">{tr(L, 'Preview')}</div><div className="preview-box" style={{ width: 220, height: 170 }}><table style={{ borderCollapse: 'collapse', font: '10px Arial,Arimo,sans-serif', width: 200 }}><tbody>{preview.map((row, rowIndex) => <tr key={rowIndex}>{row.map((cell, cellIndex) => <td key={cellIndex} style={previewCellStyle(cell.style)}>{cell.text}</td>)}</tr>)}</tbody></table></div></div>
    </div>
    <fieldset><legend>{tr(L, 'Apply special formats to')}</legend>{optionLabels.map(([key, label]) => <label key={key}><input type="checkbox" checked={state.look[key]} onChange={(event) => update({ ...state, look: { ...state.look, [key]: event.currentTarget.checked } })} />{tr(L, label)}</label>)}</fieldset>
  </div>;
}

type TabsDialogState = { tabs: Array<{ value: string; label: string }>; cleared: string[]; position: string; alignment: string; leader: string; defaultTab: number };

type ParagraphDialogState = { jc: string; l: number; r: number; fl: number; b: number; a: number; line: number; rule: string; outline?: number; contextual: boolean; widow: boolean; keepNext: boolean; keepLines: boolean; pageBreakBefore: boolean; suppressLineNumbers: boolean; suppressAutoHyphens: boolean; sameIndents: boolean; changed: string[]; levels: Array<[string, string]> };

type BorderPanelState = { kind: string; sides: string[]; borders: Record<string, { val?: string; sz?: number; color?: string } | null>; pen: { val?: string; sz?: number; color?: string }; setting: string; shadow: boolean; styles: Array<[string, string]>; widths: number[] };
type BordersDialogState = { tab: number; inTable: boolean; target: string; showApply: boolean; applyOptions: Array<[string, string]>; pageOptions: Array<[string, string]>; border: BorderPanelState; page: BorderPanelState; pageApply: string; shadingApply: string; fill: string | null; pattern: string; patternColor: string | null; patterns: Array<[string, string]>; colors: string[] };

function BordersDialogBody({ L }: { L: WordxRuntime }) {
  const adapter = L.dlg as (typeof L.dlg & { reactBordersCatalog?(): BordersDialogState; reactBordersAction?(action: string, value?: unknown): BordersDialogState | void }) | undefined;
  const [state, setState] = useState<BordersDialogState>(adapter?.reactBordersCatalog?.() || { tab: 0, inTable: false, target: 'para', showApply: true, applyOptions: [], pageOptions: [], border: { kind: 'para', sides: [], borders: {}, pen: {}, setting: 'none', shadow: false, styles: [], widths: [] }, page: { kind: 'page', sides: [], borders: {}, pen: {}, setting: 'none', shadow: false, styles: [], widths: [] }, pageApply: 'section', shadingApply: 'para', fill: null, pattern: 'clear', patternColor: null, patterns: [], colors: [] });
  useEffect(() => L.bus?.on('wordx-borders-refresh', () => { const next = adapter?.reactBordersCatalog?.(); if (next) setState(next); }), [L, adapter]);
  const act = (action: string, value?: unknown) => { const next = adapter?.reactBordersAction?.(action, value); if (next) setState(next); };
  const renderPanel = (kind: 'border' | 'page') => {
    const panel = state[kind], allowed = kind === 'page' ? ['top','bottom','left','right'] : panel.sides;
    const radioOptions: Array<[string,string]> = [['none','None'],['box','Box'],['shadow','Shadow'],['3d','3-D'],['custom','Custom'],...(panel.kind === 'cell' || panel.kind === 'table' ? [['all','All'],['grid','Grid']] as Array<[string,string]> : [])];
    const line = (side: string) => { const item = panel.borders[side]; if (!item) return 'none'; const style = item.val === 'double' || item.val === 'triple' ? 'double' : item.val === 'dotted' ? 'dotted' : item.val?.includes('dash') ? 'dashed' : 'solid'; return `${item.sz || 0.5}pt ${style} ${item.color && item.color !== 'auto' ? '#'+item.color : '#222'}`; };
    const sideLabels: Record<string,string> = {top:'Top Border',bottom:'Bottom Border',left:'Left Border',right:'Right Border',between:'Inside Horizontal Border',insideH:'Inside Horizontal Border',insideV:'Inside Vertical Border'};
    return <div className="col"><div className="row" style={{ alignItems: 'flex-start' }}><fieldset><legend>{tr(L,'Setting:')}</legend>{radioOptions.map(([value,label])=><label key={value}><input type="radio" name={`react-wordx-border-setting-${kind}`} checked={panel.setting===value} onChange={()=>act(kind,{action:'preset',value})}/>{tr(L,label)}</label>)}</fieldset><div className="col"><label>{tr(L,'Style:')}<select id={kind==='border'?'react-wordx-border-style':'react-wordx-page-border-style'} value={panel.pen.val || 'single'} onChange={(event)=>act(kind,{action:'pen',value:{key:'val',value:event.currentTarget.value}})}>{panel.styles.map(([value,label])=><option key={value} value={value}>{tr(L,label)}</option>)}</select></label><label>{tr(L,'Color:')}<button type="button" className="btn small" onClick={(event)=>L.ui?.colorMenu?.(event.currentTarget,{mode:'border'},(color)=>act(kind,{action:'pen',value:{key:'color',value:typeof color==='string'?color.replace('#','').toUpperCase():'auto'}}))}>{tr(L,'Choose color')}<span className="cpk-sw" style={{background:panel.pen.color&&panel.pen.color!=='auto'?'#'+panel.pen.color:'#222'}}/></button></label><label>{tr(L,'Width:')}<select id={kind==='border'?'react-wordx-border-width':'react-wordx-page-border-width'} value={panel.pen.sz || .5} onChange={(event)=>act(kind,{action:'pen',value:{key:'sz',value:Number(event.currentTarget.value)}})}>{panel.widths.map((width)=><option key={width} value={width}>{width} pt</option>)}</select></label></div><div className="col"><div className="cd-cap">{tr(L,'Preview')}</div><div className="bd-preview" aria-label={tr(L,'Border preview')}><div className="bd-box" style={{borderTop:line('top'),borderBottom:line('bottom'),borderLeft:line('left'),borderRight:line('right')}}/><div className="bd-text" style={{borderBottom:line(panel.borders.between?'between':'insideH'),borderRight:line('insideV')}}>{tr(L,'Sample text')}</div>{allowed.map((side)=><button key={side} type="button" className={`bd-btn${panel.borders[side]?' on':''}`} data-border-side={side} aria-label={tr(L,sideLabels[side]||side)} aria-pressed={!!panel.borders[side]} onClick={()=>act(kind,{action:'toggle',value:side})}>{tr(L,sideLabels[side]||side)}</button>)}</div></div></div>{kind==='border'&&state.showApply&&state.applyOptions.length>0?<label>{tr(L,'Apply to:')}<select value={state.target} onChange={(event)=>act('target',event.currentTarget.value)}>{state.applyOptions.map(([value,label])=><option key={value} value={value}>{tr(L,label)}</option>)}</select></label>:null}{kind==='page'?<label>{tr(L,'Apply to:')}<select value={state.pageApply} onChange={(event)=>act('pageApply',event.currentTarget.value)}>{state.pageOptions.map(([value,label])=><option key={value} value={value}>{tr(L,label)}</option>)}</select></label>:null}</div>;
  };
  return <div className="col" id="react-wordx-borders-body"><div className="tabs-head" role="tablist">{['Borders','Page Border','Shading'].map((label,index)=><button type="button" role="tab" aria-selected={state.tab===index} className={state.tab===index?'on':''} key={label} onClick={()=>act('tab',index)}>{tr(L,label)}</button>)}</div>{state.tab===0?renderPanel('border'):state.tab===1?renderPanel('page'):<div className="col"><div className="row"><div className="col"><div className="cd-cap">{tr(L,'Fill')}</div><button type="button" className="btn small" onClick={()=>act('fill',null)}>{tr(L,'No Fill')}</button><div className="cp-grid" style={{display:'grid',gridTemplateColumns:'repeat(8,18px)'}}>{state.colors.map((color)=><button type="button" className={`cp-sw${state.fill===color.replace('#','').toUpperCase()?' on':''}`} key={color} style={{background:color}} aria-label={color} title={color} onClick={()=>act('fill',color.replace('#','').toUpperCase())}/>)}</div><button type="button" className="btn small" onClick={(event)=>L.ui?.moreColors?.(state.fill?'#'+state.fill:null,(color)=>act('fill',color?.replace('#','').toUpperCase()||null))}>{tr(L,'More Colors...')}</button></div><div className="col"><div className="cd-cap">{tr(L,'Preview')}</div><div className="preview-box" style={{width:150,height:90,background:state.fill?'#'+state.fill:'#fff'}}/></div></div><fieldset><legend>{tr(L,'Patterns')}</legend><label>{tr(L,'Style:')}<select id="react-wordx-shading-pattern" value={state.pattern} onChange={(event)=>act('pattern',event.currentTarget.value)}>{state.patterns.map(([value,label])=><option key={value} value={value}>{tr(L,label)}</option>)}</select></label><label>{tr(L,'Color:')}<button type="button" className="btn small" onClick={(event)=>L.ui?.colorMenu?.(event.currentTarget,{mode:'border'},(color)=>act('patternColor',typeof color==='string'?color.replace('#','').toUpperCase():'auto'))}>{tr(L,'Choose color')}<span className="cpk-sw" style={{background:state.patternColor&&state.patternColor!=='auto'?'#'+state.patternColor:'#222'}}/></button></label></fieldset>{state.showApply&&state.applyOptions.length>0?<label>{tr(L,'Apply to:')}<select value={state.shadingApply} onChange={(event)=>act('shadingApply',event.currentTarget.value)}>{state.applyOptions.map(([value,label])=><option key={value} value={value}>{tr(L,label)}</option>)}</select></label>:null}</div>}</div>;
}

function ParagraphDialogBody({ L }: { L: WordxRuntime }) {
  const adapter = L.dlg as (typeof L.dlg & { reactParagraphCatalog?(): ParagraphDialogState; reactParagraphAction?(action: string, value?: unknown): ParagraphDialogState | void }) | undefined;
  const [state, setState] = useState<ParagraphDialogState>(adapter?.reactParagraphCatalog?.() || { jc: 'left', l: 0, r: 0, fl: 0, b: 0, a: 0, line: 1, rule: 'auto', contextual: false, widow: true, keepNext: false, keepLines: false, pageBreakBefore: false, suppressLineNumbers: false, suppressAutoHyphens: false, sameIndents: true, changed: [], levels: [] });
  const [tab, setTab] = useState(0);
  useEffect(() => L.bus?.on('wordx-paragraph-refresh', () => { const next = adapter?.reactParagraphCatalog?.(); if (next) setState(next); }), [L, adapter]);
  const update = (key: string, value: string | number | boolean) => {
    const next = adapter?.reactParagraphAction?.('set', { key, value });
    setState(next || { ...state, [key]: value, changed: [...state.changed, key] });
  };
  const alignments: Array<[string, string]> = [['left', 'Left'], ['center', 'Centered'], ['right', 'Right'], ['both', 'Justified'], ['distribute', 'Distributed']];
  const specials: Array<[string, string]> = [['none', '(none)'], ['first', 'First line'], ['hang', 'Hanging']];
  const special = state.fl > 0 ? 'first' : state.fl < 0 ? 'hang' : 'none';
  const lineMode = state.rule === 'exact' || state.rule === 'atLeast' ? state.rule : state.line === 1 ? 'single' : state.line === 1.5 ? '1.5' : state.line === 2 ? 'double' : 'multiple';
  const lineModes: Array<[string, string]> = [['single', 'Single'], ['1.5', '1.5 lines'], ['double', 'Double'], ['atLeast', 'At least'], ['exact', 'Exactly'], ['multiple', 'Multiple']];
  const checks: Array<[keyof ParagraphDialogState, string]> = [['widow', 'Widow/Orphan control'], ['keepNext', 'Keep with next'], ['keepLines', 'Keep lines together'], ['pageBreakBefore', 'Page break before'], ['suppressLineNumbers', 'Suppress line numbers'], ['suppressAutoHyphens', "Don't hyphenate"]];
  const previewStyle: CSSProperties = { margin: `${Math.max(2, state.b * .35)}px ${Math.max(0, state.r * .35)}px ${Math.max(2, state.a * .35)}px ${Math.max(0, state.l * .35)}px`, display: 'flex', flexDirection: 'column', gap: `${state.rule === 'auto' ? Math.max(1, (state.line - 1) * 5 + 2) : 2}px`, textAlign: state.jc === 'both' ? 'justify' : state.jc === 'center' || state.jc === 'right' ? state.jc : 'left' };
  return <div className="col" id="react-wordx-paragraph-body">
    <div className="tabs-head" role="tablist">{['Indents and Spacing', 'Line and Page Breaks'].map((label, index) => <button type="button" role="tab" aria-selected={tab === index} className={tab === index ? 'on' : ''} key={label} onClick={() => setTab(index)}>{tr(L, label)}</button>)}</div>
    {tab === 0 ? <div role="tabpanel" className="col">
      <fieldset><legend>{tr(L, 'General')}</legend><label>{tr(L, 'Alignment:')} <select id="react-wordx-paragraph-alignment" value={state.jc} onChange={(event) => update('jc', event.currentTarget.value)}>{alignments.map(([value, label]) => <option key={value} value={value}>{tr(L, label)}</option>)}</select></label><label>{tr(L, 'Outline level:')} <select id="react-wordx-paragraph-outline" value={state.outline == null || state.outline >= 9 ? '' : String(state.outline)} onChange={(event) => update('outline', event.currentTarget.value === '' ? 9 : Number(event.currentTarget.value))}>{state.levels.map(([value, label]) => <option key={value} value={value}>{tr(L, label)}</option>)}</select></label></fieldset>
      <fieldset><legend>{tr(L, 'Indentation')}</legend><label>{tr(L, 'Left:')} <input aria-label={tr(L, 'Left:')} type="number" step={0.25} min={-1584} value={state.sameIndents ? state.l : ''} onChange={(event) => update('l', Number(event.currentTarget.value))} /></label><label>{tr(L, 'Right:')} <input aria-label={tr(L, 'Right:')} type="number" step={0.25} min={-1584} value={state.sameIndents ? state.r : ''} onChange={(event) => update('r', Number(event.currentTarget.value))} /></label><label>{tr(L, 'Special:')} <select value={special} onChange={(event) => update('fl', event.currentTarget.value === 'none' ? 0 : (event.currentTarget.value === 'hang' ? -1 : 1) * Math.max(36, Math.abs(state.fl)))}>{specials.map(([value, label]) => <option key={value} value={value}>{tr(L, label)}</option>)}</select></label><label>{tr(L, 'By:')} <input type="number" step={0.25} min={0} value={Math.abs(state.fl)} onChange={(event) => update('fl', (special === 'hang' ? -1 : 1) * Number(event.currentTarget.value))} /></label></fieldset>
      <fieldset><legend>{tr(L, 'Spacing')}</legend><label>{tr(L, 'Before:')} <input type="number" min={0} max={1584} step={1} value={state.b} onChange={(event) => update('b', Number(event.currentTarget.value))} /> pt</label><label>{tr(L, 'After:')} <input type="number" min={0} max={1584} step={1} value={state.a} onChange={(event) => update('a', Number(event.currentTarget.value))} /> pt</label><label>{tr(L, 'Line spacing:')} <select value={lineMode} onChange={(event) => adapter?.reactParagraphAction?.('spacingMode', event.currentTarget.value) && setState(adapter.reactParagraphCatalog?.() || state)}>{lineModes.map(([value, label]) => <option key={value} value={value}>{tr(L, label)}</option>)}</select></label><label>{tr(L, 'At:')} <input type="number" min={0} step={0.5} value={state.line} disabled={lineMode === 'single' || lineMode === '1.5' || lineMode === 'double'} onChange={(event) => { update('line', Number(event.currentTarget.value)); update('rule', lineMode === 'atLeast' || lineMode === 'exact' ? lineMode : 'auto'); }} /> {lineMode === 'atLeast' || lineMode === 'exact' ? 'pt' : 'lines'}</label><label><input type="checkbox" checked={state.contextual} onChange={(event) => update('contextual', event.currentTarget.checked)} />{tr(L, "Don't add space between paragraphs of the same style")}</label></fieldset>
      <fieldset><legend>{tr(L, 'Preview')}</legend><div className="para-preview" style={{ height: 110, overflow: 'hidden', border: '1px solid var(--ui-border,#aaa)', padding: 6 }}><div style={previewStyle}>{Array.from({ length: 5 }, (_, index) => <div key={index} style={{ height: 4, width: index === 4 ? '55%' : '100%', background: '#555', marginLeft: index === 0 ? `${state.fl * .35}px` : 0 }} />)}</div></div></fieldset>
    </div> : <div role="tabpanel" className="col"><fieldset><legend>{tr(L, 'Pagination')}</legend>{checks.slice(0, 4).map(([key, label]) => <label key={key}><input type="checkbox" id={key === 'keepNext' ? 'react-wordx-paragraph-keepNext' : undefined} checked={Boolean(state[key])} onChange={(event) => update(key, event.currentTarget.checked)} />{tr(L, label)}</label>)}</fieldset><fieldset><legend>{tr(L, 'Formatting exceptions')}</legend>{checks.slice(4).map(([key, label]) => <label key={key}><input type="checkbox" checked={Boolean(state[key])} onChange={(event) => update(key, event.currentTarget.checked)} />{tr(L, label)}</label>)}</fieldset></div>}
  </div>;
}

function TabsDialogBody({ L }: { L: WordxRuntime }) {
  const adapter = L.dlg as (typeof L.dlg & {
    reactTabsCatalog?(): TabsDialogState;
    reactTabsAction?(action: string, value?: unknown): void;
  }) | undefined;
  const [state, setState] = useState(adapter?.reactTabsCatalog?.() || { tabs: [], cleared: [], position: '', alignment: 'left', leader: 'none', defaultTab: 36 });
  useEffect(() => L.bus?.on('wordx-tabs-refresh', () => { const next = adapter?.reactTabsCatalog?.(); if (next) setState(next); }), [L, adapter]);
  const update = (field: 'position' | 'alignment' | 'leader' | 'defaultTab', value: string | number) => {
    const next = { ...state, [field]: value };
    setState(next);
    adapter?.reactTabsAction?.('change', next);
  };
  const alignments: Array<[string, string]> = [['left', 'Left'], ['center', 'Center'], ['right', 'Right'], ['decimal', 'Decimal'], ['bar', 'Bar']];
  const leaders: Array<[string, string]> = [['none', '1 None'], ['dot', '2 .......'], ['hyphen', '3 -------'], ['underscore', '4 ______']];
  return <div className="col">
    <div className="row" style={{ alignItems: 'flex-start' }}>
      <div className="col" style={{ flex: 1 }}><label htmlFor="react-wordx-tabs-position">{tr(L, 'Tab stop position:')}</label><input id="react-wordx-tabs-position" type="text" value={state.position} onChange={(event) => update('position', event.currentTarget.value)} /><select aria-label={tr(L, 'Tab stops')} size={5} value={state.tabs.some((tab) => tab.value === state.position) ? state.position : ''} onChange={(event) => adapter?.reactTabsAction?.('select', event.currentTarget.value)}>{state.tabs.map((tab) => <option key={tab.value} value={tab.value}>{tab.label}</option>)}</select></div>
      <div className="col" style={{ flex: 1 }}><label htmlFor="react-wordx-tabs-default">{tr(L, 'Default tab stops:')}</label><input id="react-wordx-tabs-default" type="number" min={1} value={state.defaultTab} onChange={(event) => update('defaultTab', Number(event.currentTarget.value))} /><div className="tp-note">{state.cleared.length ? (L.t?.('Tab stops to be cleared: {0}', state.cleared.join(', ')) || '') : ''}</div></div>
    </div>
    <div className="row"><fieldset><legend>{tr(L, 'Alignment')}</legend>{alignments.map(([value, label]) => <label key={value}><input type="radio" name="react-wordx-tabs-alignment" value={value} checked={state.alignment === value} onChange={() => update('alignment', value)} />{tr(L, label)}</label>)}</fieldset><fieldset><legend>{tr(L, 'Leader')}</legend>{leaders.map(([value, label]) => <label key={value}><input type="radio" name="react-wordx-tabs-leader" value={value} checked={state.leader === value} onChange={() => update('leader', value)} />{tr(L, label)}</label>)}</fieldset></div>
    <div className="tp-btns"><button type="button" className="btn small" onClick={() => adapter?.reactTabsAction?.('set')}>{tr(L, 'Set')}</button><button type="button" className="btn small" onClick={() => adapter?.reactTabsAction?.('clear')}>{tr(L, 'Clear')}</button><button type="button" className="btn small" onClick={() => adapter?.reactTabsAction?.('clearAll')}>{tr(L, 'Clear All')}</button></div>
  </div>;
}

function ChangeCaseDialogBody({ L }: { L: WordxRuntime }) {
  const adapter = L.dlg as (typeof L.dlg & { reactChangeCaseCatalog?(): Array<[string, string]> }) | undefined;
  const options = adapter?.reactChangeCaseCatalog?.() || [['sentence', 'Sentence case.'], ['lower', 'lowercase'], ['upper', 'UPPERCASE'], ['title', 'Title Case'], ['toggle', 'tOGGLE cASE']];
  const [selected, setSelected] = useState('sentence');
  return <fieldset>{options.map(([value, label]) => <label key={value}><input type="radio" name="react-wordx-change-case" value={value} checked={selected === value} onChange={() => setSelected(value)} />{tr(L, label)}</label>)}</fieldset>;
}

function TextDirectionDialogBody({ L }: { L: WordxRuntime }) {
  const adapter = L.dlg as (typeof L.dlg & { reactTextDirectionState?(): string }) | undefined;
  const [direction, setDirection] = useState(adapter?.reactTextDirectionState?.() || 'horz');
  const options: Array<[string, string]> = [['horz', 'Horizontal'], ['vert', 'Vertical (top to bottom)'], ['vert270', 'Vertical (bottom to top)']];
  return <div className="col"><div className="tp-note">{tr(L, 'Text direction applies to text boxes and table cells.')}</div><fieldset>{options.map(([value, label]) => <label key={value}><input type="radio" name="react-wordx-text-direction" value={value} checked={direction === value} onChange={() => setDirection(value)} />{tr(L, label)}</label>)}</fieldset></div>;
}

type ThemeOption = { value: string; label: string; body: string; head: string; headingColor: string; background: string | null; bodyStack: string; headStack: string };
function ThemeDialogBody({ L }: { L: WordxRuntime }) {
  const adapter = L.dlg as (typeof L.dlg & { reactThemeCatalog?(): ThemeOption[]; reactThemeAction?(action: string, value?: string): void }) | undefined;
  const themes = adapter?.reactThemeCatalog?.() || [];
  const [selected, setSelected] = useState('');
  const theme = themes.find((option) => option.value === selected) || themes[0];
  const onSelect = (value: string) => { setSelected(value); adapter?.reactThemeAction?.('select', value); };
  return <div className="row" style={{ alignItems: 'flex-start' }}><div className="col" style={{ width: 150 }}><label htmlFor="react-wordx-theme-select">{tr(L, 'Choose a Theme:')}</label><select id="react-wordx-theme-select" size={10} value={selected} onChange={(event) => onSelect(event.currentTarget.value)}>{themes.map((option) => <option key={option.value || '(none)'} value={option.value}>{tr(L, option.label)}</option>)}</select></div><div id="react-wordx-theme-preview" className="preview-box col" style={{ width: 260, height: 180, flexDirection: 'column', alignItems: 'flex-start', justifyContent: 'flex-start', padding: 12, gap: 6, background: theme?.background ? '#' + theme.background : '#fff' }}><div style={{ font: `bold 18px ${theme?.headStack || 'Arial'}`, color: '#' + (theme?.headingColor || '000') }}>{tr(L, 'Heading 1 style')}</div><div style={{ font: `bold 14px ${theme?.headStack || 'Arial'}`, color: '#' + (theme?.headingColor || '000') }}>{tr(L, 'Heading 2 style')}</div><div style={{ font: `12px ${theme?.bodyStack || 'Times New Roman'}` }}>{tr(L, 'Regular Text Sample — the quick brown fox jumps over the lazy dog.')}</div><div style={{ fontSize: 12 }}><u style={{ color: '#00f' }}>{tr(L, 'Hyperlink')}</u></div></div></div>;
}

function LanguageDialogBody({ L }: { L: WordxRuntime }) {
  const catalog = L.dlg?.reactLanguageCatalog?.();
  const [language, setLanguage] = useState(catalog?.language || 'en-US');
  const [noProof, setNoProof] = useState(catalog?.noProof || false);
  const [detect, setDetect] = useState(catalog?.detect ?? true);
  return <div className="col"><div>{tr(L, 'Mark selected text as:')}</div>
    <select id="react-wordx-language-select" size={8} value={language} onChange={(event) => { const value = event.currentTarget.value; setLanguage(value); L.dlg?.reactLanguageAction?.('language', value); }}>
      {(catalog?.languages || []).map((option) => <option key={option.value} value={option.value}>{tr(L, option.label)}</option>)}
    </select>
    <div className="tp-note">{tr(L, 'The speller and other proofing tools use the dictionaries of the selected language.')}</div>
    <label><input id="react-wordx-language-no-proof" type="checkbox" checked={noProof} onChange={(event) => { const value = event.currentTarget.checked; setNoProof(value); L.dlg?.reactLanguageAction?.('noProof', String(value)); }} />{tr(L, 'Do not check spelling or grammar')}</label>
    <label><input id="react-wordx-language-detect" type="checkbox" checked={detect} onChange={(event) => setDetect(event.currentTarget.checked)} />{tr(L, 'Detect language automatically')}</label>
  </div>;
}

function DropCapDialogBody({ L }: { L: WordxRuntime }) {
  const catalog = L.dlg?.reactDropCapCatalog?.();
  const [position, setPosition] = useState(catalog?.position || 'none');
  const [font, setFont] = useState(catalog?.font || 'Times New Roman');
  const [lines, setLines] = useState(catalog?.lines || 3);
  const [distance, setDistance] = useState(catalog?.distanceText || '0 pt');
  const samples: Record<string, string> = { none: 'M6 8h40M6 14h40M6 20h40M6 26h40M6 32h40', drop: 'M22 8h24M22 14h24M22 20h24M6 26h40M6 32h40', margin: 'M14 8h34M14 14h34M14 20h34M14 26h34M14 32h34' };
  const selectPosition = (value: string) => { setPosition(value); L.dlg?.reactDropCapAction?.('position', value); };
  return <div className="col"><fieldset><legend>{tr(L, 'Position')}</legend><div className="preset-grid">{(catalog?.positions || []).map(([value, label]) => <button key={value} type="button" className={`drop-cap-preset${position === value ? ' on' : ''}`} aria-pressed={position === value} onClick={() => selectPosition(value)}><svg width="52" height="40" viewBox="0 0 52 40" aria-hidden="true"><path d={samples[value] || samples.none} stroke="#999" /><text x={value === 'margin' ? 0 : value === 'drop' ? 4 : 6} y="22" fontSize={value === 'none' ? 10 : 20} fontFamily="Times" fontWeight="bold">W</text></svg><span>{tr(L, label.replace('&', ''))}</span></button>)}</div></fieldset>
    <fieldset><legend>{tr(L, 'Options')}</legend><label htmlFor="react-wordx-drop-cap-font">{tr(L, 'Font:')}</label><select id="react-wordx-drop-cap-font" translate="no" value={font} onChange={(event) => { const value = event.currentTarget.value; setFont(value); L.dlg?.reactDropCapAction?.('font', value); }}>{(catalog?.fonts || []).map((item) => <option key={item} value={item}>{item}</option>)}</select>
      <label htmlFor="react-wordx-drop-cap-lines">{tr(L, 'Lines to drop:')}</label><input id="react-wordx-drop-cap-lines" type="number" min={1} max={10} value={lines} onChange={(event) => { const value = Math.max(1, Math.min(10, Number(event.currentTarget.value) || 1)); setLines(value); L.dlg?.reactDropCapAction?.('lines', String(value)); }} />
      <label htmlFor="react-wordx-drop-cap-distance">{tr(L, 'Distance from text:')}</label><input id="react-wordx-drop-cap-distance" type="text" inputMode="decimal" value={distance} onChange={(event) => { const value = event.currentTarget.value; setDistance(value); L.dlg?.reactDropCapAction?.('distance', value); }} />
    </fieldset></div>;
}

type LineNumbersState = { enabled: boolean; start: number; distance: string; countBy: number; restart: string; restarts: Array<[string, string]> };
function LineNumbersDialogBody({ L }: { L: WordxRuntime }) {
  const adapter = L.dlg;
  const [state, setState] = useState<LineNumbersState>(adapter?.reactLineNumbersCatalog?.() || { enabled: false, start: 1, distance: 'Auto', countBy: 1, restart: 'newPage', restarts: [['newPage', 'Restart each page'], ['newSection', 'Restart each section'], ['continuous', 'Continuous']] });
  const update = (next: Partial<LineNumbersState>) => { const value = { ...state, ...next }; setState(value); adapter?.reactLineNumbersAction?.('change', value); };
  return <div className="col">
    <label><input id="react-wordx-line-numbers-enabled" type="checkbox" checked={state.enabled} onChange={(event) => update({ enabled: event.currentTarget.checked })} />{tr(L, 'Add line numbering')}</label>
    <label htmlFor="react-wordx-line-numbers-start">{tr(L, 'Start at:')}</label><input id="react-wordx-line-numbers-start" type="number" min={1} value={state.start} disabled={!state.enabled} onChange={(event) => update({ start: Math.max(1, Number(event.currentTarget.value) || 1) })} />
    <label htmlFor="react-wordx-line-numbers-distance">{tr(L, 'From text:')}</label><input id="react-wordx-line-numbers-distance" type="text" inputMode="decimal" value={state.distance} disabled={!state.enabled} onChange={(event) => update({ distance: event.currentTarget.value })} />
    <label htmlFor="react-wordx-line-numbers-count">{tr(L, 'Count by:')}</label><input id="react-wordx-line-numbers-count" type="number" min={1} max={100} value={state.countBy} disabled={!state.enabled} onChange={(event) => update({ countBy: Math.max(1, Math.min(100, Number(event.currentTarget.value) || 1)) })} />
    <fieldset disabled={!state.enabled}><legend>{tr(L, 'Numbering:')}</legend>{state.restarts.map(([value, label]) => <label key={value}><input type="radio" name="react-wordx-line-numbers-restart" value={value} checked={state.restart === value} onChange={() => update({ restart: value })} />{tr(L, label)}</label>)}</fieldset>
  </div>;
}

type ColumnsDialogState = { n: number; eq: boolean; sep: boolean; rows: Array<{ width: string; spacing: string }>; applyTo: string; newColumn: boolean; applyOptions?: Array<[string, string]> };
function ColumnsDialogBody({ L }: { L: WordxRuntime }) {
  const adapter = L.dlg;
  const [state, setState] = useState<ColumnsDialogState>(adapter?.reactColumnsCatalog?.() || { n: 1, eq: true, sep: false, rows: [], applyTo: 'all', newColumn: false });
  const update = (next: Partial<ColumnsDialogState>) => { const value = { ...state, ...next }; setState(value); adapter?.reactColumnsAction?.('change', value); };
  const presets: Array<{ label: string; n: number; eq: boolean; fractions?: number[] }> = [{ label: 'One', n: 1, eq: true }, { label: 'Two', n: 2, eq: true }, { label: 'Three', n: 3, eq: true }, { label: 'Left', n: 2, eq: false, fractions: [.35, .65] }, { label: 'Right', n: 2, eq: false, fractions: [.65, .35] }];
  const rows = state.rows.slice(0, Math.min(state.n, 4));
  const weights = rows.map((row) => state.eq ? 1 : Math.max(1, Number.parseFloat(row.width) || 1));
  const total = weights.reduce((sum, value) => sum + value, 0) || 1;
  const updateRow = (index: number, field: 'width' | 'spacing', value: string) => update({ rows: state.rows.map((row, i) => i === index ? { ...row, [field]: value } : row) });
  return <div className="col"><fieldset><legend>{tr(L, 'Presets')}</legend><div className="preset-grid">{presets.map((preset) => <button key={preset.label} type="button" className="column-preset" aria-pressed={state.n === preset.n && state.eq === preset.eq} onClick={() => { adapter?.reactColumnsAction?.('preset', { n: preset.n, eq: preset.eq, fractions: preset.fractions }); setState(adapter?.reactColumnsCatalog?.() || state); }}><span className="colsample" aria-hidden="true">{Array.from({ length: preset.n }, (_, i) => <span key={i} style={{ flex: preset.fractions?.[i] || 1 }} />)}</span><span>{tr(L, preset.label)}</span></button>)}</div></fieldset>
    <div className="row"><label htmlFor="react-wordx-columns-count">{tr(L, 'Number of columns:')}</label><input id="react-wordx-columns-count" type="number" min={1} max={13} value={state.n} onChange={(event) => { const n = Math.max(1, Math.min(13, Number(event.currentTarget.value) || 1)); const nextRows = Array.from({ length: Math.min(n, 4) }, (_, i) => state.rows[i] || { width: '', spacing: '36 pt' }); update({ n, rows: nextRows }); }} /><label><input id="react-wordx-columns-separator" type="checkbox" checked={state.sep} onChange={(event) => update({ sep: event.currentTarget.checked })} />{tr(L, 'Line between')}</label></div>
    <fieldset><legend>{tr(L, 'Width and spacing')}</legend><label><input id="react-wordx-columns-equal" type="checkbox" checked={state.eq} onChange={(event) => update({ eq: event.currentTarget.checked })} />{tr(L, 'Equal column width')}</label><div className="grid3" style={{ fontWeight: 'bold' }}><span>{tr(L, 'Col #:')}</span><span>{tr(L, 'Width:')}</span><span>{tr(L, 'Spacing:')}</span></div>{rows.map((row, index) => <div className="grid3" key={index}><span>{index + 1}:</span><input aria-label={L.t?.('Column {0} width', index + 1) || `Column ${index + 1} width`} value={row.width} disabled={state.eq} onChange={(event) => updateRow(index, 'width', event.currentTarget.value)} /><input aria-label={L.t?.('Column {0} spacing', index + 1) || `Column ${index + 1} spacing`} value={row.spacing} disabled={state.eq || index === rows.length - 1} onChange={(event) => updateRow(index, 'spacing', event.currentTarget.value)} /></div>)}</fieldset>
    <div className="row"><label htmlFor="react-wordx-columns-apply">{tr(L, 'Apply to:')}</label><select id="react-wordx-columns-apply" value={state.applyTo} onChange={(event) => update({ applyTo: event.currentTarget.value })}>{(state.applyOptions || []).map(([value, label]) => <option key={value} value={value}>{tr(L, label)}</option>)}</select><label><input id="react-wordx-columns-new" type="checkbox" checked={state.newColumn} onChange={(event) => update({ newColumn: event.currentTarget.checked })} />{tr(L, 'Start new column')}</label></div>
    <div className="col"><div className="cd-cap">{tr(L, 'Preview')}</div><div className="colsample" id="react-wordx-columns-preview" style={{ width: 110, height: 130 }}>{weights.map((weight, index) => <span key={index} style={{ flex: weight / total, borderRight: state.sep && index < weights.length - 1 ? '1px solid currentColor' : undefined }}><i /><i /><i /><i /></span>)}</div></div>
  </div>;
}

type PageSetupState = { mt: number; mb: number; ml: number; mr: number; gutter: number; pgW: number; pgH: number; orient: string; type: string; titlePg: boolean; hdr: number; ftr: number; vAlign: string; lnNum: null | { start: number; distance: number | null; countBy: number; restart: string }; evenOdd: boolean; multi: string; gutterAtTop: boolean; applyTo: string; tab: number; paper: string; papers: Array<{ value: string; label: string; width: number; height: number }>; starts: Array<[string, string]>; applyOptions: Array<[string, string]> };
function PageSetupDialogBody({ L }: { L: WordxRuntime }) {
  const adapter = L.dlg;
  const [state, setState] = useState<PageSetupState>(adapter?.reactPageSetupCatalog?.() || { mt: 72, mb: 72, ml: 72, mr: 72, gutter: 0, pgW: 612, pgH: 792, orient: 'portrait', type: 'nextPage', titlePg: false, hdr: 36, ftr: 36, vAlign: 'top', lnNum: null, evenOdd: false, multi: 'normal', gutterAtTop: false, applyTo: 'all', tab: 0, paper: 'Letter', papers: [], starts: [], applyOptions: [] });
  const update = (next: Partial<PageSetupState>) => { const value = { ...state, ...next }; setState(value); adapter?.reactPageSetupAction?.('change', value); };
  const paper = state.papers.find((item) => item.value === state.paper);
  const scale = 120 / Math.max(state.pgW, state.pgH);
  const pageWidth = Math.max(60, state.pgW * scale), pageHeight = Math.max(80, state.pgH * scale);
  const margin = (value: number) => `${Math.max(0, Math.abs(value) * scale)}px`;
  const dimensionChange = (field: 'pgW' | 'pgH', text: string) => update({ [field]: Math.max(72, Number(text) || 72), paper: 'Custom size' });
  const choosePaper = (value: string) => { const item = state.papers.find((p) => p.value === value); update(item ? { paper: value, pgW: state.orient === 'landscape' ? item.height : item.width, pgH: state.orient === 'landscape' ? item.width : item.height } : { paper: value }); };
  const orientation = (value: 'portrait' | 'landscape') => { if (value === state.orient) return; update({ orient: value, pgW: state.pgH, pgH: state.pgW, paper: state.paper }); };
  const tab = (value: number) => update({ tab: value });
  return <div className="col page-setup-react">
    <div role="tablist" aria-label={tr(L, 'Page Setup')} className="tabs-head">{['Margins', 'Paper', 'Layout'].map((label, index) => <button key={label} type="button" role="tab" aria-selected={state.tab === index} onClick={() => tab(index)}>{tr(L, label)}</button>)}</div>
    <div role="tabpanel" className="page-setup-panel">
      {state.tab === 0 && <><fieldset><legend>{tr(L, 'Margins')}</legend><div className="grid2">{([['mt', 'Top:'], ['ml', 'Left:'], ['mb', 'Bottom:'], ['mr', 'Right:'], ['gutter', 'Gutter:']] as Array<[keyof PageSetupState, string]>).map(([field, label]) => <label key={field}>{tr(L, label)}<input type="number" step="1" value={Number(state[field])} onChange={(event) => update({ [field]: Number(event.currentTarget.value) } as Partial<PageSetupState>)} /> pt</label>)}</div><label>{tr(L, 'Gutter position:')}<select value={state.gutterAtTop ? 'top' : 'left'} onChange={(event) => update({ gutterAtTop: event.currentTarget.value === 'top' })}><option value="left">{tr(L, 'Left')}</option><option value="top">{tr(L, 'Top')}</option></select></label></fieldset>
        <fieldset><legend>{tr(L, 'Orientation')}</legend>{(['portrait', 'landscape'] as const).map((value) => <button key={value} type="button" aria-pressed={state.orient === value} className={state.orient === value ? 'on' : ''} onClick={() => orientation(value)}>{tr(L, value === 'portrait' ? 'Portrait' : 'Landscape')}</button>)}</fieldset>
        <label>{tr(L, 'Multiple pages:')}<select value={state.multi} onChange={(event) => update({ multi: event.currentTarget.value })}><option value="normal">{tr(L, 'Normal')}</option><option value="mirror">{tr(L, 'Mirror margins')}</option><option value="twoOnOne">{tr(L, '2 pages per sheet')}</option><option value="bookFold">{tr(L, 'Book fold')}</option></select></label></>}
      {state.tab === 1 && <><label>{tr(L, 'Paper size:')}<select value={state.paper} onChange={(event) => choosePaper(event.currentTarget.value)}>{state.papers.map((item) => <option key={item.value} value={item.value}>{tr(L, item.label)}</option>)}<option value="Custom size">{tr(L, 'Custom size')}</option></select></label><div className="grid2"><label>{tr(L, 'Width:')}<input type="number" min={72} value={state.pgW} onChange={(event) => dimensionChange('pgW', event.currentTarget.value)} /> pt</label><label>{tr(L, 'Height:')}<input type="number" min={72} value={state.pgH} onChange={(event) => dimensionChange('pgH', event.currentTarget.value)} /> pt</label></div><fieldset><legend>{tr(L, 'Paper source')}</legend><div className="tp-note">{tr(L, 'Default tray (Automatically Select). Paper source is chosen in the print dialog of the program that prints the PDF.')}</div></fieldset></>}
      {state.tab === 2 && <><label>{tr(L, 'Section start:')}<select value={state.type} onChange={(event) => update({ type: event.currentTarget.value })}>{state.starts.map(([value, label]) => <option key={value} value={value}>{tr(L, label)}</option>)}</select></label><fieldset><legend>{tr(L, 'Headers and footers')}</legend><label><input type="checkbox" checked={state.evenOdd} onChange={(event) => update({ evenOdd: event.currentTarget.checked })} />{tr(L, 'Different odd and even')}</label><label><input type="checkbox" checked={state.titlePg} onChange={(event) => update({ titlePg: event.currentTarget.checked })} />{tr(L, 'Different first page')}</label><div className="grid2"><label>{tr(L, 'From edge: Header:')}<input type="number" value={state.hdr} onChange={(event) => update({ hdr: Number(event.currentTarget.value) })} /> pt</label><label>{tr(L, 'Footer:')}<input type="number" value={state.ftr} onChange={(event) => update({ ftr: Number(event.currentTarget.value) })} /> pt</label></div></fieldset><label>{tr(L, 'Vertical alignment:')}<select value={state.vAlign} onChange={(event) => update({ vAlign: event.currentTarget.value })}>{([['top', 'Top'], ['center', 'Center'], ['both', 'Justified'], ['bottom', 'Bottom']] as Array<[string, string]>).map(([value, label]) => <option key={value} value={value}>{tr(L, label)}</option>)}</select></label><div className="tp-btns"><button type="button" className="btn small" onClick={() => adapter?.reactPageSetupAction?.('lineNumbers')}>{tr(L, 'Line Numbers...')}</button><button type="button" className="btn small" onClick={() => adapter?.reactPageSetupAction?.('borders')}>{tr(L, 'Borders...')}</button></div></>}
    </div>
    <div className="row"><label htmlFor="react-wordx-page-setup-apply">{tr(L, 'Apply to:')}</label><select id="react-wordx-page-setup-apply" value={state.applyTo} onChange={(event) => update({ applyTo: event.currentTarget.value })}>{state.applyOptions.map(([value, label]) => <option key={value} value={value}>{tr(L, label)}</option>)}</select></div>
    <div className="col" style={{ alignItems: 'center' }}><div className="cd-cap">{tr(L, 'Preview')}</div><div className="ps-page" id="react-wordx-page-setup-preview" style={{ position: 'relative', width: pageWidth, height: pageHeight, margin: 6 }}><div style={{ position: 'absolute', top: margin(state.mt), bottom: margin(state.mb), left: margin(state.ml + state.gutter), right: margin(state.mr), overflow: 'hidden' }}>{Array.from({ length: 12 }, (_, i) => <i key={i} style={{ display: 'block', borderTop: '1px solid #999', marginTop: 5 }} />)}</div></div></div>
  </div>;
}

type CaptionNumbering = { fmt: string; chapter: boolean; chapterStyle: number; sep: string };
type CaptionDialogState = { labels: string[]; label: string; position: string; exclude: boolean; text: string; numbering: CaptionNumbering };
function CaptionDialogBody({ L }: { L: WordxRuntime }) {
  const adapter = L.dlg;
  const [state, setState] = useState<CaptionDialogState>(adapter?.reactCaptionCatalog?.() || { labels: ['Figure', 'Table'], label: 'Figure', position: 'below', exclude: false, text: 'Figure 1', numbering: { fmt: 'ARABIC', chapter: false, chapterStyle: 1, sep: '-' } });
  const refresh = () => setState(adapter?.reactCaptionCatalog?.() || state);
  useEffect(() => L.bus?.on('wordx-caption-refresh', refresh), [L]);
  const number = (value: CaptionNumbering) => value.fmt === 'alphabetic' ? 'a' : value.fmt === 'ALPHABETIC' ? 'A' : value.fmt === 'roman' ? 'i' : value.fmt === 'ROMAN' ? 'I' : '1';
  const preview = (label: string, exclude: boolean, numbering: CaptionNumbering) => `${exclude ? '' : `${label} `}${numbering.chapter ? `1${numbering.sep}` : ''}${number(numbering)}`;
  const update = (next: Partial<CaptionDialogState>, regenerate = false) => { const value = { ...state, ...next }; if (regenerate) value.text = preview(value.label, value.exclude, value.numbering); setState(value); void adapter?.reactCaptionAction?.('change', value); };
  const labelChange = (label: string) => update({ label }, true);
  return <div className="col" id="react-wordx-caption-body">
    <label htmlFor="react-wordx-caption-text">{tr(L, 'Caption:')}</label><input autoFocus id="react-wordx-caption-text" type="text" value={state.text} onChange={(event) => update({ text: event.currentTarget.value })} />
    <fieldset><legend>{tr(L, 'Options')}</legend><label htmlFor="react-wordx-caption-label">{tr(L, 'Label:')}</label><select id="react-wordx-caption-label" value={state.label} onChange={(event) => labelChange(event.currentTarget.value)}>{state.labels.map((label) => <option key={label} value={label} translate="no">{label}</option>)}</select>
      <label htmlFor="react-wordx-caption-position">{tr(L, 'Position:')}</label><select id="react-wordx-caption-position" value={state.position} onChange={(event) => update({ position: event.currentTarget.value })}><option value="below">{tr(L, 'Below selected item')}</option><option value="above">{tr(L, 'Above selected item')}</option></select>
      <label><input id="react-wordx-caption-no-label" type="checkbox" checked={state.exclude} onChange={(event) => update({ exclude: event.currentTarget.checked }, true)} />{tr(L, 'Exclude label from caption')}</label>
    </fieldset>
    <div className="tp-btns"><button type="button" className="btn small" onClick={async () => { await adapter?.reactCaptionAction?.('newLabel'); refresh(); }}>{tr(L, 'New Label...')}</button><button type="button" className="btn small" disabled={!state.label || state.labels.length < 1} onClick={async () => { await adapter?.reactCaptionAction?.('deleteLabel'); refresh(); }}>{tr(L, 'Delete Label')}</button><button type="button" className="btn small" onClick={() => adapter?.reactCaptionAction?.('numbering')}>{tr(L, 'Numbering...')}</button></div>
  </div>;
}

function CaptionNumberingDialogBody({ L }: { L: WordxRuntime }) {
  const adapter = L.dlg;
  const [state, setState] = useState<CaptionNumbering>(adapter?.reactCaptionNumberingCatalog?.() || { fmt: 'ARABIC', chapter: false, chapterStyle: 1, sep: '-' });
  const update = (next: Partial<CaptionNumbering>) => { const value = { ...state, ...next }; setState(value); adapter?.reactCaptionNumberingAction?.(value); };
  return <div className="col" id="react-wordx-caption-numbering-body"><label htmlFor="react-wordx-caption-format">{tr(L, 'Format:')}</label><select id="react-wordx-caption-format" value={state.fmt} onChange={(event) => update({ fmt: event.currentTarget.value })}>{([['ARABIC','1, 2, 3, ...'],['alphabetic','a, b, c, ...'],['ALPHABETIC','A, B, C, ...'],['roman','i, ii, iii, ...'],['ROMAN','I, II, III, ...']] as Array<[string,string]>).map(([value,label]) => <option key={value} value={value}>{tr(L,label)}</option>)}</select>
    <label><input id="react-wordx-caption-chapter" type="checkbox" checked={state.chapter} onChange={(event) => update({ chapter: event.currentTarget.checked })} />{tr(L, 'Include chapter number')}</label>
    <label htmlFor="react-wordx-caption-chapter-style">{tr(L, 'Chapter starts with style:')}</label><select id="react-wordx-caption-chapter-style" value={state.chapterStyle} disabled={!state.chapter} onChange={(event) => update({ chapterStyle: Number(event.currentTarget.value) })}>{Array.from({ length: 9 }, (_, i) => i + 1).map((level) => <option key={level} value={level} translate="no">Heading {level}</option>)}</select>
    <label htmlFor="react-wordx-caption-separator">{tr(L, 'Use separator:')}</label><select id="react-wordx-caption-separator" value={state.sep} disabled={!state.chapter} onChange={(event) => update({ sep: event.currentTarget.value })}>{([['-','- (hyphen)'],['.','. (period)'],[':',': (colon)'],['—','— (em dash)'],['–','– (en dash)']] as Array<[string,string]>).map(([value,label]) => <option key={value} value={value}>{tr(L,label)}</option>)}</select>
  </div>;
}

type WatermarkState = { kind: string; media: unknown; pictureName: string; scale: string | number; washout: boolean; text: string; font: string; size: string | number; color: string; semi: boolean; layout: string };
function WatermarkDialogBody({ L }: { L: WordxRuntime }) {
  const adapter = L.dlg as (typeof L.dlg & { reactWatermarkCatalog?(): WatermarkState & { fonts: string[] }; reactWatermarkAction?(action: string, value?: Partial<WatermarkState>): Promise<void> | void }) | undefined;
  const initial = adapter?.reactWatermarkCatalog?.();
  const [state, setState] = useState<WatermarkState & { fonts: string[] }>(initial || { kind: 'none', media: null, pictureName: '', scale: 'auto', washout: true, text: 'ASAP', font: 'Times New Roman', size: 'auto', color: 'C0C0C0', semi: true, layout: 'diagonal', fonts: [] });
  useEffect(() => L.bus?.on('wordx-watermark-refresh', () => { const next = adapter?.reactWatermarkCatalog?.(); if (next) setState(next); }), [L]);
  const update = (next: Partial<WatermarkState>) => { const value = { ...state, ...next }; setState(value); void adapter?.reactWatermarkAction?.('change', value); };
  const scales: Array<[string, string]> = [['auto','Auto'],['5','500%'],['2','200%'],['1.5','150%'],['1','100%'],['0.5','50%']];
  const sizes: Array<[string, string]> = [['auto','Auto'],...([36,40,44,48,54,60,66,72,80,90,96,105,120,144].map((value) => [String(value),String(value)] as [string,string]))];
  const color = `#${state.color.replace(/^#/, '').slice(0, 6)}`;
  return <div className="col" id="react-wordx-watermark-body"><fieldset><legend>{tr(L, 'Watermark')}</legend>{([['none','No watermark'],['picture','Picture watermark'],['text','Text watermark']] as Array<[string,string]>).map(([value,label]) => <label key={value}><input type="radio" name="react-wordx-watermark-kind" value={value} checked={state.kind === value} onChange={() => update({ kind: value })} />{tr(L,label)}</label>)}</fieldset>
    <fieldset disabled={state.kind !== 'picture'}><legend>{tr(L, 'Picture')}</legend><button type="button" className="btn small" onClick={() => void adapter?.reactWatermarkAction?.('picture')}>{tr(L, 'Select Picture...')}</button><span className="tp-note">{state.pictureName}</span><label>{tr(L, 'Scale:')}<select value={String(state.scale)} onChange={(event) => update({ scale: event.currentTarget.value === 'auto' ? 'auto' : Number(event.currentTarget.value) })}>{scales.map(([value,label]) => <option key={value} value={value}>{tr(L,label)}</option>)}</select></label><label><input type="checkbox" checked={state.washout} onChange={(event) => update({ washout: event.currentTarget.checked })} />{tr(L, 'Washout')}</label></fieldset>
    <fieldset disabled={state.kind !== 'text'}><legend>{tr(L, 'Text')}</legend><label htmlFor="react-wordx-watermark-text">{tr(L, 'Text:')}</label><input id="react-wordx-watermark-text" list="react-wordx-watermark-text-list" value={state.text} onChange={(event) => update({ text: event.currentTarget.value })} /><datalist id="react-wordx-watermark-text-list">{['ASAP','ATTORNEY-CLIENT PRIVILEGE','CONFIDENTIAL','COPY','DO NOT COPY','DRAFT','ORIGINAL','PERSONAL','SAMPLE','TOP SECRET','URGENT'].map((value) => <option key={value} value={value} />)}</datalist>
      <label>{tr(L, 'Font:')}<select translate="no" value={state.font} onChange={(event) => update({ font: event.currentTarget.value })}>{state.fonts.map((font) => <option key={font} value={font}>{font}</option>)}</select></label><label>{tr(L, 'Size:')}<select value={String(state.size)} onChange={(event) => update({ size: event.currentTarget.value === 'auto' ? 'auto' : Number(event.currentTarget.value) })}>{sizes.map(([value,label]) => <option key={value} value={value}>{tr(L,label)}</option>)}</select></label>
      <label>{tr(L, 'Color:')}<input type="color" aria-label={tr(L, 'Color:')} value={color} onChange={(event) => update({ color: event.currentTarget.value.slice(1).toUpperCase() })} /></label><label><input type="checkbox" checked={state.semi} onChange={(event) => update({ semi: event.currentTarget.checked })} />{tr(L, 'Semitransparent')}</label><fieldset><legend>{tr(L, 'Layout:')}</legend>{([['diagonal','Diagonal'],['horizontal','Horizontal']] as Array<[string,string]>).map(([value,label]) => <label key={value}><input type="radio" name="react-wordx-watermark-layout" value={value} checked={state.layout === value} onChange={() => update({ layout: value })} />{tr(L,label)}</label>)}</fieldset>
    </fieldset>
  </div>;
}

type HyperlinkPlace = { value: string; label: string; group?: boolean };
type HyperlinkState = { kind: number; display: string; address: string; place: string; email: string; subject: string; screenTip: string; places: HyperlinkPlace[] };
function HyperlinkDialogBody({ L }: { L: WordxRuntime }) {
  const adapter = L.dlg as (typeof L.dlg & { reactHyperlinkCatalog?(): HyperlinkState; reactHyperlinkAction?(action: string, value?: Partial<HyperlinkState>): Promise<void> | void }) | undefined;
  const [state, setState] = useState<HyperlinkState>(adapter?.reactHyperlinkCatalog?.() || { kind: 0, display: '', address: '', place: '', email: '', subject: '', screenTip: '', places: [] });
  useEffect(() => L.bus?.on('wordx-hyperlink-refresh', () => { const next = adapter?.reactHyperlinkCatalog?.(); if (next) setState(next); }), [L]);
  const update = (next: Partial<HyperlinkState>) => { const value = { ...state, ...next }; setState(value); void adapter?.reactHyperlinkAction?.('change', value); };
  const tabs: Array<[number, string]> = [[0, 'Existing File or Web Page'], [1, 'Place in This Document'], [2, 'E-mail Address']];
  const placeGroups: Array<{ label: string; options: HyperlinkPlace[] }> = [];
  for (const place of state.places) {
    if (place.group) placeGroups.push({ label: place.label, options: [] });
    else if (placeGroups.length) placeGroups.at(-1)?.options.push(place);
  }
  return <div className="col" id="react-wordx-hyperlink-body">
    <div className="row"><label htmlFor="react-wordx-hyperlink-display">{tr(L, 'Text to display:')}</label><input id="react-wordx-hyperlink-display" type="text" value={state.display} style={{ flex: 1, minWidth: 0 }} onChange={(event) => update({ display: event.currentTarget.value })} /><button className="btn small" type="button" onClick={() => void adapter?.reactHyperlinkAction?.('screenTip')}>{tr(L, 'ScreenTip...')}</button></div>
    <div role="tablist" aria-label={tr(L, 'Hyperlink type')} className="tabs-head">{tabs.map(([value, label]) => <button key={value} type="button" role="tab" aria-selected={state.kind === value} onClick={() => update({ kind: value })}>{tr(L, label)}</button>)}</div>
    {state.kind === 0 && <div role="tabpanel" className="col"><label htmlFor="react-wordx-hyperlink-address">{tr(L, 'Address:')}</label><input id="react-wordx-hyperlink-address" type="text" value={state.address} onChange={(event) => update({ address: event.currentTarget.value })} /><div className="tp-note">{tr(L, 'Type a web address (https://…) or the name of a file.')}</div></div>}
    {state.kind === 1 && <div role="tabpanel" className="col"><label htmlFor="react-wordx-hyperlink-place">{tr(L, 'Select a place in this document:')}</label><select id="react-wordx-hyperlink-place" size={9} value={state.place} onChange={(event) => update({ place: event.currentTarget.value })}><option value="">{tr(L, 'Top of the Document')}</option>{placeGroups.map((group) => <optgroup key={group.label} label={tr(L, group.label)}>{group.options.map((place, index) => <option key={`${place.value}:${index}`} value={place.value}>{tr(L, place.label)}</option>)}</optgroup>)}</select></div>}
    {state.kind === 2 && <div role="tabpanel" className="col"><label htmlFor="react-wordx-hyperlink-email">{tr(L, 'E-mail address:')}</label><input id="react-wordx-hyperlink-email" type="text" value={state.email} onChange={(event) => update({ email: event.currentTarget.value })} /><label htmlFor="react-wordx-hyperlink-subject">{tr(L, 'Subject:')}</label><input id="react-wordx-hyperlink-subject" type="text" value={state.subject} onChange={(event) => update({ subject: event.currentTarget.value })} /></div>}
  </div>;
}

type SymbolGlyph = { ch: string; raw?: string; font: string; code: string };
type SymbolState = { font: string; subset: number; tab: number; selected: string | null; code: string; recent: string[]; special: string; fonts: string[]; subsets: Array<{ value: number; label: string }>; chars: SymbolGlyph[]; specials: Array<{ value: string; label: string }> };
function SymbolDialogBody({ L }: { L: WordxRuntime }) {
  const adapter = L.dlg as (typeof L.dlg & { reactSymbolCatalog?(): SymbolState; reactSymbolAction?(action: string, value?: unknown): SymbolState }) | undefined;
  const [state, setState] = useState<SymbolState | null>(adapter?.reactSymbolCatalog?.() || null);
  useEffect(() => L.bus?.on('wordx-symbol-refresh', () => setState(adapter?.reactSymbolCatalog?.() || null)), [L]);
  if (!state) return null;
  const act = (action: string, value?: unknown) => { const next = adapter?.reactSymbolAction?.(action, value); if (next) setState(next); };
  return <div className="col" id="react-wordx-symbol-body">
    <div role="tablist" aria-label={tr(L, 'Symbol type')} className="tabs-head"><button type="button" role="tab" aria-selected={state.tab === 0} onClick={() => act('tab', 0)}>{tr(L, 'Symbols')}</button><button type="button" role="tab" aria-selected={state.tab === 1} onClick={() => act('tab', 1)}>{tr(L, 'Special Characters')}</button></div>
    {state.tab === 0 ? <><div className="row"><label htmlFor="react-wordx-symbol-font">{tr(L, 'Font:')}</label><select id="react-wordx-symbol-font" translate="no" value={state.font} onChange={(event) => act('font', event.currentTarget.value)}>{state.fonts.map((font) => <option key={font || 'normal'} value={font}>{font || tr(L, '(normal text)')}</option>)}</select><label htmlFor="react-wordx-symbol-subset">{tr(L, 'Subset:')}</label><select id="react-wordx-symbol-subset" value={state.subset} onChange={(event) => act('subset', event.currentTarget.value)}>{state.subsets.map((subset) => <option key={subset.value} value={subset.value}>{tr(L, subset.label)}</option>)}</select></div>
      <div className="sym-grid">{state.chars.map((glyph, index) => <button key={`${glyph.code}:${index}`} type="button" className={state.selected === glyph.ch ? 'on' : ''} aria-label={`${glyph.ch} ${glyph.code}`} style={glyph.font && !/Symbol|Wingdings|Webdings/.test(glyph.font) ? { fontFamily: glyph.font } : undefined} onClick={() => act('char', glyph)} onDoubleClick={() => { act('char', glyph); act('insert'); }}>{glyph.ch}</button>)}</div>
      <div className="cd-cap">{tr(L, 'Recently used symbols:')}</div><div className="sym-recent">{state.recent.map((char, index) => <button key={`${char}:${index}`} type="button" aria-label={char} onClick={() => act('recent', char)} onDoubleClick={() => { act('recent', char); act('insert'); }}>{char}</button>)}</div>
      <div className="row"><div className="tp-note" style={{ minHeight: 16 }}>{state.code ? L.t?.('Unicode character {0}', state.code) : ''}</div><label htmlFor="react-wordx-symbol-code">{tr(L, 'Character code:')}</label><input id="react-wordx-symbol-code" type="text" value={state.code} style={{ width: 70 }} onChange={(event) => act('code', event.currentTarget.value)} /></div>
    </> : <div className="col" role="tabpanel"><label htmlFor="react-wordx-symbol-special">{tr(L, 'Special Characters')}</label><select id="react-wordx-symbol-special" size={10} value={state.special} onChange={(event) => act('special', event.currentTarget.value)}>{state.specials.map((item) => <option key={item.value} value={item.value}>{item.label}</option>)}</select></div>}
  </div>;
}

type ListPreset = { index: number; lines: string[]; headings?: boolean };
type BulletsState = { tab: number; restart: boolean; selected: string | null; bullet: ListPreset[]; number: ListPreset[]; outline: ListPreset[] };
function BulletsDialogBody({ L }: { L: WordxRuntime }) {
  const adapter = L.dlg as (typeof L.dlg & { reactBulletsCatalog?(): BulletsState; reactBulletsAction?(action: string, value?: unknown): BulletsState | void }) | undefined;
  const [state, setState] = useState<BulletsState>(adapter?.reactBulletsCatalog?.() || { tab: 0, restart: false, selected: null, bullet: [], number: [], outline: [] });
  const act = (action: string, value?: unknown) => { const next = adapter?.reactBulletsAction?.(action, value); if (next) setState(next); };
  const galleries: Array<[string, string, ListPreset[]]> = [['bullet', 'Bulleted', state.bullet], ['number', 'Numbered', state.number], ['outline', 'Outline Numbered', state.outline]];
  const preview = (lines: string[]) => <div className="list-preview">{lines.map((line, index) => <div key={index} className={`l${index}`}><span>{line}</span><i /></div>)}</div>;
  return <div className="col" id="react-wordx-bullets-body"><div role="tablist" aria-label={tr(L, 'List type')} className="tabs-head">{galleries.map(([kind, label], index) => <button key={kind} type="button" role="tab" aria-selected={state.tab === index} onClick={() => act('tab', index)}>{tr(L, label)}</button>)}</div>
    {state.tab === 1 && <label><input type="checkbox" checked={state.restart} onChange={(event) => act('restart', event.currentTarget.checked)} />{tr(L, 'Restart numbering')}</label>}
    <div role="tabpanel" className="gal-grid"><button type="button" className={`none${state.selected === `${galleries[state.tab]?.[0]}:none` ? ' on' : ''}`} onClick={() => act('pick', { kind: galleries[state.tab]?.[0], none: true })}>{tr(L, 'None')}</button>{galleries[state.tab]?.[2].map((item) => <button key={item.index} type="button" className={state.selected === `${galleries[state.tab]?.[0]}:${item.index}` ? 'on' : ''} aria-label={item.lines.join(' ')} onClick={() => act('pick', { kind: galleries[state.tab]?.[0], index: item.index })} onDoubleClick={() => { act('pick', { kind: galleries[state.tab]?.[0], index: item.index }); act('apply'); }}>{preview(item.lines)}</button>)}</div>
  </div>;
}

type CustomizeListState = { multi:boolean; level:number; levels:Array<[number,string]>; text:string; format:string; formats:Array<[string,string]>; start:number; align:string; alignedAt:number; tabAfter:number; indentAt:number; follow:string; follows:Array<[string,string]>; style:string; styles:Array<[string,string]>; legal:boolean; bullet:string; preview:string[] };
function CustomizeListDialogBody({ L }: { L: WordxRuntime }) {
  const adapter=L.dlg as (typeof L.dlg & {reactCustomizeListCatalog?():CustomizeListState;reactCustomizeListAction?(action:string,value?:any):CustomizeListState})|undefined;
  const [state,setState]=useState<CustomizeListState>(adapter?.reactCustomizeListCatalog?.()||{multi:false,level:0,levels:[],text:'',format:'decimal',formats:[],start:1,align:'left',alignedAt:0,tabAfter:36,indentAt:36,follow:'tab',follows:[],style:'',styles:[],legal:false,bullet:'•',preview:[]});
  const act=(action:string,value?:unknown)=>{const next=adapter?.reactCustomizeListAction?.(action,value);if(next)setState(next);};
  const update=(key:string,value:unknown)=>act('change',{key,value});
  const field=(key:string,label:string,value:number)=><label key={key}>{tr(L,label)}<input type="number" min={0} step={0.1} value={value} onChange={(event)=>update(key,Number(event.currentTarget.value))}/></label>;
  return <div className="row" id="react-wordx-customize-list-body">{state.multi&&<label>{tr(L,'Level')}<select value={state.level} onChange={(event)=>act('level',Number(event.currentTarget.value))}>{state.levels.map(([value,label])=><option key={value} value={value}>{tr(L,label)}</option>)}</select></label>}<div className="col" style={{flex:1,minWidth:250}}>
    {state.format==='bullet'?<fieldset className="col"><legend>{tr(L,'Bullet')}</legend><div className="row"><button type="button" className="btn" onClick={()=>act('symbol')}>{state.bullet}</button><button type="button" className="btn small" onClick={()=>act('bullet-font')}>{tr(L,'Font...')}</button></div></fieldset>:<fieldset className="col"><legend>{tr(L,'Number format')}</legend><label>{tr(L,'Number format:')}<input value={state.text} onChange={(event)=>update('text',event.currentTarget.value)}/></label><button type="button" className="btn small" onClick={()=>act('number-font')}>{tr(L,'Font...')}</button></fieldset>}
    <div className="row"><label>{tr(L,'Number style:')}<select value={state.format} onChange={(event)=>update('format',event.currentTarget.value)}>{state.formats.map(([value,label])=><option key={value} value={value}>{tr(L,label)}</option>)}</select></label>{field('start','Start at:',state.start)}</div>
    {state.multi&&<div className="row"><label>{tr(L,'Link level to style:')}<select value={state.style} onChange={(event)=>update('style',event.currentTarget.value)}>{state.styles.map(([value,label])=><option key={value} value={value}>{tr(L,label)}</option>)}</select></label><label><input type="checkbox" checked={state.legal} onChange={(event)=>update('legal',event.currentTarget.checked)}/>{tr(L,'Legal style numbering')}</label></div>}
    <fieldset className="col"><legend>{tr(L,'Number position')}</legend><div className="row"><label>{tr(L,'Alignment:')}<select value={state.align} onChange={(event)=>update('align',event.currentTarget.value)}>{([['left','Left'],['center','Centered'],['right','Right']] as const).map(([value,label])=><option key={value} value={value}>{tr(L,label)}</option>)}</select></label>{field('alignedAt','Aligned at:',state.alignedAt)}</div></fieldset>
    <fieldset className="col"><legend>{tr(L,'Text position')}</legend><div className="row">{field('tabAfter','Tab space after:',state.tabAfter)}{field('indentAt','Indent at:',state.indentAt)}</div><label>{tr(L,'Follow number with:')}<select value={state.follow} onChange={(event)=>update('follow',event.currentTarget.value)}>{state.follows.map(([value,label])=><option key={value} value={value}>{tr(L,label)}</option>)}</select></label></fieldset>
    </div><div className="col"><div className="cd-cap">{tr(L,'Preview')}</div><div className="para-preview" style={{height:150,width:180,font:'9px Arial,Arimo,sans-serif',gap:5}}>{state.preview.map((line,index)=><div key={index} style={{display:'flex',gap:4,alignItems:'center'}}><span style={{minWidth:10}}>{line}</span><i style={{flex:1,height:3,background:'#999'}}/></div>)}</div></div></div>;
}

function WordartDialogBody({ L }: { L: WordxRuntime }) {
  const adapter = L.dlg as (typeof L.dlg & { reactWordartCatalog?(): { selected: number; previews: string[] }; reactWordartAction?(action: string, value?: number): { selected: number; previews: string[] } | void }) | undefined;
  const [state, setState] = useState(adapter?.reactWordartCatalog?.() || { selected: 0, previews: [] });
  const act = (action: string, value?: number) => { const next = adapter?.reactWordartAction?.(action, value); if (next) setState(next); };
  return <div className="col" id="react-wordx-wordart-body"><div>{tr(L, 'Select a LettersArt style:')}</div><div className="wa-grid">{state.previews.map((svg, index) => <button key={index} type="button" aria-label={L.t?.('LettersArt style {0}', index + 1)} className={state.selected === index ? 'on' : ''} onClick={() => act('select', index)} onDoubleClick={() => { act('select', index); act('apply'); }} dangerouslySetInnerHTML={{ __html: svg }} />)}</div></div>;
}

function AutoShapeMenu({ L }: { L: WordxRuntime }) {
  const shapes = L.app?.reactShapeMenuCatalog?.() || [];
  return <div className="shape-grid">{shapes.map((shape) => <button key={shape.kind} type="button" data-tip={tr(L, shape.label)} aria-label={tr(L, shape.label)} onClick={() => L.app?.reactShapeMenuAction?.(shape.kind)} dangerouslySetInnerHTML={{ __html: shape.svg }} />)}</div>;
}

function ObjectShadowMenu({ L }: { L: WordxRuntime }) {
  const choices = L.app?.reactShadowMenuCatalog?.() || [];
  return <div className="shape-grid" style={{ gridTemplateColumns: 'repeat(4,30px)' }}>{choices.map((choice) => <button key={choice.index} type="button" style={{ width: 30, height: 30 }} aria-label={tr(L, choice.label)} onClick={() => L.app?.reactShadowMenuAction?.(choice.index)} dangerouslySetInnerHTML={{ __html: choice.svg }} />)}</div>;
}

type LettersArtTextState = { text: string; font: string; sz: number; b: boolean; i: boolean; fonts: string[]; sizes: number[] };
function LettersArtTextDialogBody({ L }: { L: WordxRuntime }) {
  const adapter = L.dlg as (typeof L.dlg & { reactLettersArtTextCatalog?(): LettersArtTextState; reactLettersArtTextAction?(action: string, value?: Partial<LettersArtTextState>): LettersArtTextState }) | undefined;
  const [state, setState] = useState<LettersArtTextState>(adapter?.reactLettersArtTextCatalog?.() || { text: '', font: 'Arial Black', sz: 36, b: false, i: false, fonts: [], sizes: [] });
  useLayoutEffect(() => { const timer = window.setTimeout(() => document.querySelector<HTMLTextAreaElement>('#react-wordx-lettersart-text-body textarea')?.focus(), 50); return () => window.clearTimeout(timer); }, []);
  const update = (value: Partial<LettersArtTextState>) => { const next = adapter?.reactLettersArtTextAction?.('change', value); if (next) setState(next); };
  return <div className="col" id="react-wordx-lettersart-text-body">
    <div className="row">
      <label>{tr(L, 'Font:')}<select aria-label={tr(L, 'Font')} value={state.font} onChange={(event) => update({ font: event.currentTarget.value })}>{state.fonts.map((font) => <option key={font} value={font}>{font}</option>)}</select></label>
      <label>{tr(L, 'Size:')}<select aria-label={tr(L, 'Size')} value={state.sz} onChange={(event) => update({ sz: Number(event.currentTarget.value) })}>{state.sizes.map((size) => <option key={size} value={size}>{size}</option>)}</select></label>
      <label><input type="checkbox" checked={state.b} onChange={(event) => update({ b: event.currentTarget.checked })} />{tr(L, 'Bold')}</label>
      <label><input type="checkbox" checked={state.i} onChange={(event) => update({ i: event.currentTarget.checked })} />{tr(L, 'Italic')}</label>
    </div>
    <label htmlFor="react-wordx-lettersart-text">{tr(L, 'Text:')}</label>
    <textarea autoFocus id="react-wordx-lettersart-text" rows={5} value={state.text} style={{ width: '100%', fontSize: 20, fontFamily: state.font, fontWeight: state.b ? 'bold' : 'normal', fontStyle: state.i ? 'italic' : 'normal' }} onChange={(event) => update({ text: event.currentTarget.value })} />
  </div>;
}

type WordCountState = { selection: boolean; includeNotes: boolean; rows: Array<[string, string]> };
function WordCountDialogBody({ L }: { L: WordxRuntime }) {
  const adapter = L.dlg as (typeof L.dlg & { reactWordCountCatalog?(): WordCountState; reactWordCountAction?(value: boolean): WordCountState }) | undefined;
  const [state, setState] = useState<WordCountState>(adapter?.reactWordCountCatalog?.() || { selection: false, includeNotes: false, rows: [] });
  return <div className="col" id="react-wordx-word-count-body"><div className="cd-cap">{state.selection ? tr(L, 'Statistics: (selection)') : tr(L, 'Statistics:')}</div><table className="wc-table"><tbody>{state.rows.map(([label, value], index) => <tr key={`${label}:${index}`}><td>{tr(L, label)}</td><td>{value}</td></tr>)}</tbody></table><label><input type="checkbox" checked={state.includeNotes} onChange={(event) => { const next = adapter?.reactWordCountAction?.(event.currentTarget.checked); if (next) setState(next); }} />{tr(L, 'Include footnotes and endnotes')}</label></div>;
}

function AutoSummarizeDialogBody({ L }: { L: WordxRuntime }) {
  const [kind, setKind] = useState('highlight');
  const [length, setLength] = useState('25%');
  const [update, setUpdate] = useState(true);
  const types: Array<[string, string]> = [['highlight', 'Highlight key points'], ['top', 'Insert an executive summary or abstract at the top of the document'], ['new', 'Create a new document and put the summary there'], ['hide', 'Hide everything but the summary without leaving the original document']];
  const lengths: Array<[string, string]> = [['10','10 sentences'],['20','20 sentences'],['100','100 words or less'],['500','500 words or less'],['10%','10%'],['25%','25%'],['50%','50%'],['75%','75%']];
  return <div className="col" id="react-wordx-summarize-body"><div className="tp-note">{tr(L, 'Wordx has examined the document and picked the sentences most relevant to the main theme.')}</div><fieldset><legend>{tr(L, 'Type of summary')}</legend>{types.map(([value, label]) => <label key={value}><input type="radio" name="react-wordx-summary-kind" value={value} checked={kind === value} onChange={() => setKind(value)} />{tr(L, label)}</label>)}</fieldset><fieldset><legend>{tr(L, 'Length of summary')}</legend><label htmlFor="react-wordx-summary-length">{tr(L, 'Percent of original:')}</label><select id="react-wordx-summary-length" value={length} onChange={(event) => setLength(event.currentTarget.value)}>{lengths.map(([value, label]) => <option key={value} value={value}>{tr(L, label)}</option>)}</select></fieldset><label><input id="react-wordx-summary-update" type="checkbox" checked={update} onChange={(event) => setUpdate(event.currentTarget.checked)} />{tr(L, 'Update document statistics')}</label></div>;
}

function ZoomDialogBody({ L }: { L: WordxRuntime }) {
  const adapter = L.dlg as (typeof L.dlg & { reactZoomCatalog?(): { current: number; preset: string } }) | undefined;
  const initial = adapter?.reactZoomCatalog?.() || { current: 100, preset: '100' };
  const [preset, setPreset] = useState(initial.preset);
  const [percent, setPercent] = useState(initial.current);
  const options: Array<[string, string]> = [['200','200%'],['100','100%'],['75','75%'],['pw','Page width'],['tw','Text width'],['wp','Whole page'],['many','Many pages'],['pct','Percent']];
  return <div className="row" id="react-wordx-zoom-body"><fieldset><legend>{tr(L, 'Zoom to')}</legend>{options.map(([value, label]) => <label key={value}><input type="radio" name="react-wordx-zoom" value={value} checked={preset === value} onChange={() => setPreset(value)} />{tr(L, label)}</label>)}</fieldset><div className="col"><label htmlFor="react-wordx-zoom-percent">{tr(L, 'Percent:')}</label><div className="row"><input id="react-wordx-zoom-percent" type="number" min={10} max={500} step={10} value={percent} onChange={(event) => { setPercent(Number(event.currentTarget.value)); setPreset('pct'); }} />%</div><div className="cd-cap">{tr(L, 'Preview')}</div><div className="preview-box" style={{ width: 150, height: 110, overflow: 'hidden' }}><span style={{ transform: `scale(${Math.max(10, Math.min(500, percent)) / 100})`, transformOrigin: 'center' }}>AaBbCcDdEeXxYyZz</span></div></div></div>;
}

type HyphenationState = { auto: boolean; caps: boolean; zone: number; limit: number | '' };
function HyphenationDialogBody({ L }: { L: WordxRuntime }) {
  const adapter = L.dlg as (typeof L.dlg & { reactHyphenationCatalog?(): HyphenationState }) | undefined;
  const [state, setState] = useState<HyphenationState>(adapter?.reactHyphenationCatalog?.() || { auto: false, caps: true, zone: 18, limit: '' });
  const update = (next: Partial<HyphenationState>) => setState((current) => ({ ...current, ...next }));
  return <div className="col" id="react-wordx-hyphenation-body"><label><input id="react-wordx-hyphen-auto" type="checkbox" checked={state.auto} onChange={(event) => update({ auto: event.currentTarget.checked })} />{tr(L, 'Automatically hyphenate document')}</label><label><input id="react-wordx-hyphen-caps" type="checkbox" checked={state.caps} onChange={(event) => update({ caps: event.currentTarget.checked })} />{tr(L, 'Hyphenate words in CAPS')}</label><label htmlFor="react-wordx-hyphen-zone">{tr(L, 'Hyphenation zone:')}</label><input id="react-wordx-hyphen-zone" type="number" min={0} step={1} value={state.zone} onChange={(event) => update({ zone: Number(event.currentTarget.value) })} /><label htmlFor="react-wordx-hyphen-limit">{tr(L, 'Limit consecutive hyphens to:')}</label><input id="react-wordx-hyphen-limit" type="number" min={0} value={state.limit} placeholder={tr(L, 'No limit')} onChange={(event) => update({ limit: event.currentTarget.value === '' ? '' : Number(event.currentTarget.value) })} /></div>;
}

function PasteSpecialDialogBody({ L }: { L: WordxRuntime }) {
  const adapter = L.dlg as (typeof L.dlg & { reactPasteSpecialCatalog?(): { hasClipboard: boolean } }) | undefined;
  const [mode, setMode] = useState('paste');
  const [format, setFormat] = useState('rich');
  const hasClipboard = adapter?.reactPasteSpecialCatalog?.().hasClipboard || false;
  const formats: Array<[string, string]> = [['rich','Formatted Text (RTF)'],['html','HTML Format'],['text','Unformatted Text'],['utext','Unformatted Unicode Text']];
  return <div className="col" id="react-wordx-paste-special-body"><div className="tp-note">{tr(L, hasClipboard ? 'Source: Wordx document' : 'Source: (the Wordx clipboard is empty — copy something first, or press Ctrl+V)')}</div><fieldset><legend>{tr(L, 'Paste Special')}</legend><label><input type="radio" name="react-wordx-paste-special-mode" value="paste" checked={mode === 'paste'} onChange={() => setMode('paste')} />{tr(L, 'Paste:')}</label><label><input type="radio" name="react-wordx-paste-special-mode" value="link" checked={mode === 'link'} onChange={() => setMode('link')} />{tr(L, 'Paste link:')}</label></fieldset><label htmlFor="react-wordx-paste-special-format">{tr(L, 'Format:')}</label><select id="react-wordx-paste-special-format" size={4} value={format} onChange={(event) => setFormat(event.currentTarget.value)}>{formats.map(([value, label]) => <option key={value} value={value}>{tr(L, label)}</option>)}</select><div className="tp-note">{tr(L, 'Formatted text keeps fonts and formatting; unformatted text takes the formatting of the destination.')}</div></div>;
}

function MacrosDialogBody({ L }: { L: WordxRuntime }) {
  return <div className="col" id="react-wordx-macros-body"><label htmlFor="react-wordx-macro-name">{tr(L, 'Macro name:')}</label><input id="react-wordx-macro-name" type="text" style={{ width: '100%' }} disabled /><select aria-label={tr(L, 'Macros')} size={7} disabled /><div className="tp-note">{tr(L, 'Macros are not run in Wordx. VBA projects in .docm files are kept when the document is saved as .docm by Word, but Wordx saves the document text only.')}</div></div>;
}

type OptionsViewState = { options: Array<{ key: string; label: string; checked: boolean; layout: boolean }>; fieldShading: string };
function OptionsViewDialogBody({ L }: { L: WordxRuntime }) {
  const adapter = L.dlg as (typeof L.dlg & { reactOptionsViewCatalog?(): OptionsViewState; reactOptionsViewSet?(key: string, value: boolean | string): OptionsViewState | undefined }) | undefined;
  const [state, setState] = useState<OptionsViewState>(adapter?.reactOptionsViewCatalog?.() || { options: [], fieldShading: 'selected' });
  const change = (key: string, value: boolean | string) => { const next = adapter?.reactOptionsViewSet?.(key, value); if (next) setState(next); };
  const group = (title: string, options: OptionsViewState['options']) => <fieldset className="col"><legend>{tr(L, title)}</legend><div className="grid2">{options.map((option) => <label key={option.key}><input type="checkbox" data-option-key={option.key} checked={option.checked} onChange={(event) => change(option.key, event.currentTarget.checked)} />{tr(L, option.label)}</label>)}</div></fieldset>;
  return <div className="col" id="react-wordx-options-view-body">{group('Show', state.options.filter((item) => !['showMarks','showHidden','textBoundaries','vruler','whiteSpace'].includes(item.key)))}<label htmlFor="react-wordx-options-field-shading">{tr(L, 'Field shading:')}<select id="react-wordx-options-field-shading" value={state.fieldShading} onChange={(event) => change('fieldShading', event.currentTarget.value)}>{([['never','Never'],['always','Always'],['selected','When selected']] as const).map(([value,label])=><option key={value} value={value}>{tr(L,label)}</option>)}</select></label>{group('Formatting marks', state.options.filter((item) => ['showMarks','showHidden'].includes(item.key)))}{group('Print and Web Layout options', state.options.filter((item) => ['textBoundaries','vruler','whiteSpace'].includes(item.key)))}</div>;
}

type OptionsGeneralState = { options: Array<{ key: string; checked: boolean }>; unit: string; units: Array<[string,string]>; look: string; looks: Array<{ id: string; name: string }>; language: string; languages: Array<{ id: string; name: string }> };
function OptionsGeneralDialogBody({ L }: { L: WordxRuntime }) {
  const adapter = L.dlg as (typeof L.dlg & { reactOptionsGeneralCatalog?(): OptionsGeneralState; reactOptionsGeneralSet?(key: string, value: boolean | string): OptionsGeneralState }) | undefined;
  const [state, setState] = useState<OptionsGeneralState>(adapter?.reactOptionsGeneralCatalog?.() || { options: [], unit: 'in', units: [], look: 'classic', looks: [], language: 'auto', languages: [] });
  const change = (key: string, value: boolean | string) => { const next = adapter?.reactOptionsGeneralSet?.(key, value); if (next) setState(next); };
  return <div className="col" id="react-wordx-options-general-body">
    {state.looks.length > 0 && <label htmlFor="react-wordx-options-look">{tr(L,'Theme:')}<select id="react-wordx-options-look" value={state.look} onChange={(event)=>change('look',event.currentTarget.value)}>{state.looks.map((item)=><option key={item.id} value={item.id}>{tr(L,item.name)}</option>)}</select></label>}
    {state.languages.length > 0 && <label htmlFor="react-wordx-options-language">{tr(L,'Language:')}<select id="react-wordx-options-language" value={state.language} onChange={(event)=>change('language',event.currentTarget.value)}>{state.languages.map((item)=><option key={item.id} value={item.id}>{item.name}</option>)}</select></label>}
    {state.options.map((item)=>{const label={bgRepag:'Background repagination',blueBg:'Blue background, white text',sound:'Provide feedback with sound'}[item.key]||item.key;return <label key={item.key}><input type="checkbox" data-option-key={item.key} checked={item.checked} onChange={(event)=>change(item.key,event.currentTarget.checked)}/>{tr(L,label)}</label>;})}
    <label htmlFor="react-wordx-options-unit">{tr(L,'Measurement units:')}<select id="react-wordx-options-unit" value={state.unit} onChange={(event)=>change('unit',event.currentTarget.value)}>{state.units.map(([value,label])=><option key={value} value={value}>{tr(L,label)}</option>)}</select></label>
  </div>;
}

type OptionsEditItem = { key: string; label: string; checked: boolean };
function OptionsEditDialogBody({ L }: { L: WordxRuntime }) {
  const adapter = L.dlg as (typeof L.dlg & { reactOptionsEditCatalog?(): OptionsEditItem[]; reactOptionsEditSet?(key: string, value: boolean): OptionsEditItem[] }) | undefined;
  const [items, setItems] = useState<OptionsEditItem[]>(adapter?.reactOptionsEditCatalog?.() || []);
  const change = (key: string, value: boolean) => { const next = adapter?.reactOptionsEditSet?.(key, value); if (next) setItems(next); };
  const group = (title: string, values: OptionsEditItem[]) => <fieldset className="col"><legend>{tr(L,title)}</legend>{values.map((item)=><label key={item.key}><input type="checkbox" data-option-key={item.key} checked={item.checked} onChange={(event)=>change(item.key,event.currentTarget.checked)}/>{tr(L,item.label)}</label>)}</fieldset>;
  return <div className="col" id="react-wordx-options-edit-body">{group('Editing options',items.filter((item)=>item.key!=='pasteOptions'))}{group('Cut and paste options',items.filter((item)=>item.key==='pasteOptions'))}</div>;
}

type OptionsPrintState = { printing: OptionsEditItem[]; include: OptionsEditItem[] };
function OptionsPrintDialogBody({ L }: { L: WordxRuntime }) {
  const adapter = L.dlg as (typeof L.dlg & { reactOptionsPrintCatalog?(): OptionsPrintState; reactOptionsPrintSet?(key: string, value: boolean): OptionsPrintState }) | undefined;
  const [state, setState] = useState<OptionsPrintState>(adapter?.reactOptionsPrintCatalog?.() || { printing: [], include: [] });
  const change = (key: string, value: boolean) => { const next = adapter?.reactOptionsPrintSet?.(key, value); if (next) setState(next); };
  const group = (title: string, items: OptionsEditItem[]) => <fieldset className="col"><legend>{tr(L,title)}</legend>{items.map((item)=><label key={item.key}><input type="checkbox" data-option-key={item.key} checked={item.checked} onChange={(event)=>change(item.key,event.currentTarget.checked)}/>{tr(L,item.label)}</label>)}</fieldset>;
  return <div className="col" id="react-wordx-options-print-body">{group('Printing options',state.printing)}{group('Include with document',state.include)}</div>;
}

type OptionsSaveState = { promptProps: boolean; autoRecover: boolean; autoSaveMin: number; formats: Array<[string,string]> };
function OptionsSaveDialogBody({ L }: { L: WordxRuntime }) {
  const adapter = L.dlg as (typeof L.dlg & { reactOptionsSaveCatalog?(): OptionsSaveState; reactOptionsSaveSet?(key: string, value: boolean): OptionsSaveState }) | undefined;
  const [state, setState] = useState<OptionsSaveState>(adapter?.reactOptionsSaveCatalog?.() || { promptProps: false, autoRecover: true, autoSaveMin: 10, formats: [] });
  const change = (key: string, value: boolean) => { const next = adapter?.reactOptionsSaveSet?.(key, value); if (next) setState(next); };
  return <div className="col" id="react-wordx-options-save-body"><label><input type="checkbox" data-option-key="promptProps" checked={state.promptProps} onChange={(event)=>change('promptProps',event.currentTarget.checked)}/>{tr(L,'Prompt for document properties')}</label><div className="row"><label><input type="checkbox" data-option-key="autoRecover" checked={state.autoRecover} onChange={(event)=>change('autoRecover',event.currentTarget.checked)}/>{tr(L,'Save AutoRecover info every:')}</label><input id="react-wordx-options-autorecover-min" type="number" min={1} max={120} value={state.autoSaveMin} onChange={(event)=>setState({...state,autoSaveMin:Number(event.currentTarget.value)})}/><span>{tr(L,'minutes')}</span></div><div className="tp-note">{tr(L,'AutoRecover keeps a copy in this browser only; it is offered when Wordx opens after an unexpected close.')}</div><label htmlFor="react-wordx-options-save-format">{tr(L,'Save Word files as:')}<select id="react-wordx-options-save-format" defaultValue="docx">{state.formats.map(([value,label])=><option key={value} value={value}>{tr(L,label)}</option>)}</select></label></div>;
}

type OptionsTrackState = { color: string; colors: Array<[string,string]>; balloons: string; balloonOptions: Array<[string,string]> };
function OptionsTrackDialogBody({ L }: { L: WordxRuntime }) {
  const adapter = L.dlg as (typeof L.dlg & { reactOptionsTrackCatalog?(): OptionsTrackState; reactOptionsTrackSet?(key: string, value: string): OptionsTrackState }) | undefined;
  const [state, setState] = useState<OptionsTrackState>(adapter?.reactOptionsTrackCatalog?.() || { color: 'author', colors: [], balloons: 'comments', balloonOptions: [] });
  const change = (key: string, value: string) => { const next = adapter?.reactOptionsTrackSet?.(key, value); if (next) setState(next); };
  return <div className="col" id="react-wordx-options-track-body"><fieldset className="col"><legend>{tr(L,'Markup')}</legend><label htmlFor="react-wordx-options-insertion-color">{tr(L,'Insertions color:')}<select id="react-wordx-options-insertion-color" value={state.color} onChange={(event)=>change('color',event.currentTarget.value)}>{state.colors.map(([value,label])=><option key={value} value={value}>{tr(L,label)}</option>)}</select></label></fieldset><fieldset className="col"><legend>{tr(L,'Balloons')}</legend><label htmlFor="react-wordx-options-balloons">{tr(L,'Use Balloons (Print and Web Layout):')}<select id="react-wordx-options-balloons" value={state.balloons} onChange={(event)=>change('balloons',event.currentTarget.value)}>{state.balloonOptions.map(([value,label])=><option key={value} value={value}>{tr(L,label)}</option>)}</select></label></fieldset></div>;
}

type OptionsUserState = { name: string; initials: string; address: string };
function OptionsUserDialogBody({ L }: { L: WordxRuntime }) {
  const adapter = L.dlg as (typeof L.dlg & { reactOptionsUserCatalog?(): OptionsUserState }) | undefined;
  const state = adapter?.reactOptionsUserCatalog?.() || { name: '', initials: '', address: '' };
  return <div className="col" id="react-wordx-options-user-body"><label htmlFor="react-wordx-options-user-name">{tr(L,'Name:')}<input id="react-wordx-options-user-name" type="text" defaultValue={state.name} /></label><label htmlFor="react-wordx-options-user-initials">{tr(L,'Initials:')}<input id="react-wordx-options-user-initials" type="text" defaultValue={state.initials} /></label><label htmlFor="react-wordx-options-user-address">{tr(L,'Mailing address:')}<textarea id="react-wordx-options-user-address" rows={3} defaultValue={state.address} /></label></div>;
}

type OptionsCompatState = { version: string; versions: Array<[string,string]> };
function OptionsCompatDialogBody({ L }: { L: WordxRuntime }) {
  const adapter = L.dlg as (typeof L.dlg & { reactOptionsCompatCatalog?(): OptionsCompatState; reactOptionsCompatSet?(value: string): OptionsCompatState }) | undefined;
  const [state, setState] = useState<OptionsCompatState>(adapter?.reactOptionsCompatCatalog?.() || { version: '11', versions: [] });
  const change = (value: string) => { const next = adapter?.reactOptionsCompatSet?.(value); if (next) setState(next); };
  return <div className="col" id="react-wordx-options-compat-body"><label htmlFor="react-wordx-options-compat-version">{tr(L,'Lay out this document as if created in:')}<select id="react-wordx-options-compat-version" value={state.version} onChange={(event)=>change(event.currentTarget.value)}>{state.versions.map(([value,label])=><option key={value} value={value}>{tr(L,label)}</option>)}</select></label><div className="tp-note">{tr(L,'Word 2003 layout positions tables by their cell margins and keeps legacy spacing rules; later versions align tables with the text.')}</div></div>;
}

type OptionsSecurityState = { removePersonal: boolean; warnTracked: boolean; password: string };
function OptionsSecurityDialogBody({ L }: { L: WordxRuntime }) {
  const adapter = L.dlg as (typeof L.dlg & { reactOptionsSecurityCatalog?(): OptionsSecurityState; reactOptionsSecuritySet?(key: string, value: boolean): OptionsSecurityState }) | undefined;
  const [state, setState] = useState<OptionsSecurityState>(adapter?.reactOptionsSecurityCatalog?.() || { removePersonal: false, warnTracked: false, password: '' });
  const change = (key: string, value: boolean) => { const next = adapter?.reactOptionsSecuritySet?.(key, value); if (next) setState(next); };
  return <div className="col" id="react-wordx-options-security-body"><fieldset className="col"><legend>{tr(L,'Privacy options')}</legend><label><input type="checkbox" data-option-key="removePersonal" checked={state.removePersonal} onChange={(event)=>change('removePersonal',event.currentTarget.checked)}/>{tr(L,'Remove personal information from file properties on save')}</label><label><input type="checkbox" data-option-key="warnTracked" checked={state.warnTracked} onChange={(event)=>change('warnTracked',event.currentTarget.checked)}/>{tr(L,'Warn before printing, saving or sending a file that contains tracked changes or comments')}</label></fieldset><fieldset className="col"><legend>{tr(L,'File encryption options for this document')}</legend><label htmlFor="react-wordx-options-password">{tr(L,'Password to open:')}<input id="react-wordx-options-password" type="password" autoComplete="new-password" defaultValue={state.password} /></label><div className="tp-note">{tr(L,'Saved .docx files are encrypted with AES-256, as Word 2010 and later do. Passwords are case-sensitive; if you forget it, the document cannot be recovered.')}</div></fieldset><div className="tp-note">{tr(L,'To restrict editing instead, use Tools ▸ Protect Document.')}</div></div>;
}

function ConfirmPasswordDialogBody({ L }: { L: WordxRuntime }) {
  return <div className="col" id="react-wordx-confirm-password-body"><label htmlFor="react-wordx-confirm-password">{tr(L,'Reenter password to open:')}<input id="react-wordx-confirm-password" type="password" autoComplete="new-password" /></label><div className="tp-note">{tr(L,'Caution: If you lose or forget the password, it cannot be recovered.')}</div></div>;
}

type OptionsSpellingState = { spelling: OptionsEditItem[]; grammar: OptionsEditItem[]; hidden: boolean };
function OptionsSpellingDialogBody({ L }: { L: WordxRuntime }) {
  const adapter = L.dlg as (typeof L.dlg & { reactOptionsSpellingCatalog?(): OptionsSpellingState; reactOptionsSpellingSet?(key: string, value: boolean): OptionsSpellingState; reactOptionsSpellingAction?(action: string): void }) | undefined;
  const [state, setState] = useState<OptionsSpellingState>(adapter?.reactOptionsSpellingCatalog?.() || { spelling: [], grammar: [], hidden: false });
  const change = (key: string, value: boolean) => { const next = adapter?.reactOptionsSpellingSet?.(key, value); if (next) setState(next); };
  const group = (title: string, items: OptionsEditItem[]) => <fieldset className="col"><legend>{tr(L,title)}</legend>{items.map((item)=><label key={item.key}><input type="checkbox" data-option-key={item.key} checked={item.checked} onChange={(event)=>change(item.key,event.currentTarget.checked)}/>{tr(L,item.label)}</label>)}{title==='Spelling'&&<label><input type="checkbox" data-option-key="hidden" checked={state.hidden} onChange={(event)=>change('hidden',event.currentTarget.checked)}/>{tr(L,'Hide spelling and grammar errors in this document')}</label>}{title==='Spelling'&&<button className="btn small" type="button" onClick={()=>adapter?.reactOptionsSpellingAction?.('dictionary')}>{tr(L,'Custom Dictionaries...')}</button>}{title==='Grammar'&&<button className="btn small" type="button" onClick={()=>adapter?.reactOptionsSpellingAction?.('readability')}>{tr(L,'Readability Statistics...')}</button>}</fieldset>;
  return <div className="col" id="react-wordx-options-spelling-body">{group('Spelling',state.spelling)}{group('Grammar',state.grammar)}<button className="btn small" type="button" onClick={()=>adapter?.reactOptionsSpellingAction?.('recheck')}>{tr(L,'Recheck Document')}</button></div>;
}

function CustomDictionaryDialogBody({ L }: { L: WordxRuntime }) {
  const adapter=L.dlg as (typeof L.dlg & {reactCustomDictionaryCatalog?():{words:string}})|undefined;
  const state=adapter?.reactCustomDictionaryCatalog?.()||{words:''};
  return <div className="col" id="react-wordx-custom-dictionary-body"><label htmlFor="react-wordx-custom-dictionary-words">{tr(L,'Words (one per line):')}</label><textarea id="react-wordx-custom-dictionary-words" rows={12} defaultValue={state.words}/></div>;
}

function FormatObjectWebDialogBody({ L }: { L: WordxRuntime }) {
  const adapter = L.dlg as (typeof L.dlg & { reactFormatObjectWebCatalog?(): { alt: string } }) | undefined;
  const state = adapter?.reactFormatObjectWebCatalog?.() || { alt: '' };
  return <div className="col" id="react-wordx-format-object-web-body"><label htmlFor="react-wordx-format-object-alt">{tr(L,'Alternative text:')}<textarea id="react-wordx-format-object-alt" rows={4} defaultValue={state.alt} /></label><div className="tp-note">{tr(L,'Web browsers display alternative text while pictures are loading or if they are missing. Screen readers read it aloud.')}</div></div>;
}

type FormatObjectColorsState = { fill: string | null; fillType: string; line: string | null; transparency: number; dash: string; dashes: Array<{value:string;label:string}>; weight: number; head: string; tail: string; arrows: Array<{value:string;label:string}>; showFill: boolean; showArrows: boolean };
function FormatObjectColorsDialogBody({ L }: { L: WordxRuntime }) {
  const adapter=L.dlg as (typeof L.dlg & {reactFormatObjectColorsCatalog?():FormatObjectColorsState;reactFormatObjectColorsSet?(key:string,value:string|number):FormatObjectColorsState;reactFormatObjectColorsAction?(action:string,anchor:HTMLElement):void})|undefined;
  const [state,setState]=useState<FormatObjectColorsState>(adapter?.reactFormatObjectColorsCatalog?.()||{fill:null,fillType:'none',line:null,transparency:0,dash:'solid',dashes:[],weight:.75,head:'none',tail:'none',arrows:[],showFill:true,showArrows:false});
  useEffect(()=>L.bus?.on('wordx-format-object-colors-refresh',()=>{const next=adapter?.reactFormatObjectColorsCatalog?.();if(next)setState(next);}),[L,adapter]);
  const update=(key:string,value:string|number)=>{const next=adapter?.reactFormatObjectColorsSet?.(key,value);if(next)setState(next);};
  const colorButton=(kind:'fill'|'line',color:string|null)=> <button type="button" className="btn small" aria-label={tr(L,'Color:')} onClick={(event)=>adapter?.reactFormatObjectColorsAction?.(kind,event.currentTarget)}><span className="cpk-sw" style={{background:color||'transparent'}}/>{tr(L,color?'Color':kind==='fill'?'No Fill':'No Line')}</button>;
  return <div className="col" id="react-wordx-format-object-colors-body">
    {state.showFill&&<fieldset className="col"><legend>{tr(L,'Fill')}</legend><div className="row">{colorButton('fill',state.fill)}<button type="button" className="btn small" onClick={(event)=>adapter?.reactFormatObjectColorsAction?.('effects',event.currentTarget)}>{tr(L,'Fill Effects...')}</button></div>{state.fillType!=='solid'&&state.fillType!=='none'&&<div className="tp-note">{tr(L,'Current fill:')} {tr(L,({grad:'Gradient',patt:'Pattern',img:'Picture'} as Record<string,string>)[state.fillType]||state.fillType)}</div>}<label>{tr(L,'Transparency:')}<input type="number" min={0} max={100} step={5} value={state.transparency} onChange={(event)=>update('transparency',Number(event.currentTarget.value))}/>%</label></fieldset>}
    <fieldset className="col"><legend>{tr(L,'Line')}</legend><div className="row">{colorButton('line',state.line)}</div><label>{tr(L,'Dashed:')}<select value={state.dash} onChange={(event)=>update('dash',event.currentTarget.value)}>{state.dashes.map((item)=><option key={item.value} value={item.value}>{tr(L,item.label)}</option>)}</select></label><label>{tr(L,'Weight:')}<input type="number" min={0} max={1584} step={0.25} value={state.weight} onChange={(event)=>update('weight',Number(event.currentTarget.value))}/> pt</label></fieldset>
    {state.showArrows&&<fieldset className="col"><legend>{tr(L,'Arrows')}</legend><label>{tr(L,'Begin style:')}<select value={state.head} onChange={(event)=>update('head',event.currentTarget.value)}>{state.arrows.map((item)=><option key={item.value} value={item.value}>{tr(L,item.label)}</option>)}</select></label><label>{tr(L,'End style:')}<select value={state.tail} onChange={(event)=>update('tail',event.currentTarget.value)}>{state.arrows.map((item)=><option key={item.value} value={item.value}>{tr(L,item.label)}</option>)}</select></label></fieldset>}
  </div>;
}

type FillEffectsState = { tab:string;colors:string;color1:string;color2:string;preset:string;style:string;pattern:string;foreground:string;background:string;picture:string|null;pictureName:string;presets:string[];styles:Array<[string,string]>;patterns:Array<{key:string;name:string;svg:string}> };
function FillEffectsDialogBody({ L }: { L: WordxRuntime }) {
  const adapter=L.dlg as (typeof L.dlg & {reactFillEffectsCatalog?():FillEffectsState;reactFillEffectsSet?(key:string,value:string):FillEffectsState;reactFillEffectsAction?(action:string,anchor?:HTMLElement):Promise<FillEffectsState>|FillEffectsState})|undefined;
  const [state,setState]=useState<FillEffectsState>(adapter?.reactFillEffectsCatalog?.()||{tab:'gradient',colors:'two',color1:'#3366FF',color2:'#FFFFFF',preset:'Early Sunset',style:'h',pattern:'pct50',foreground:'#000000',background:'#FFFFFF',picture:null,pictureName:'No picture selected',presets:[],styles:[],patterns:[]});
  useEffect(()=>L.bus?.on('wordx-fill-effects-refresh',()=>{const next=adapter?.reactFillEffectsCatalog?.();if(next)setState(next);}),[L,adapter]);
  const update=(key:string,value:string)=>{const next=adapter?.reactFillEffectsSet?.(key,value);if(next)setState(next);};
  const pick=(key:'color1'|'color2'|'foreground'|'background')=><button type="button" className="btn small" data-fill-color={key} onClick={(event)=>adapter?.reactFillEffectsAction?.(key,event.currentTarget)}><span className="cpk-sw" style={{background:state[key]}}/>{tr(L,state[key])}</button>;
  const tabs:Array<[string,string]>=[['gradient','Gradient'],['texture','Texture'],['pattern','Pattern'],['picture','Picture']];
  return <div className="col" id="react-wordx-fill-effects-body"><div className="tabs" role="tablist">{tabs.map(([key,label])=><button type="button" role="tab" aria-selected={state.tab===key} className={state.tab===key?'on':''} key={key} onClick={()=>update('tab',key)}>{tr(L,label)}</button>)}</div>
    {state.tab==='gradient'&&<div className="col" role="tabpanel"><fieldset className="col"><legend>{tr(L,'Colors')}</legend>{([['one','One color'],['two','Two colors'],['preset','Preset']] as const).map(([value,label])=><label key={value}><input type="radio" name="react-wordx-fill-effects-colors" value={value} checked={state.colors===value} onChange={()=>update('colors',value)}/>{tr(L,label)}</label>)}{state.colors!=='preset'?<div className="row">{pick('color1')}{state.colors==='two'&&pick('color2')}</div>:<label>{tr(L,'Preset colors:')}<select value={state.preset} onChange={(event)=>update('preset',event.currentTarget.value)}>{state.presets.map((name)=><option key={name}>{tr(L,name)}</option>)}</select></label>}</fieldset><fieldset className="col"><legend>{tr(L,'Shading styles')}</legend><div className="row">{state.styles.map(([value,label])=><button type="button" className={`btn small${state.style===value?' on':''}`} aria-pressed={state.style===value} key={value} onClick={()=>update('style',value)}>{tr(L,label)}</button>)}</div></fieldset></div>}
    {state.tab==='texture'&&<div className="tp-note" role="tabpanel">{tr(L,'Textures are drawn as patterns; choose a pattern or a picture.')}</div>}
    {state.tab==='pattern'&&<div className="col" role="tabpanel"><div className="shape-grid" style={{gridTemplateColumns:'repeat(8,26px)'}}>{state.patterns.map((item)=><button type="button" key={item.key} className={state.pattern===item.key?'on':''} aria-label={tr(L,item.name)} title={tr(L,item.name)} onClick={()=>update('pattern',item.key)} dangerouslySetInnerHTML={{__html:item.svg}}/>)}</div><div className="row">{pick('foreground')}{pick('background')}</div></div>}
    {state.tab==='picture'&&<div className="col" role="tabpanel"><button type="button" className="btn small" onClick={()=>adapter?.reactFillEffectsAction?.('picture')}>{tr(L,'Select Picture...')}</button><div className="tp-note">{tr(L,state.pictureName)}</div></div>}
  </div>;
}

type FormatObjectSizeState = { width: number; height: number; rotation: number; locked: boolean; naturalWidth: number; naturalHeight: number; hasNatural: boolean };
function FormatObjectSizeDialogBody({ L }: { L: WordxRuntime }) {
  const adapter = L.dlg as (typeof L.dlg & { reactFormatObjectSizeCatalog?(): FormatObjectSizeState }) | undefined;
  const [state,setState] = useState<FormatObjectSizeState>(adapter?.reactFormatObjectSizeCatalog?.() || {width:72,height:72,rotation:0,locked:false,naturalWidth:72,naturalHeight:72,hasNatural:false});
  const ratio=state.height/state.width;
  const update=(key:'width'|'height'|'rotation'|'locked'|'scaleWidth'|'scaleHeight',value:number|boolean)=>setState((current)=>{
    if(key==='locked')return {...current,locked:!!value};
    if(key==='rotation')return {...current,rotation:Number(value)||0};
    if(key==='width'){const width=Math.max(1,Number(value)||1);return {...current,width,height:current.locked?width*ratio:current.height};}
    if(key==='height'){const height=Math.max(1,Number(value)||1);return {...current,height,width:current.locked?height/ratio:current.width};}
    const percent=Math.max(1,Number(value)||1)/100;
    return key==='scaleWidth'?{...current,width:current.naturalWidth*percent,height:current.locked?current.naturalHeight*percent:current.height}:{...current,height:current.naturalHeight*percent,width:current.locked?current.naturalWidth*percent:current.width};
  });
  const scaleWidth=state.width/state.naturalWidth*100,scaleHeight=state.height/state.naturalHeight*100;
  return <div className="col" id="react-wordx-format-object-size-body"><fieldset className="col"><legend>{tr(L,'Size and rotate')}</legend><div className="row"><label>{tr(L,'Height:')}<input type="number" min={1} step={0.1} data-size-field="height" value={state.height} onChange={(e)=>update('height',Number(e.currentTarget.value))}/></label><label>{tr(L,'Width:')}<input type="number" min={1} step={0.1} data-size-field="width" value={state.width} onChange={(e)=>update('width',Number(e.currentTarget.value))}/></label></div><label>{tr(L,'Rotation:')}<input type="number" min={-3600} max={3600} step={0.1} data-size-field="rotation" value={state.rotation} onChange={(e)=>update('rotation',Number(e.currentTarget.value))}/>°</label></fieldset><fieldset className="col"><legend>{tr(L,'Scale')}</legend><div className="row"><label>{tr(L,'Height:')}<input type="number" min={1} max={5000} data-size-field="scale-height" value={Math.round(scaleHeight)} onChange={(e)=>update('scaleHeight',Number(e.currentTarget.value))}/>%</label><label>{tr(L,'Width:')}<input type="number" min={1} max={5000} data-size-field="scale-width" value={Math.round(scaleWidth)} onChange={(e)=>update('scaleWidth',Number(e.currentTarget.value))}/>%</label></div><label><input type="checkbox" data-size-field="locked" checked={state.locked} onChange={(e)=>update('locked',e.currentTarget.checked)}/>{tr(L,'Lock aspect ratio')}</label></fieldset>{state.hasNatural&&<div className="tp-note">{L.t?.('Original size — Height: {0}, Width: {1}',state.naturalHeight,state.naturalWidth)}</div>}<div className="tp-btns"><button type="button" className="btn small" onClick={()=>setState((current)=>({...current,width:current.hasNatural?current.naturalWidth:current.width,height:current.hasNatural?current.naturalHeight:current.height,rotation:0}))}>{tr(L,'Reset')}</button></div></div>;
}

type FormatObjectLayoutState = { wrap: string; alignment: string; wraps: Array<{ id: string; label: string }> };
function FormatObjectLayoutDialogBody({ L }: { L: WordxRuntime }) {
  const adapter=L.dlg as (typeof L.dlg & { reactFormatObjectLayoutCatalog?():FormatObjectLayoutState;reactFormatObjectLayoutSet?(key:string,value:string):FormatObjectLayoutState })|undefined;
  const [state,setState]=useState<FormatObjectLayoutState>(adapter?.reactFormatObjectLayoutCatalog?.()||{wrap:'inline',alignment:'other',wraps:[]});
  useEffect(()=>L.bus?.on('wordx-format-object-layout-refresh',()=>{const next=adapter?.reactFormatObjectLayoutCatalog?.();if(next)setState(next);}),[L,adapter]);
  const update=(key:string,value:string)=>{const next=adapter?.reactFormatObjectLayoutSet?.(key,value);if(next)setState(next);};
  const diagram=(wrap:string)=>{const lengths=({inline:[8,36,36],square:[10,10,10],tight:[12,9,12],behind:[36,36,36],none:[36,36,36]} as Record<string,number[]>)[wrap]||[36,36,36];return <svg width="40" height="36" viewBox="0 0 40 36" aria-hidden="true"><path d={`M2 6h36M2 12h${lengths[0]}M2 18h${lengths[1]}M2 24h${lengths[2]}M2 30h36`} stroke="#999"/><rect x="12" y="8" width="16" height="18" fill="#9cf" stroke="#333"/></svg>;};
  return <div className="col" id="react-wordx-format-object-layout-body"><fieldset className="col"><legend>{tr(L,'Wrapping style')}</legend><div className="preset-grid">{state.wraps.map((item)=><button type="button" key={item.id} className={`preset${state.wrap===item.id?' on':''}`} data-wrap={item.id} aria-pressed={state.wrap===item.id} onClick={()=>update('wrap',item.id)}>{diagram(item.id)}<span>{tr(L,item.label)}</span></button>)}</div></fieldset><fieldset className="col"><legend>{tr(L,'Horizontal alignment')}</legend><div className="row">{([['left','Left'],['center','Center'],['right','Right'],['other','Other']] as const).map(([value,label])=><button type="button" key={value} className={`btn small${state.alignment===value?' on':''}`} aria-pressed={state.alignment===value} onClick={()=>update('alignment',value)}>{tr(L,label)}</button>)}</div></fieldset><div className="tp-btns"><button type="button" className="btn small" onClick={()=>adapter?.reactFormatObjectLayoutSet?.('advanced','')}>{tr(L,'Advanced...')}</button></div></div>;
}

type FormatObjectAdvancedState = { tab:string;hMode:string;hAlign:string;hRel:string;hOff:number;vMode:string;vAlign:string;vRel:string;vOff:number;moveWith:boolean;overlap:boolean;top:number;bottom:number;left:number;right:number;hRels:Array<[string,string]>;vRels:Array<[string,string]> };
function FormatObjectAdvancedDialogBody({ L }: { L: WordxRuntime }) {
  const adapter=L.dlg as (typeof L.dlg & {reactFormatObjectAdvancedCatalog?():FormatObjectAdvancedState;reactFormatObjectAdvancedSet?(key:string,value:string|number|boolean):FormatObjectAdvancedState})|undefined;
  const [state,setState]=useState<FormatObjectAdvancedState>(adapter?.reactFormatObjectAdvancedCatalog?.()||{tab:'position',hMode:'abs',hAlign:'left',hRel:'column',hOff:0,vMode:'abs',vAlign:'top',vRel:'paragraph',vOff:0,moveWith:true,overlap:true,top:0,bottom:0,left:9,right:9,hRels:[],vRels:[]});
  const update=(key:string,value:string|number|boolean)=>{const next=adapter?.reactFormatObjectAdvancedSet?.(key,value);if(next)setState(next);};
  const mode=(key:'hMode'|'vMode',value:string,label:string)=><label key={`${key}-${value}`}><input type="radio" name={`react-wordx-format-object-${key}`} value={value} checked={state[key]===value} onChange={()=>update(key,value)}/>{tr(L,label)}</label>;
  const tabs:Array<[string,string]>=[['position','Picture &Position'],['wrapping','Text Wrapping']];
  return <div className="col" id="react-wordx-format-object-advanced-body"><div className="tabs" role="tablist">{tabs.map(([value,label])=><button type="button" role="tab" aria-selected={state.tab===value} className={state.tab===value?'on':''} key={value} onClick={()=>update('tab',value)}>{tr(L,label)}</button>)}</div>
    {state.tab==='position'?<div className="col" role="tabpanel"><fieldset className="col"><legend>{tr(L,'Horizontal')}</legend>{mode('hMode','align','Alignment')}{mode('hMode','abs','Absolute position')}<div className="row"><select aria-label={tr(L,'Horizontal alignment')} value={state.hAlign} disabled={state.hMode!=='align'} onChange={(event)=>update('hAlign',event.currentTarget.value)}>{([['left','Left'],['center','Centered'],['right','Right']] as const).map(([value,label])=><option key={value} value={value}>{tr(L,label)}</option>)}</select><span>{tr(L,'relative to')}</span><select aria-label={tr(L,'Horizontal relative position')} value={state.hRel} onChange={(event)=>update('hRel',event.currentTarget.value)}>{state.hRels.map(([value,label])=><option key={value} value={value}>{tr(L,label)}</option>)}</select></div>{state.hMode==='abs'&&<label>{tr(L,'to the right of')}<input type="number" min={0} step={0.1} value={state.hOff} onChange={(event)=>update('hOff',Number(event.currentTarget.value))}/></label>}</fieldset><fieldset className="col"><legend>{tr(L,'Vertical')}</legend>{mode('vMode','align','Alignment')}{mode('vMode','abs','Absolute position')}<div className="row"><select aria-label={tr(L,'Vertical alignment')} value={state.vAlign} disabled={state.vMode!=='align'} onChange={(event)=>update('vAlign',event.currentTarget.value)}>{([['top','Top'],['center','Centered'],['bottom','Bottom']] as const).map(([value,label])=><option key={value} value={value}>{tr(L,label)}</option>)}</select><span>{tr(L,'relative to')}</span><select aria-label={tr(L,'Vertical relative position')} value={state.vRel} onChange={(event)=>update('vRel',event.currentTarget.value)}>{state.vRels.map(([value,label])=><option key={value} value={value}>{tr(L,label)}</option>)}</select></div>{state.vMode==='abs'&&<label>{tr(L,'below')}<input type="number" min={0} step={0.1} value={state.vOff} onChange={(event)=>update('vOff',Number(event.currentTarget.value))}/></label>}</fieldset><fieldset className="col"><legend>{tr(L,'Options')}</legend><label><input type="checkbox" checked={state.moveWith} onChange={(event)=>update('moveWith',event.currentTarget.checked)}/>{tr(L,'Move object with text')}</label><label><input type="checkbox" checked={state.overlap} onChange={(event)=>update('overlap',event.currentTarget.checked)}/>{tr(L,'Allow overlap')}</label></fieldset></div>:<fieldset className="col" role="tabpanel"><legend>{tr(L,'Distance from text')}</legend>{([['top','Top'],['left','Left'],['bottom','Bottom'],['right','Right']] as const).map(([key,label])=><label key={key}>{tr(L,label)}<input type="number" min={0} step={0.1} value={state[key]} onChange={(event)=>update(key,Number(event.currentTarget.value))}/></label>)}</fieldset>}
  </div>;
}

type FormatObjectTextBoxState = { margins: {left:number;right:number;top:number;bottom:number}; anchor: string; autofit: boolean };
function FormatObjectTextBoxDialogBody({ L }: { L: WordxRuntime }) {
  const adapter=L.dlg as (typeof L.dlg & {reactFormatObjectTextBoxCatalog?():FormatObjectTextBoxState})|undefined;
  const state=adapter?.reactFormatObjectTextBoxCatalog?.()||{margins:{left:7.2,right:7.2,top:3.6,bottom:3.6},anchor:'t',autofit:false};
  return <div className="col" id="react-wordx-format-object-textbox-body"><fieldset className="col"><legend>{tr(L,'Internal margin')}</legend><div className="row">{([['left','Left'],['top','Top'],['right','Right'],['bottom','Bottom']] as const).map(([key,label])=><label key={key}>{tr(L,label)}<input type="number" min={0} step={0.1} data-box-margin={key} defaultValue={state.margins[key]}/></label>)}</div></fieldset><label htmlFor="react-wordx-format-object-anchor">{tr(L,'Vertical alignment:')}<select id="react-wordx-format-object-anchor" defaultValue={state.anchor}>{([['t','Top'],['ctr','Center'],['b','Bottom']] as const).map(([value,label])=><option key={value} value={value}>{tr(L,label)}</option>)}</select></label><label><input type="checkbox" data-box-autofit defaultChecked={state.autofit}/>{tr(L,'Resize AutoShape to fit text')}</label></div>;
}

type FormatObjectPictureState={crop:{left:number;right:number;top:number;bottom:number};mode:string;brightness:number;contrast:number;naturalWidth:number;naturalHeight:number};
function FormatObjectPictureDialogBody({ L }: { L: WordxRuntime }) {
  const adapter=L.dlg as (typeof L.dlg & {reactFormatObjectPictureCatalog?():FormatObjectPictureState})|undefined;
  const state=adapter?.reactFormatObjectPictureCatalog?.()||{crop:{left:0,right:0,top:0,bottom:0},mode:'',brightness:50,contrast:50,naturalWidth:72,naturalHeight:72};
  const [mode,setMode]=useState(state.mode),[brightness,setBrightness]=useState(state.brightness),[contrast,setContrast]=useState(state.contrast);
  return <div className="col" id="react-wordx-format-object-picture-body"><fieldset className="col"><legend>{tr(L,'Crop from')}</legend><div className="row">{([['left','Left'],['top','Top'],['right','Right'],['bottom','Bottom']] as const).map(([key,label])=><label key={key}>{tr(L,label)}<input type="number" min={0} step={0.1} data-picture-crop={key} defaultValue={state.crop[key]}/></label>)}</div></fieldset><fieldset className="col"><legend>{tr(L,'Image control')}</legend><label htmlFor="react-wordx-format-object-picture-mode">{tr(L,'Color:')}<select id="react-wordx-format-object-picture-mode" value={mode} onChange={(e)=>setMode(e.currentTarget.value)}>{([['','Automatic'],['gray','Grayscale'],['bw','Black & White'],['washout','Washout']] as const).map(([value,label])=><option key={value||'auto'} value={value}>{tr(L,label)}</option>)}</select></label><label>{tr(L,'Brightness:')}<input type="number" min={0} max={100} step={3} data-picture-tone="brightness" value={brightness} onChange={(e)=>setBrightness(Number(e.currentTarget.value))}/>%</label><label>{tr(L,'Contrast:')}<input type="number" min={0} max={100} step={3} data-picture-tone="contrast" value={contrast} onChange={(e)=>setContrast(Number(e.currentTarget.value))}/>%</label></fieldset></div>;
}

function AutoCorrectExceptionsDialogBody({ L }: { L: WordxRuntime }) {
  const adapter = L.dlg as (typeof L.dlg & { reactAutoCorrectExceptionsCatalog?(): { words: string }; reactAutoCorrectExceptionsAction?(value: string): { words: string } }) | undefined;
  const [words, setWords] = useState(adapter?.reactAutoCorrectExceptionsCatalog?.().words || '');
  return <div className="col" id="react-wordx-autocorrect-exceptions-body"><label htmlFor="react-wordx-autocorrect-exceptions">{tr(L, "Don't capitalize after:")}</label><textarea id="react-wordx-autocorrect-exceptions" rows={10} value={words} onChange={(event) => setWords(event.currentTarget.value)} style={{ width: '100%' }} /></div>;
}

type AutoCorrectMainState={options:Record<string,boolean>;entries:Array<[string,string]>;selected:string};
function AutoCorrectMainDialogBody({ L }: { L: WordxRuntime }) {
  const adapter=L.dlg as (typeof L.dlg & {reactAutoCorrectMainCatalog?():AutoCorrectMainState;reactAutoCorrectMainAction?(action:string,value?:any):AutoCorrectMainState})|undefined;
  const [state,setState]=useState<AutoCorrectMainState>(adapter?.reactAutoCorrectMainCatalog?.()||{options:{},entries:[],selected:''});
  const [from,setFrom]=useState(state.selected),[to,setTo]=useState(state.entries.find(([key])=>key===state.selected)?.[1]||'');
  const update=(action:string,value?:unknown)=>{const next=adapter?.reactAutoCorrectMainAction?.(action,value);if(next)setState(next);};
  const select=(key:string)=>{setFrom(key);setTo(state.entries.find(([entry])=>entry===key)?.[1]||'');update('select',{from:key});};
  const options:Array<[string,string]>=[['acButtons','Show AutoCorrect Options buttons'],['twoCaps','Correct TWo INitial CApitals'],['capSentence','Capitalize first letter of sentences'],['capCells','Capitalize first letter of table cells'],['capDays','Capitalize names of days'],['capsLock','Correct accidental usage of cAPS LOCK key'],['autocorrect','Replace text as you type']];
  return <div className="col" id="react-wordx-autocorrect-main-body">{options.map(([key,label])=><label key={key}><input type="checkbox" data-autocorrect-main={key} checked={state.options[key]!==false} onChange={(event)=>update('option',{key,enabled:event.currentTarget.checked})}/>{tr(L,label)}</label>)}<button type="button" className="btn small" onClick={()=>update('exceptions')}>{tr(L,'Exceptions...')}</button><div className="row"><label>{tr(L,'Replace:')}<input id="react-wordx-ac-from" value={from} onChange={(event)=>setFrom(event.currentTarget.value)}/></label><label>{tr(L,'With:')}<input id="react-wordx-ac-to" value={to} onChange={(event)=>setTo(event.currentTarget.value)}/></label></div><select aria-label={tr(L,'AutoCorrect replacements')} size={7} value={state.selected} onChange={(event)=>select(event.currentTarget.value)}>{state.entries.map(([key,value])=><option key={key} value={key}>{key}  →  {value}</option>)}</select><div className="tp-btns"><button type="button" className="btn small" onClick={()=>update('entry',{from,to})}>{tr(L,'Add')}</button><button type="button" className="btn small" disabled={!from} onClick={()=>{update('delete',{from});setFrom('');setTo('');}}>{tr(L,'Delete')}</button></div></div>;
}

function AutoCorrectFormatDialogBody({ L }: { L: WordxRuntime }) {
  const adapter=L.dlg as (typeof L.dlg & {reactAutoCorrectFormatCatalog?():{options:Record<string,boolean>};reactAutoCorrectFormatAction?(key:string,enabled:boolean):{options:Record<string,boolean>}})|undefined;
  const [options,setOptions]=useState(adapter?.reactAutoCorrectFormatCatalog?.().options||{});
  const groups:Array<[string,Array<[string,string]>]>=[['Replace as you type',[['smartQuotes','"Straight quotes" with “smart quotes”'],['ordinals','Ordinals (1st) with superscript'],['fractions','Fractions (1/2) with fraction character (½)'],['dashes','Hyphens (--) with dash (—)'],['autoFormat','*Bold* and _italic_ with real formatting'],['autoUrl','Internet and network paths with hyperlinks']]],['Apply as you type',[['autoBullets','Automatic bulleted lists'],['autoBorders','Border lines'],['autoTables','Tables'],['autoNumbers','Automatic numbered lists'],['autoHeadings','Built-in Heading styles']]]];
  return <div className="col" id="react-wordx-autocorrect-format-body">{groups.map(([title,items])=><fieldset className="col" key={title}><legend>{tr(L,title)}</legend>{items.map(([key,label])=><label key={key}><input type="checkbox" data-autocorrect-format={key} checked={!!options[key]} onChange={(event)=>{const next=adapter?.reactAutoCorrectFormatAction?.(key,event.currentTarget.checked);if(next)setOptions(next.options);}}/>{tr(L,label)}</label>)}</fieldset>)}</div>;
}

type AutoTextState={entries:Array<{name:string;text:string}>;selected:string;name:string;preview:string;autoComplete:boolean};
function AutoTextDialogBody({ L }: { L: WordxRuntime }) {
  const adapter=L.dlg as (typeof L.dlg & {reactAutoTextCatalog?():AutoTextState;reactAutoTextAction?(action:string,value?:string):AutoTextState;reactAutoCorrectMainAction?(action:string,value?:any):unknown})|undefined;
  const catalog=adapter?.reactAutoTextCatalog?.()||{entries:[],selected:'',name:'',preview:'',autoComplete:true};
  const [state,setState]=useState<AutoTextState>(catalog),[name,setName]=useState(catalog.name);
  const act=(action:string,value?:string)=>{const next=adapter?.reactAutoTextAction?.(action,value);if(next){setState(next);setName(next.name);}};
  return <div className="col" id="react-wordx-autotext-body"><label><input type="checkbox" checked={catalog.autoComplete!==false} onChange={(event)=>adapter?.reactAutoCorrectMainAction?.('option',{key:'autoComplete',enabled:event.currentTarget.checked})}/>{tr(L,'Show AutoComplete suggestions')}</label><label htmlFor="react-wordx-autotext-name">{tr(L,'Enter AutoText entries here:')}</label><input id="react-wordx-autotext-name" value={name} onChange={(event)=>setName(event.currentTarget.value)}/><select aria-label={tr(L,'AutoText entries')} size={7} value={state.selected} onChange={(event)=>act('select',event.currentTarget.value)}>{state.entries.map((entry)=><option key={entry.name} value={entry.name}>{entry.name}</option>)}</select><div className="cd-cap">{tr(L,'Preview')}</div><div className="tp-box" style={{minHeight:40,whiteSpace:'pre-wrap'}} translate="no">{state.preview}</div><div className="tp-btns"><button type="button" className="btn small" onClick={()=>act('add',name)}>{tr(L,'Add')}</button><button type="button" className="btn small" disabled={!state.selected} onClick={()=>act('delete',state.selected)}>{tr(L,'Delete')}</button><button type="button" className="btn small" disabled={!state.selected} onClick={()=>act('insert',state.selected)}>{tr(L,'Insert')}</button></div></div>;
}

type FontState={font?:string;sz?:number;b?:boolean;i?:boolean;u?:string;uColor?:string;color?:string;strike?:boolean;dstrike?:boolean;vert?:string;shadow?:boolean;outline?:boolean;emboss?:boolean;imprint?:boolean;smallCaps?:boolean;caps?:boolean;hidden?:boolean;w?:number;spacing?:number;position?:number;kern?:number;effect?:string;tab:number;target:boolean;fonts:string[];sizes:string[];underlines:Array<[string,string]>;animations:Array<[string,string]>;effects:Array<[string,string]>;sample:string;caption:string;[key:string]:any};
function FontDialogBody({ L }: { L: WordxRuntime }) {
  const adapter=L.dlg as (typeof L.dlg & {reactFontCatalog?():FontState;reactFontAction?(action:string,value?:any):FontState})|undefined;
  const [state,setState]=useState<FontState>(adapter?.reactFontCatalog?.()||{tab:0,target:false,fonts:[],sizes:[],underlines:[],animations:[],effects:[],sample:'Sample Text',caption:''});
  const [tab,setTab]=useState(state.tab);
  const act=(action:string,value?:unknown)=>{const next=adapter?.reactFontAction?.(action,value);if(next)setState(next);};
  const set=(key:string,value:unknown)=>act('set',{key,value});
  const color=(key:'color'|'uColor',label:string)=><button type="button" className="btn small" onClick={(event)=>L.ui?.colorMenu?.(event.currentTarget,{mode:'font'},(value)=>set(key,typeof value==='object'&&value!==null&&'auto'in value?undefined:value))}>{tr(L,label)}<span className="cpk-sw" style={{background:state[key]?'#'+state[key]:'#000'}}/></button>;
  const styles=state.b&&state.i?'Bold Italic':state.b?'Bold':state.i?'Italic':'Regular';
  const effects=(key:string)=>key==='sup'?state.vert==='superscript':key==='sub'?state.vert==='subscript':!!state[key];
  const changeEffect=(key:string,enabled:boolean)=>set(key,enabled);
  const spacingKind=!state.spacing?'normal':state.spacing>0?'exp':'cond',positionKind=!state.position?'normal':state.position>0?'raised':'lowered';
  return <div className="col" id="react-wordx-font-body"><div className="tabs-head" role="tablist">{(['Font','Character Spacing','Text Effects'] as const).map((label,index)=><button type="button" role="tab" aria-selected={tab===index} className={tab===index?'on':''} key={label} onClick={()=>setTab(index)}>{tr(L,label)}</button>)}</div>
    {tab===0?<div className="col" role="tabpanel"><div className="row"><label style={{flex:2}}>{tr(L,'Font:')}<input list="react-wordx-font-list" value={state.font||''} onChange={(event)=>set('font',event.currentTarget.value)}/><datalist id="react-wordx-font-list">{state.fonts.map((font)=><option key={font} value={font}/>)}</datalist></label><label style={{flex:1}}>{tr(L,'Font style:')}<select value={styles} onChange={(event)=>set('style',event.currentTarget.value)}>{['Regular','Italic','Bold','Bold Italic'].map(value=><option key={value}>{tr(L,value)}</option>)}</select></label><label style={{width:80}}>{tr(L,'Size:')}<input type="number" min={1} max={1638} step={0.5} value={state.sz||12} onChange={(event)=>set('sz',Number(event.currentTarget.value))}/></label></div><div className="row">{color('color','Font color:')}<label>{tr(L,'Underline style:')}<select value={state.u||'none'} onChange={(event)=>set('u',event.currentTarget.value==='none'?(state.target?undefined:'none'):event.currentTarget.value)}>{state.underlines.map(([value,label])=><option key={value} value={value}>{tr(L,label)}</option>)}</select></label>{color('uColor','Underline color:')}</div><fieldset className="grid3"><legend>{tr(L,'Effects')}</legend>{state.effects.map(([key,label])=><label key={key}><input type="checkbox" checked={effects(key)} onChange={(event)=>changeEffect(key,event.currentTarget.checked)}/>{tr(L,label)}</label>)}</fieldset></div>:tab===1?<div className="col" role="tabpanel"><label>{tr(L,'Scale:')}<select value={state.w||100} onChange={(event)=>set('w',Number(event.currentTarget.value)===100?undefined:Number(event.currentTarget.value))}>{[200,150,100,90,80,66,50,33].map(value=><option key={value} value={value}>{value}%</option>)}</select></label><div className="row"><label>{tr(L,'Spacing:')}<select value={spacingKind} onChange={(event)=>{const amount=Math.abs(state.spacing||1);set('spacing',event.currentTarget.value==='normal'?undefined:event.currentTarget.value==='cond'?-amount:amount)}}>{([['normal','Normal'],['exp','Expanded'],['cond','Condensed']] as const).map(([value,label])=><option key={value} value={value}>{tr(L,label)}</option>)}</select></label><label>{tr(L,'By:')}<input type="number" min={0} max={1584} step={0.1} value={Math.abs(state.spacing||0)} onChange={(event)=>set('spacing',(spacingKind==='cond'?-1:1)*Number(event.currentTarget.value))}/></label></div><div className="row"><label>{tr(L,'Position:')}<select value={positionKind} onChange={(event)=>{const amount=Math.abs(state.position||3);set('position',event.currentTarget.value==='normal'?undefined:event.currentTarget.value==='lowered'?-amount:amount)}}>{([['normal','Normal'],['raised','Raised'],['lowered','Lowered']] as const).map(([value,label])=><option key={value} value={value}>{tr(L,label)}</option>)}</select></label><label>{tr(L,'By:')}<input type="number" min={0} max={1584} step={1} value={Math.abs(state.position||0)} onChange={(event)=>set('position',(positionKind==='lowered'?-1:1)*Number(event.currentTarget.value))}/></label></div><div className="row"><label><input type="checkbox" checked={!!state.kern} onChange={(event)=>set('kern',event.currentTarget.checked?state.kern||14:undefined)}/>{tr(L,'Kerning for fonts:')}</label><label>{tr(L,'Points and above')}<input type="number" min={1} value={state.kern||14} onChange={(event)=>set('kern',Number(event.currentTarget.value))}/></label></div></div>:<div className="col" role="tabpanel"><label>{tr(L,'Animations:')}<select size={7} value={state.effect||''} onChange={(event)=>set('effect',event.currentTarget.value||undefined)}>{state.animations.map(([value,label])=><option key={value} value={value}>{tr(L,label)}</option>)}</select></label><div className="tp-note">{tr(L,'Animated text effects are kept in the file; they display as plain text, as in later versions of Word.')}</div></div>}
    <div className="cd-cap">{tr(L,'Preview')}</div><div className="preview-box font-preview" translate="no"><span style={{fontFamily:state.font||'Times New Roman',fontSize:`${state.sz||12}pt`,fontWeight:state.b?'bold':'normal',fontStyle:state.i?'italic':'normal',color:state.color||'inherit',textDecoration:state.u&&state.u!=='none'?'underline':undefined}}>{state.sample}</span></div><div className="prev-cap">{state.caption}</div></div>;
}

type CustomizeState = { toolbars: Array<{ id: string; label: string; shown: boolean }>; commands: Array<{ id: string; label: string; key: string }>; shortcuts: Array<{ key: string; label: string }>; screenTips: boolean };
function CustomizeDialogBody({ L }: { L: WordxRuntime }) {
  const adapter = L.dlg as (typeof L.dlg & { reactCustomizeCatalog?(): CustomizeState; reactCustomizeAction?(action: string, value?: { id?: string; enabled?: boolean }): CustomizeState }) | undefined;
  const [state, setState] = useState<CustomizeState>(adapter?.reactCustomizeCatalog?.() || { toolbars: [], commands: [], shortcuts: [], screenTips: true });
  const [tab, setTab] = useState('toolbars');
  const [command, setCommand] = useState('');
  const refresh = (action: string, value?: { id?: string; enabled?: boolean }) => { const next = adapter?.reactCustomizeAction?.(action, value); if (next) setState(next); };
  return <div className="col" id="react-wordx-customize-body">
    <div role="tablist" className="tabs-head">{([['toolbars','Toolbars'],['commands','Commands'],['options','Options'],['keyboard','Keyboard']] as const).map(([id,label])=><button key={id} type="button" role="tab" aria-selected={tab===id} className={`tab${tab===id?' on':''}`} onClick={()=>setTab(id)}>{tr(L,label)}</button>)}</div>
    {tab==='toolbars'&&<div className="col" role="tabpanel">{state.toolbars.map((item)=><label key={item.id}><input type="checkbox" checked={item.shown} onChange={()=>refresh('toolbar',{id:item.id})}/>{tr(L,item.label.replace(/&/g,''))}</label>)}</div>}
    {tab==='commands'&&<div className="col" role="tabpanel"><div className="tp-note">{tr(L,'All Wordx commands. Double-click one to run it.')}</div><select aria-label={tr(L,'Commands')} size={12} value={command} onChange={(event)=>setCommand(event.currentTarget.value)} onDoubleClick={()=>command&&refresh('command',{id:command})}>{state.commands.map((item)=><option key={item.id} value={item.id}>{tr(L,item.label)}{item.key?`   ${item.key}`:''}</option>)}</select></div>}
    {tab==='options'&&<div className="col" role="tabpanel"><label><input type="checkbox" checked={state.screenTips} onChange={(event)=>refresh('screenTips',{enabled:event.currentTarget.checked})}/>{tr(L,'Show ScreenTips on toolbars')}</label><label><input type="checkbox" disabled defaultChecked/>{tr(L,'Always show full menus')}</label><label><input type="checkbox" disabled defaultChecked/>{tr(L,'Show shortcut keys in ScreenTips')}</label></div>}
    {tab==='keyboard'&&<div role="tabpanel" style={{maxHeight:300,overflow:'auto'}}><table className="kbd-table"><tbody>{state.shortcuts.map((item,index)=><tr key={`${item.key}-${index}`}><td>{item.key}</td><td>{tr(L,item.label)}</td></tr>)}</tbody></table></div>}
  </div>;
}

type DrawingGridState = { snap: boolean; horizontal: number; vertical: number };
function DrawingGridDialogBody({ L }: { L: WordxRuntime }) {
  const adapter = L.dlg as (typeof L.dlg & { reactDrawingGridCatalog?(): DrawingGridState; reactDrawingGridAction?(value: DrawingGridState): DrawingGridState }) | undefined;
  const state = adapter?.reactDrawingGridCatalog?.() || { snap: false, horizontal: 9, vertical: 9 };
  return <div className="col" id="react-wordx-drawing-grid-body"><label><input name="react-wordx-grid-snap" type="checkbox" defaultChecked={state.snap}/>{tr(L,'Snap objects to grid')}</label><fieldset className="col"><legend>{tr(L,'Grid settings')}</legend><label htmlFor="react-wordx-grid-horizontal">{tr(L,'Horizontal spacing:')}<input id="react-wordx-grid-horizontal" type="number" min={1} step={0.1} defaultValue={state.horizontal}/></label><label htmlFor="react-wordx-grid-vertical">{tr(L,'Vertical spacing:')}<input id="react-wordx-grid-vertical" type="number" min={1} step={0.1} defaultValue={state.vertical}/></label></fieldset></div>;
}

type AutoFormatState = { mode: string; kind: string; options: Record<string, boolean> };
function AutoFormatDialogBody({ L }: { L: WordxRuntime }) {
  const adapter = L.dlg as (typeof L.dlg & { reactAutoFormatCatalog?(): AutoFormatState }) | undefined;
  const state = adapter?.reactAutoFormatCatalog?.() || { mode: 'now', kind: 'general', options: {} };
  const [mode, setMode] = useState(state.mode);
  const [kind, setKind] = useState(state.kind);
  const [options, setOptions] = useState(state.options);
  const labels: Record<string,string> = { headings:'Built-in Heading styles', lists:'List styles', quotes:'Straight quotes with smart quotes', dashes:'Hyphens (--) with dash (—)', ordinals:'Ordinals (1st) with superscript', fractions:'Fractions (1/2) with fraction character', urls:'Internet paths with hyperlinks', empty:'Remove extra empty paragraphs' };
  return <div className="col" id="react-wordx-autoformat-body"><div className="tp-note">{tr(L,'Wordx will automatically format the document.')}</div><fieldset className="col"><legend>{tr(L,'Mode')}</legend><label><input type="radio" name="react-wordx-autoformat-mode" value="now" checked={mode==='now'} onChange={()=>setMode('now')}/>{tr(L,'AutoFormat now')}</label><label><input type="radio" name="react-wordx-autoformat-mode" value="review" checked={mode==='review'} onChange={()=>setMode('review')}/>{tr(L,'AutoFormat and review each change')}</label></fieldset><label htmlFor="react-wordx-autoformat-kind">{tr(L,'Document type')}<select id="react-wordx-autoformat-kind" value={kind} onChange={(event)=>setKind(event.currentTarget.value)}><option value="general">{tr(L,'General document')}</option><option value="letter">{tr(L,'Letter')}</option><option value="email">{tr(L,'E-mail')}</option></select></label><fieldset className="col"><legend>{tr(L,'Apply / replace')}</legend>{Object.entries(labels).map(([id,label])=><label key={id}><input type="checkbox" data-autoformat-option={id} checked={!!options[id]} onChange={(event)=>setOptions({...options,[id]:event.currentTarget.checked})}/>{tr(L,label)}</label>)}</fieldset></div>;
}

function AutoFormatReviewDialogBody({ L }: { L: WordxRuntime }) {
  const adapter=L.dlg as (typeof L.dlg & {reactAutoFormatReviewCatalog?():{changes:number}})|undefined;
  const changes=adapter?.reactAutoFormatReviewCatalog?.().changes||0;
  const message=changes===1?L.t?.('Formatting completed: {0} change. You can accept all of them, or reject all and return the document to its previous state.',changes):L.t?.('Formatting completed: {0} changes. You can accept all of them, or reject all and return the document to its previous state.',changes);
  return <div className="tp-note" id="react-wordx-autoformat-review-body">{message}</div>;
}

type EnvelopeState = { tab: number; selectedText: string; address: string; returnAddress: string; omit: boolean; size: string; labelAddress: string; products: Array<{ id: string; name: string }>; product: string; printMode: string; row: number; column: number };
function EnvelopesDialogBody({ L }: { L: WordxRuntime }) {
  const adapter = L.dlg as (typeof L.dlg & { reactEnvelopeCatalog?(): EnvelopeState; reactEnvelopeTab?(tab: string): void }) | undefined;
  const state = adapter?.reactEnvelopeCatalog?.() || { tab: 0, selectedText: '', address: '', returnAddress: '', omit: false, size: '10', labelAddress: '', products: [], product: '5160', printMode: 'full', row: 1, column: 1 };
  const [tab, setTab] = useState(state.tab === 1 ? 'labels' : 'envelopes');
  const [address, setAddress] = useState(state.address || state.selectedText);
  const [returnAddress, setReturnAddress] = useState(state.returnAddress);
  const [omit, setOmit] = useState(state.omit);
  const [size, setSize] = useState(state.size);
  const [labelAddress, setLabelAddress] = useState(state.labelAddress || state.selectedText);
  const [product, setProduct] = useState(state.product);
  const [printMode, setPrintMode] = useState(state.printMode);
  const [row, setRow] = useState(state.row);
  const [column, setColumn] = useState(state.column);
  return <div className="col" id="react-wordx-envelopes-body"><div className="tabs-head" role="tablist">{([['envelopes','Envelopes'],['labels','Labels']] as const).map(([id,label])=><button type="button" key={id} role="tab" aria-selected={tab===id} className={`tab${tab===id?' on':''}`} onClick={()=>{setTab(id);adapter?.reactEnvelopeTab?.(id);}}>{tr(L,label)}</button>)}</div>{tab==='envelopes'?<div className="col" role="tabpanel"><label htmlFor="react-wordx-envelope-address">{tr(L,'Delivery address:')}</label><textarea id="react-wordx-envelope-address" rows={5} value={address} onChange={(event)=>setAddress(event.currentTarget.value)}/><label htmlFor="react-wordx-envelope-return">{tr(L,'Return address:')}<input id="react-wordx-envelope-omit" type="checkbox" checked={omit} onChange={(event)=>setOmit(event.currentTarget.checked)}/>{tr(L,'Omit')}</label><textarea id="react-wordx-envelope-return" rows={4} value={returnAddress} onChange={(event)=>setReturnAddress(event.currentTarget.value)}/><label htmlFor="react-wordx-envelope-size">{tr(L,'Envelope size:')}<select id="react-wordx-envelope-size" value={size} onChange={(event)=>setSize(event.currentTarget.value)}><option value="10">{tr(L,'Size 10 (4 1/8 x 9 1/2 in)')}</option><option value="DL">{tr(L,'DL (110 x 220 mm)')}</option><option value="C5">{tr(L,'C5 (162 x 229 mm)')}</option><option value="Monarch">{tr(L,'Monarch (3 7/8 x 7 1/2 in)')}</option></select></label></div>:<div className="col" role="tabpanel"><label htmlFor="react-wordx-label-address">{tr(L,'Address:')}<textarea id="react-wordx-label-address" rows={5} value={labelAddress} onChange={(event)=>setLabelAddress(event.currentTarget.value)}/></label><label htmlFor="react-wordx-label-product">{tr(L,'Label product:')}<select id="react-wordx-label-product" value={product} onChange={(event)=>setProduct(event.currentTarget.value)}>{state.products.map((item)=><option key={item.id} value={item.id}>{item.name}</option>)}</select></label><fieldset className="col"><legend>{tr(L,'Print')}</legend><label><input type="radio" name="react-wordx-label-mode" value="full" checked={printMode==='full'} onChange={()=>setPrintMode('full')}/>{tr(L,'Full page of the same label')}</label><label><input type="radio" name="react-wordx-label-mode" value="single" checked={printMode==='single'} onChange={()=>setPrintMode('single')}/>{tr(L,'Single label')}</label><label htmlFor="react-wordx-label-row">{tr(L,'Row:')}<input id="react-wordx-label-row" type="number" min={1} max={30} value={row} onChange={(event)=>setRow(Number(event.currentTarget.value))}/></label><label htmlFor="react-wordx-label-column">{tr(L,'Column:')}<input id="react-wordx-label-column" type="number" min={1} max={4} value={column} onChange={(event)=>setColumn(Number(event.currentTarget.value))}/></label></fieldset></div>}</div>;
}

type IndexTablesState = { initialTab: string; styles: Array<{ id: string; name: string; level?: number }>; labels: string[]; leaders: Array<[string,string]> };
function IndexTablesDialogBody({ L }: { L: WordxRuntime }) {
  const adapter = L.dlg as (typeof L.dlg & { reactIndexTablesCatalog?(): IndexTablesState; reactIndexTablesAction?(action: string, value?: unknown): void }) | undefined;
  const catalog = adapter?.reactIndexTablesCatalog?.() || { initialTab:'toc', styles:[], labels:['Figure','Table','Equation'], leaders:[] };
  const [tab,setTab]=useState(catalog.initialTab),[options,setOptions]=useState(false);
  const [tocShow,setTocShow]=useState(true),[tocRight,setTocRight]=useState(true),[tocLinks,setTocLinks]=useState(true),[tocLeader,setTocLeader]=useState('dot'),[tocLevels,setTocLevels]=useState(3),[tocFormat,setTocFormat]=useState('template');
  const [tocStyles,setTocStyles]=useState<Record<string,number>>(()=>Object.fromEntries(catalog.styles.filter((item)=>/^Heading\d$/.test(item.id)&&item.level!=null&&item.level<3).map((item)=>[item.id,(item.level||0)+1])));
  const [outline,setOutline]=useState(true),[fields,setFields]=useState(false);
  const [figure,setFigure]=useState(catalog.labels[0]||'Figure'),[include,setInclude]=useState(true),[figureShow,setFigureShow]=useState(true);
  const [indexType,setIndexType]=useState('indent'),[columns,setColumns]=useState(2),[indexRight,setIndexRight]=useState(false),[headings,setHeadings]=useState(false);
  const tabs=[['index','Index'],['toc','Table of Contents'],['tof','Table of Figures'],['toa','Table of Authorities']] as const;
  return <div className="col" id="react-wordx-index-tables-body"><div className="tabs-head" role="tablist">{tabs.map(([id,label])=><button key={id} type="button" role="tab" data-index-tab={id} aria-selected={tab===id} className={`tab${tab===id?' on':''}`} onClick={()=>setTab(id)}>{tr(L,label)}</button>)}</div>
    {tab==='toc'&&<div className="row" role="tabpanel"><div className="col" style={{flex:1}}><div className="cd-cap">{tr(L,'Print Preview')}</div><div className="preview-box" style={{height:150,flexDirection:'column',alignItems:'stretch',padding:6}}>{Array.from({length:Math.min(3,tocLevels)},(_,i)=><div key={i} style={{display:'flex',paddingLeft:i*12,fontWeight:i===0?'bold':undefined}}><span>{tr(L,'Heading {0}'.replace('{0}',String(i+1)))}</span><span style={{flex:1,borderBottom:tocRight&&tocLeader!=='none'?'1px dotted #000':undefined,margin:'0 2px 3px'}}/><span>{tocShow?String(i*2+1):''}</span></div>)}</div><label><input id="react-wordx-toc-show" type="checkbox" checked={tocShow} onChange={(e)=>setTocShow(e.currentTarget.checked)}/>{tr(L,'Show page numbers')}</label><label><input id="react-wordx-toc-right" type="checkbox" checked={tocRight} onChange={(e)=>setTocRight(e.currentTarget.checked)}/>{tr(L,'Right align page numbers')}</label><label>{tr(L,'Tab leader:')}<select id="react-wordx-toc-leader" value={tocLeader} onChange={(e)=>setTocLeader(e.currentTarget.value)}>{catalog.leaders.map(([v,l])=><option key={v} value={v}>{tr(L,l)}</option>)}</select></label></div><div className="col" style={{flex:1}}><div className="cd-cap">{tr(L,'Web Preview')}</div><div className="preview-box" style={{height:150,flexDirection:'column',alignItems:'flex-start',padding:6,gap:3}}>{[1,2,3].map((n)=><u key={n} style={{color:'#00f',marginLeft:(n-1)*10}}>{tr(L,'Heading {0}'.replace('{0}',String(n)))}</u>)}</div><label><input id="react-wordx-toc-links" type="checkbox" checked={tocLinks} onChange={(e)=>setTocLinks(e.currentTarget.checked)}/>{tr(L,'Use hyperlinks instead of page numbers (Web)')}</label></div><div className="col" style={{flexBasis:'100%'}}><label>{tr(L,'Formats:')}<select id="react-wordx-toc-format" value={tocFormat} onChange={(e)=>setTocFormat(e.currentTarget.value)}>{([['template','From template'],['classic','Classic'],['distinctive','Distinctive'],['formal','Formal'],['simple','Simple']] as const).map(([v,l])=><option key={v} value={v}>{tr(L,l)}</option>)}</select></label><label>{tr(L,'Show levels:')}<input id="react-wordx-toc-levels" type="number" min={1} max={9} value={tocLevels} onChange={(e)=>setTocLevels(Number(e.currentTarget.value))}/></label><button type="button" className="btn small" onClick={()=>setOptions(true)}>{tr(L,'Options...')}</button><button type="button" className="btn small" onClick={()=>adapter?.reactIndexTablesAction?.('modify')}>{tr(L,'Modify...')}</button>{options&&<fieldset className="col"><legend>{tr(L,'Table of Contents Options')}</legend><div style={{maxHeight:200,overflow:'auto'}}>{catalog.styles.map((style)=><label key={style.id}>{style.name}<input type="number" min={1} max={9} data-toc-style={style.id} value={tocStyles[style.id]??''} onChange={(e)=>{const next={...tocStyles};const value=Number(e.currentTarget.value);if(value)next[style.id]=value;else delete next[style.id];setTocStyles(next)}}/></label>)}</div><label><input id="react-wordx-toc-outline" type="checkbox" checked={outline} onChange={(e)=>setOutline(e.currentTarget.checked)}/>{tr(L,'Outline levels')}</label><label><input id="react-wordx-toc-fields" type="checkbox" checked={fields} onChange={(e)=>setFields(e.currentTarget.checked)}/>{tr(L,'Table entry fields')}</label></fieldset>}</div></div>}
    {tab==='tof'&&<div className="col" role="tabpanel"><label>{tr(L,'Caption label:')}<select id="react-wordx-tof-label" value={figure} onChange={(e)=>setFigure(e.currentTarget.value)}>{catalog.labels.map((label)=><option key={label}>{label}</option>)}</select></label><label><input id="react-wordx-tof-include" type="checkbox" checked={include} onChange={(e)=>setInclude(e.currentTarget.checked)}/>{tr(L,'Include label and number')}</label><label><input id="react-wordx-tof-show" type="checkbox" checked={figureShow} onChange={(e)=>setFigureShow(e.currentTarget.checked)}/>{tr(L,'Show page numbers')}</label><div className="tp-note">{tr(L,'Captions inserted with Insert ▸ Reference ▸ Caption are collected here.')}</div></div>}
    {tab==='index'&&<div className="col" role="tabpanel"><fieldset className="col"><legend>{tr(L,'Type')}</legend><label><input type="radio" name="react-wordx-index-type" value="indent" checked={indexType==='indent'} onChange={()=>setIndexType('indent')}/>{tr(L,'Indented')}</label><label><input type="radio" name="react-wordx-index-type" value="runin" checked={indexType==='runin'} onChange={()=>setIndexType('runin')}/>{tr(L,'Run-in')}</label></fieldset><label>{tr(L,'Columns:')}<input id="react-wordx-index-columns" type="number" min={1} max={4} value={columns} onChange={(e)=>setColumns(Number(e.currentTarget.value))}/></label><label><input id="react-wordx-index-right" type="checkbox" checked={indexRight} onChange={(e)=>setIndexRight(e.currentTarget.checked)}/>{tr(L,'Right align page numbers')}</label><label><select aria-label={tr(L,'Tab leader:')} value={tocLeader} onChange={(e)=>setTocLeader(e.currentTarget.value)}>{catalog.leaders.map(([v,l])=><option key={v} value={v}>{tr(L,l)}</option>)}</select></label><label><input id="react-wordx-index-headings" type="checkbox" checked={headings} onChange={(e)=>setHeadings(e.currentTarget.checked)}/>{tr(L,'Show letter headings')}</label><button type="button" className="btn small" onClick={()=>{setTab('index');adapter?.reactIndexTablesAction?.('mark')}}>{tr(L,'Mark Entry...')}</button><div className="tp-note">{tr(L,'Mark entries first (Alt+Shift+X), then insert the index where it should appear.')}</div></div>}
    {tab==='toa'&&<div className="tp-note" role="tabpanel">{tr(L,'Tables of authorities need citations marked with TA fields. Use Insert ▸ Field ▸ TA to mark citations, then Field ▸ TOA to build the table.')}</div>}
  </div>;
}

type LetterWizardState = { dates: Array<{ value: string; label: string }>; styles: Array<[string,string]>; salutations: string[]; closings: string[]; name: string; address: string };
function LetterWizardDialogBody({ L }: { L: WordxRuntime }) {
  const adapter=L.dlg as (typeof L.dlg & {reactLetterWizardCatalog?():LetterWizardState})|undefined;
  const catalog=adapter?.reactLetterWizardCatalog?.()||{dates:[],styles:[],salutations:['Dear '],closings:['Sincerely,'],name:'',address:''};
  const [tab,setTab]=useState('format'),[date,setDate]=useState(true),[datePic,setDatePic]=useState(catalog.dates[0]?.value||'MMMM d, yyyy'),[style,setStyle]=useState('full');
  const [recipient,setRecipient]=useState(''),[delivery,setDelivery]=useState(''),[salutation,setSalutation]=useState('Dear '),[reference,setReference]=useState(''),[subject,setSubject]=useState(''),[cc,setCc]=useState('');
  const [sender,setSender]=useState(catalog.name),[returnAddress,setReturnAddress]=useState(catalog.address),[closing,setClosing]=useState('Sincerely,'),[title,setTitle]=useState(''),[enclosures,setEnclosures]=useState(0);
  return <div className="col" id="react-wordx-letter-wizard-body"><div className="tabs-head" role="tablist">{([['format','Letter Format'],['recipient','Recipient Info'],['other','Other Elements'],['sender','Sender Info']] as const).map(([id,label])=><button key={id} type="button" role="tab" aria-selected={tab===id} className={`tab${tab===id?' on':''}`} onClick={()=>setTab(id)}>{tr(L,label)}</button>)}</div>
    {tab==='format'&&<div className="col" role="tabpanel"><label><input id="react-wordx-letter-date" type="checkbox" checked={date} onChange={(event)=>setDate(event.currentTarget.checked)}/>{tr(L,'Date line:')}</label><label>{tr(L,'Date format:')}<select id="react-wordx-letter-date-format" value={datePic} onChange={(event)=>setDatePic(event.currentTarget.value)}>{catalog.dates.map((item)=><option key={item.value} value={item.value}>{item.label}</option>)}</select></label><label>{tr(L,'Page design / Letter style:')}<select id="react-wordx-letter-style" value={style} onChange={(event)=>setStyle(event.currentTarget.value)}>{catalog.styles.map(([value,label])=><option key={value} value={value}>{tr(L,label)}</option>)}</select></label></div>}
    {tab==='recipient'&&<div className="col" role="tabpanel"><label>{tr(L,"Recipient's name:")}<input id="react-wordx-letter-recipient" value={recipient} onChange={(event)=>setRecipient(event.currentTarget.value)}/></label><label>{tr(L,'Delivery address:')}<textarea id="react-wordx-letter-delivery" rows={4} value={delivery} onChange={(event)=>setDelivery(event.currentTarget.value)}/></label><label>{tr(L,'Salutation:')}<input id="react-wordx-letter-salutation" list="react-wordx-letter-salutations" value={salutation} onChange={(event)=>setSalutation(event.currentTarget.value)}/><datalist id="react-wordx-letter-salutations">{catalog.salutations.map((value)=><option key={value} value={value}/>)}</datalist></label></div>}
    {tab==='other'&&<div className="col" role="tabpanel"><label>{tr(L,'Reference line:')}<input id="react-wordx-letter-reference" placeholder={tr(L,'In reply to:')} value={reference} onChange={(event)=>setReference(event.currentTarget.value)}/></label><label>{tr(L,'Subject:')}<input id="react-wordx-letter-subject" value={subject} onChange={(event)=>setSubject(event.currentTarget.value)}/></label><label>{tr(L,'Courtesy copies (cc):')}<input id="react-wordx-letter-cc" value={cc} onChange={(event)=>setCc(event.currentTarget.value)}/></label></div>}
    {tab==='sender'&&<div className="col" role="tabpanel"><label>{tr(L,"Sender's name:")}<input id="react-wordx-letter-sender" value={sender} onChange={(event)=>setSender(event.currentTarget.value)}/></label><label>{tr(L,'Return address:')}<textarea id="react-wordx-letter-return" rows={3} value={returnAddress} onChange={(event)=>setReturnAddress(event.currentTarget.value)}/></label><label>{tr(L,'Complimentary closing:')}<input id="react-wordx-letter-closing" list="react-wordx-letter-closings" value={closing} onChange={(event)=>setClosing(event.currentTarget.value)}/><datalist id="react-wordx-letter-closings">{catalog.closings.map((value)=><option key={value} value={value}/>)}</datalist></label><label>{tr(L,'Job title:')}<input id="react-wordx-letter-title" value={title} onChange={(event)=>setTitle(event.currentTarget.value)}/></label><label>{tr(L,'Enclosures:')}<input id="react-wordx-letter-enclosures" type="number" min={0} value={enclosures} onChange={(event)=>setEnclosures(Number(event.currentTarget.value))}/></label></div>}
  </div>;
}

type SaveAsCatalog = { name: string; type: string; types: Array<{ id: string; label: string }>; note: string };
function SaveAsDialogBody({ L }: { L: WordxRuntime }) {
  const adapter = L.dlg as (typeof L.dlg & { reactSaveAsCatalog?(): SaveAsCatalog }) | undefined;
  const catalog = adapter?.reactSaveAsCatalog?.() || { name: 'Document', type: 'docx', types: [], note: '' };
  const [name, setName] = useState(catalog.name);
  const [type, setType] = useState(catalog.type);
  return <div className="col" id="react-wordx-save-as-body">
    <label htmlFor="react-wordx-save-as-name">{tr(L, 'File name:')}<input id="react-wordx-save-as-name" type="text" value={name} onChange={(event) => setName(event.currentTarget.value)} style={{ width: '100%' }} /></label>
    <label htmlFor="react-wordx-save-as-type">{tr(L, 'Save as type:')}<select id="react-wordx-save-as-type" value={type} onChange={(event) => setType(event.currentTarget.value)}>{catalog.types.map((item) => <option key={item.id} value={item.id}>{item.label}</option>)}</select></label>
    <div className="tp-note">{tr(L, catalog.note)}</div>
  </div>;
}

function PrintDialogBody({ L }: { L: WordxRuntime }) {
  const [range, setRange] = useState('all');
  const [pages, setPages] = useState('');
  const [what, setWhat] = useState('doc');
  const [which, setWhich] = useState('all');
  const [copies, setCopies] = useState(1);
  const [perSheet, setPerSheet] = useState('1');
  const ranges: Array<[string,string]> = [['all','All'],['current','Current page'],['selection','Selection'],['pages','Pages:']];
  const kinds: Array<[string,string]> = [['doc','Document'],['markup','Document showing markup'],['props','Document properties'],['list','List of markup'],['styles','Styles'],['autotext','AutoText entries'],['keys','Key assignments']];
  return <div className="col" id="react-wordx-print-body">
    <fieldset className="col"><legend>{tr(L,'Printer')}</legend><label htmlFor="react-wordx-print-printer">{tr(L,'Name:')}<select id="react-wordx-print-printer" disabled><option>{tr(L,'Save as PDF (Wordx)')}</option></select></label><div className="tp-note">{tr(L,'Status: Ready — the pages are saved as a PDF that you can print from any viewer.')}</div></fieldset>
    <div className="row"><fieldset className="col"><legend>{tr(L,'Page range')}</legend>{ranges.map(([value,label])=><label key={value}><input type="radio" name="react-wordx-print-range" value={value} checked={range===value} onChange={()=>setRange(value)} />{tr(L,label)}</label>)}<input id="react-wordx-print-pages" type="text" value={pages} onFocus={()=>setRange('pages')} onChange={(event)=>{setRange('pages');setPages(event.currentTarget.value);}} placeholder="1,3,5–12" /><div className="tp-note">{tr(L,'Enter page numbers and/or page ranges separated by commas. For example, 1,3,5–12')}</div></fieldset><fieldset className="col"><legend>{tr(L,'Copies')}</legend><label htmlFor="react-wordx-print-copies">{tr(L,'Number of copies:')}<input id="react-wordx-print-copies" type="number" min={1} max={99} value={copies} onChange={(event)=>setCopies(Number(event.currentTarget.value))} /></label></fieldset></div>
    <div className="row"><label htmlFor="react-wordx-print-what">{tr(L,'Print what:')}<select id="react-wordx-print-what" value={what} onChange={(event)=>setWhat(event.currentTarget.value)}>{kinds.map(([value,label])=><option key={value} value={value}>{tr(L,label)}</option>)}</select></label><label htmlFor="react-wordx-print-which">{tr(L,'Print:')}<select id="react-wordx-print-which" value={which} onChange={(event)=>setWhich(event.currentTarget.value)}><option value="all">{tr(L,'All pages in range')}</option><option value="odd">{tr(L,'Odd pages')}</option><option value="even">{tr(L,'Even pages')}</option></select></label></div>
    <label htmlFor="react-wordx-print-persheet">{tr(L,'Pages per sheet:')}<select id="react-wordx-print-persheet" value={perSheet} onChange={(event)=>setPerSheet(event.currentTarget.value)}>{[[1,'1 page'],[2,'2 pages'],[4,'4 pages'],[6,'6 pages'],[8,'8 pages'],[16,'16 pages']].map(([count,label])=><option key={count} value={count}>{tr(L,String(label))}</option>)}</select></label>
  </div>;
}

type TemplatePickerItem = { id: string; name: string };
type TemplatePickerState = { categories: Array<{ name: string; items: TemplatePickerItem[] }> };
function TemplatePickerDialogBody({ L }: { L: WordxRuntime }) {
  const adapter = L.dlg as (typeof L.dlg & { reactTemplatePickerCatalog?(): TemplatePickerState; reactTemplatePickerAction?(action: string, id?: string): void }) | undefined;
  const catalog = adapter?.reactTemplatePickerCatalog?.() || { categories: [] };
  const first = catalog.categories.find((category) => category.name === 'General') || catalog.categories[0];
  const [category, setCategory] = useState(first?.name || '');
  const [selected, setSelected] = useState('blank');
  const active = catalog.categories.find((item) => item.name === category) || first;
  const item = active?.items.find((entry) => entry.id === selected) || active?.items[0];
  return <div className="row" id="react-wordx-template-picker-body">
    <div className="col" style={{ flex: 1, minWidth: 240 }}><div className="tabs-head" role="tablist">{catalog.categories.map((item) => <button key={item.name} type="button" role="tab" aria-selected={item.name === category} className={`tab${item.name === category ? ' on' : ''}`} onClick={() => { setCategory(item.name); if (item.items[0]) setSelected(item.items[0].id); }}>{tr(L,item.name)}</button>)}</div>
      <div className="preset-grid">{(active?.items || []).map((entry) => <button key={entry.id} type="button" className={entry.id === selected ? 'on' : ''} aria-pressed={entry.id === selected} data-template-id={entry.id} onClick={() => { setSelected(entry.id); adapter?.reactTemplatePickerAction?.('select', entry.id); }} onDoubleClick={() => adapter?.reactTemplatePickerAction?.('accept', entry.id)}><span aria-hidden="true" dangerouslySetInnerHTML={{ __html: L.icons?.get?.('new') || '' }} /><span>{tr(L,entry.name)}</span></button>)}</div>
    </div>
    <div className="col"><div className="cd-cap">{tr(L,'Preview')}</div><div className="preview-box" style={{ width: 150, height: 190 }}><div className="ps-page" style={{ width: 110, height: 140 }}>{Array.from({ length: item?.id === 'blank' ? 0 : 14 }, (_, index) => <i key={index} style={{ top: 14 + index * 8, ...(index % 5 === 4 ? { right: 40 } : {}) }} />)}</div></div></div>
  </div>;
}

type OrganizerStyle = { id: string; name: string; builtin?: boolean };
type OrganizerState = { mineName: string; otherName: string; mine: OrganizerStyle[]; theirs: OrganizerStyle[] };
function OrganizerDialogBody({ L }: { L: WordxRuntime }) {
  const adapter = L.dlg as (typeof L.dlg & { reactOrganizerCatalog?(): OrganizerState; reactOrganizerAction?(action: string, value?: { ids?: string[]; name?: string }): Promise<OrganizerState> | OrganizerState }) | undefined;
  const [state, setState] = useState<OrganizerState>(adapter?.reactOrganizerCatalog?.() || { mineName: '', otherName: '', mine: [], theirs: [] });
  const [mine, setMine] = useState<string[]>([]);
  const [theirs, setTheirs] = useState<string[]>([]);
  const [newName, setNewName] = useState('');
  const refresh = async (action: string, value?: { ids?: string[]; name?: string }) => { const next = await adapter?.reactOrganizerAction?.(action, value); if (next) { setState(next); if (action === 'open' || action === 'copy' || action === 'delete' || action === 'rename') { setMine([]); setTheirs([]); } } };
  const selectedMine = state.mine.find((style) => style.id === mine[0]);
  return <div className="row" id="react-wordx-organizer-body">
    <section className="col" style={{ flex: 1 }}><div className="cd-cap">{tr(L, state.mineName)}</div><select aria-label={tr(L, 'Styles in this document')} size={12} multiple value={mine} onChange={(event) => { const ids=Array.from(event.currentTarget.selectedOptions, (option) => option.value); setMine(ids); const selected=state.mine.find((style) => style.id === ids[0]); setNewName(selected?.name || ''); }}>{state.mine.map((style) => <option key={style.id} value={style.id}>{style.name}</option>)}</select></section>
    <div className="col" style={{ width: 110, justifyContent: 'center' }}><button type="button" className="btn small" disabled={!state.theirs.length || !theirs.length} onClick={() => { void refresh('copy', { ids: theirs }); }}>{tr(L, '<- Copy')}</button><button type="button" className="btn small" disabled={!mine.length || !state.mine.some((style) => mine.includes(style.id) && !style.builtin && style.id !== 'Normal')} onClick={() => { void refresh('delete', { ids: mine }); }}>{tr(L, 'Delete')}</button><input aria-label={tr(L, 'New style name')} value={newName} onChange={(event) => setNewName(event.currentTarget.value)} disabled={!selectedMine || !!selectedMine.builtin || selectedMine.id === 'Normal'} /><button type="button" className="btn small" disabled={!selectedMine || !!selectedMine.builtin || selectedMine.id === 'Normal' || !newName.trim()} onClick={() => { void refresh('rename', { ids: [selectedMine!.id], name: newName }); }}>{tr(L, 'Rename...')}</button></div>
    <section className="col" style={{ flex: 1 }}><div className="cd-cap">{tr(L, state.otherName)}</div><select aria-label={tr(L, 'Styles in another file')} size={12} multiple value={theirs} onChange={(event) => setTheirs(Array.from(event.currentTarget.selectedOptions, (option) => option.value))}>{state.theirs.map((style) => <option key={style.id} value={style.id}>{style.name}</option>)}</select><button type="button" className="btn small" onClick={() => { void refresh('open'); }}>{tr(L, 'Open File...')}</button></section>
  </div>;
}

type TemplatesState = { name: string; update: boolean };
function TemplatesDialogBody({ L }: { L: WordxRuntime }) {
  const adapter = L.dlg as (typeof L.dlg & { reactTemplatesCatalog?(): TemplatesState; reactTemplatesAction?(action: string, value?: boolean): Promise<TemplatesState> | TemplatesState }) | undefined;
  const [state, setState] = useState<TemplatesState>(adapter?.reactTemplatesCatalog?.() || { name: 'Normal', update: false });
  const act = async (action: string, value?: boolean) => { const next = await adapter?.reactTemplatesAction?.(action, value); if (next) setState(next); };
  return <div className="col" id="react-wordx-templates-body">
    <fieldset className="col"><legend>{tr(L, 'Document template')}</legend>
      <label htmlFor="react-wordx-template-name">{tr(L, 'Template:')}<input id="react-wordx-template-name" type="text" value={state.name} readOnly style={{ width: '100%' }} /></label>
      <label><input id="react-wordx-template-update" type="checkbox" checked={state.update} onChange={(event) => setState((current) => ({ ...current, update: event.currentTarget.checked }))} />{tr(L, 'Automatically update document styles')}</label>
      <div className="tp-btns"><button type="button" className="btn small" onClick={() => { void act('attach', state.update); }}>{tr(L, 'Attach...')}</button><button type="button" className="btn small" onClick={() => { void act('organizer'); }}>{tr(L, 'Organizer...')}</button></div>
    </fieldset>
    <div className="tp-note">{tr(L, 'Attaching a template with "Automatically update document styles" copies its styles into this document.')}</div>
  </div>;
}

type PropertiesState = { fileName: string; general: Array<[string, string]>; summary: Array<{ key: string; label: string; value: string }>; comments: string; hyperlinkBase: string; template: string; statistics: Array<[string, string | number]>; headings: Array<{ level: number; text: string }>; custom: Array<{ name: string; type: string; value: unknown }> };
function PropertiesDialogBody({ L }: { L: WordxRuntime }) {
  const adapter = L.dlg as (typeof L.dlg & { reactPropertiesCatalog?(): PropertiesState; reactPropertiesAction?(action: string, value?: unknown): PropertiesState }) | undefined;
  const [state, setState] = useState<PropertiesState>(adapter?.reactPropertiesCatalog?.() || { fileName: '', general: [], summary: [], comments: '', hyperlinkBase: '', template: 'Normal', statistics: [], headings: [], custom: [] });
  const [tab, setTab] = useState('Summary');
  const [customName, setCustomName] = useState('');
  const [customType, setCustomType] = useState('text');
  const [customValue, setCustomValue] = useState('');
  const [selectedCustom, setSelectedCustom] = useState('');
  const tabs = ['General', 'Summary', 'Statistics', 'Contents', 'Custom'];
  const custom = (action: string, value?: unknown) => { const next = adapter?.reactPropertiesAction?.(action, value); if (next) setState(next); };
  return <div className="col" id="react-wordx-properties-body"><div role="tablist" aria-label={tr(L, 'Document Properties')} className="tabs-head">{tabs.map((name) => <button key={name} type="button" role="tab" aria-selected={tab === name} onClick={() => setTab(name)}>{tr(L, name)}</button>)}</div><div role="tabpanel" className="col">
    {tab === 'General' && <><div className="row"><b translate="no">{state.fileName}.docx</b></div><table className="wc-table"><tbody>{state.general.map(([label, value]) => <tr key={label}><td>{tr(L, label)}</td><td translate="no">{value}</td></tr>)}</tbody></table></>}
    {tab === 'Summary' && <>{state.summary.map((field) => <label key={field.key} htmlFor={`react-wordx-property-${field.key}`}>{tr(L, field.label)}<input id={`react-wordx-property-${field.key}`} type="text" defaultValue={field.value} /></label>)}<label htmlFor="react-wordx-properties-comments">{tr(L, 'Comments:')}<textarea id="react-wordx-properties-comments" rows={3} defaultValue={state.comments} /></label><label htmlFor="react-wordx-properties-hyperlink-base">{tr(L, 'Hyperlink base:')}<input id="react-wordx-properties-hyperlink-base" type="text" defaultValue={state.hyperlinkBase} /></label><div className="tp-note">{L.t?.('Template: {0}', state.template)}</div></>}
    {tab === 'Statistics' && <table className="wc-table"><tbody>{state.statistics.map(([label, value], index) => <tr key={`${label}:${index}`}><td>{tr(L, label)}</td><td>{String(value)}</td></tr>)}</tbody></table>}
    {tab === 'Contents' && <><div>{tr(L, 'Document contents:')}</div><div className="tp-box" style={{ maxHeight: 200, overflow: 'auto' }}>{state.headings.length ? state.headings.map((heading, index) => <div key={`${index}:${heading.text}`} translate="no" style={{ paddingLeft: (heading.level - 1) * 12 }}>{heading.text}</div>) : <div>{tr(L, '(no headings)')}</div>}</div></>}
    {tab === 'Custom' && <><label htmlFor="react-wordx-property-custom-name">{tr(L, 'Name:')}<input id="react-wordx-property-custom-name" list="react-wordx-property-custom-names" value={customName} onChange={(event) => setCustomName(event.currentTarget.value)} /></label><datalist id="react-wordx-property-custom-names">{['Checked by','Client','Date completed','Department','Destination','Disposition','Division','Document number','Editor','Forward to','Group','Language','Mailstop','Matter','Office','Owner','Project','Publisher','Purpose','Received from','Recorded by','Recorded date','Reference','Source','Status','Telephone number','Typist'].map((name) => <option key={name} value={name} />)}</datalist><label htmlFor="react-wordx-property-custom-type">{tr(L, 'Type:')}<select id="react-wordx-property-custom-type" value={customType} onChange={(event) => setCustomType(event.currentTarget.value)}>{([['text','Text'],['date','Date'],['number','Number'],['bool','Yes or no']] as Array<[string,string]>).map(([value,label]) => <option key={value} value={value}>{tr(L,label)}</option>)}</select></label><label htmlFor="react-wordx-property-custom-value">{tr(L, 'Value:')}<input id="react-wordx-property-custom-value" type="text" value={customValue} onChange={(event) => setCustomValue(event.currentTarget.value)} /></label><div className="tp-btns"><button type="button" className="btn small" onClick={() => { const next=adapter?.reactPropertiesAction?.('addCustom',{name:customName,type:customType,value:customValue}); if(next) setState(next); setSelectedCustom(customName.trim()); }}>{tr(L, 'Add')}</button><button type="button" className="btn small" onClick={() => { const next=adapter?.reactPropertiesAction?.('deleteCustom',selectedCustom || customName.trim()); if(next) setState(next); setCustomName(''); setCustomValue(''); setSelectedCustom(''); }}>{tr(L, 'Delete')}</button></div><label htmlFor="react-wordx-properties-custom-list">{tr(L, 'Properties:')}</label><select id="react-wordx-properties-custom-list" size={6} value={selectedCustom} onChange={(event) => { const value=event.currentTarget.value, item=state.custom.find((entry) => entry.name===value); setSelectedCustom(value); setCustomName(value); if(item) { setCustomType(item.type); setCustomValue(String(item.value ?? '')); } }}>{state.custom.map((item) => <option key={item.name} value={item.name}>{item.name}: {String(item.value ?? '')}</option>)}</select></>}
  </div></div>;
}

function BookmarkDialogBody({ L }: { L: WordxRuntime }) {
  const [, refresh] = useState(0);
  const [name, setName] = useState(L.dlg?.reactBookmarkCatalog?.().name || '');
  useEffect(() => L.bus?.on('wordx-bookmark-refresh', () => refresh((value) => value + 1)), [L]);
  const catalog = L.dlg?.reactBookmarkCatalog?.();
  const act = (action: string, value?: string) => L.dlg?.reactBookmarkAction?.(action, value);
  return <div className="col">
    <label htmlFor="react-wordx-bookmark-name">{tr(L, 'Bookmark name:')}</label>
    <input id="react-wordx-bookmark-name" type="text" value={name} onChange={(event) => { const value = event.currentTarget.value; setName(value); act('name', value); }} />
    <select aria-label={tr(L, 'Bookmarks')} size={8} value={catalog?.names.includes(name) ? name : ''} onChange={(event) => { const value = event.currentTarget.value; setName(value); act('select', value); }}>
      {(catalog?.names || []).map((item) => <option key={item} value={item} translate="no">{item}</option>)}
    </select>
    <fieldset><legend>{tr(L, 'Sort by:')}</legend><label><input type="radio" name="react-wordx-bookmark-sort" checked={catalog?.sort === 'name'} onChange={() => act('sort', 'name')} />{tr(L, 'Name')}</label><label><input type="radio" name="react-wordx-bookmark-sort" checked={catalog?.sort === 'loc'} onChange={() => act('sort', 'loc')} />{tr(L, 'Location')}</label></fieldset>
    <label><input type="checkbox" checked={catalog?.hidden || false} onChange={(event) => act('hidden', String(event.currentTarget.checked))} />{tr(L, 'Hidden bookmarks')}</label>
  </div>;
}

function InsertObjectDialogBody({ L }: { L: WordxRuntime }) {
  const options = L.dlg?.reactObjectCatalog?.() || [];
  const [tab, setTab] = useState<'new' | 'file'>('new');
  const [selected, setSelected] = useState('chart');
  const act = (action: string, value?: string) => L.dlg?.reactObjectAction?.(action, value);
  return <div className="col">
    <div role="tablist" aria-label={tr(L, 'Insert Object')}>
      <button className={`btn small${tab === 'new' ? ' on' : ''}`} type="button" role="tab" aria-selected={tab === 'new'} onClick={() => { setTab('new'); act('tab', 'new'); }}>{tr(L, 'Create New')}</button>
      <button className={`btn small${tab === 'file' ? ' on' : ''}`} type="button" role="tab" aria-selected={tab === 'file'} onClick={() => { setTab('file'); act('tab', 'file'); }}>{tr(L, 'Create from File')}</button>
    </div>
    {tab === 'new' ? <div className="col"><label htmlFor="react-wordx-object-type">{tr(L, 'Object type:')}</label><select id="react-wordx-object-type" size={6} value={selected} onChange={(event) => { const value=event.currentTarget.value; setSelected(value); act('select', value); }}>{options.map((option) => <option key={option.value} value={option.value}>{option.label}</option>)}</select></div> : <div className="col"><div className="tp-note">{tr(L, 'Insert the contents of a .docx, .txt or .htm file, or a picture.')}</div><button className="btn small" type="button" onClick={() => act('browse')}>{tr(L, 'Browse...')}</button></div>}
  </div>;
}

function EquationDialogBody({ L }: { L: WordxRuntime }) {
  const [value, setValue] = useState('');
  const symbols = '±×÷≠≈≤≥∞√∛∑∏∫∂∆∇πθαβγδλμσφωΩ→⇒∈∉⊂⊃∪∩∀∃¬∧∨²³⁴ⁿ₀₁₂ᵢ½';
  const input = useRef<HTMLTextAreaElement>(null);
  const insertSymbol = (symbol: string) => {
    const node = input.current;
    const start = node?.selectionStart ?? value.length;
    const end = node?.selectionEnd ?? start;
    const next = value.slice(0, start) + symbol + value.slice(end);
    setValue(next);
    requestAnimationFrame(() => { node?.focus(); node?.setSelectionRange(start + symbol.length, start + symbol.length); });
  };
  return <div className="col">
    <div className="sym-recent" style={{ gridTemplateColumns: 'repeat(16,26px)' }}>{[...symbols].map((symbol, index) => <button key={`${symbol}-${index}`} type="button" onClick={() => insertSymbol(symbol)}>{symbol}</button>)}</div>
    <textarea ref={input} rows={3} aria-label={tr(L, 'Equation')} style={{ width: '100%', font: '16px "Cambria Math", Caladea, serif' }} value={value} onChange={(event) => setValue(event.currentTarget.value)} />
    <div className="tp-note">{tr(L, 'Type the equation; it is inserted in Cambria Math italics.')}</div>
  </div>;
}

function BreakDialogBody({ L }: { L: WordxRuntime }) {
  const [selected, setSelected] = useState('page');
  const items = L.dlg?.reactBreakCatalog?.() || [];
  const group = (title: string, options: typeof items) => <fieldset><legend>{tr(L, title)}</legend>{options.map((item) => <label key={item.value} className="row" style={{ gap: 6 }}><input type="radio" name="react-wordx-break-type" data-break-type={item.value} checked={selected === item.value} onChange={() => setSelected(item.value)} /><span>{item.label}</span></label>)}</fieldset>;
  return <div className="col" role="radiogroup" aria-label={tr(L, 'Break types')}>{group('Break types', items.slice(0, 3))}{group('Section break types', items.slice(3))}</div>;
}

function DateTimeDialogBody({ L }: { L: WordxRuntime }) {
  const options = L.dlg?.reactDateTimeCatalog?.() || [];
  const initial = L.dlg?.reactDateTimeState?.();
  const [format, setFormat] = useState(options.some((option) => option.value === initial?.format) ? initial?.format || options[0]?.value || '' : options[0]?.value || '');
  const [language, setLanguage] = useState('en-US');
  const [autoUpdate, setAutoUpdate] = useState(initial?.autoUpdate || false);
  return <div className="col">
    <label htmlFor="react-wordx-date-format">{tr(L, 'Available formats:')}</label>
    <select id="react-wordx-date-format" size={8} value={format} onChange={(event) => setFormat(event.currentTarget.value)}>{options.map((option) => <option key={option.value} value={option.value} translate="no">{option.label}</option>)}</select>
    <label htmlFor="react-wordx-date-language">{tr(L, 'Language:')}</label>
    <select id="react-wordx-date-language" value={language} onChange={(event) => setLanguage(event.currentTarget.value)}><option value="en-US">English (U.S.)</option><option value="en-GB">English (U.K.)</option></select>
    <label><input type="checkbox" checked={autoUpdate} onChange={(event) => setAutoUpdate(event.currentTarget.checked)} />{tr(L, 'Update automatically')}</label>
  </div>;
}

function PageNumbersDialogBody({ L }: { L: WordxRuntime }) {
  const [position, setPosition] = useState('bottom');
  const [alignment, setAlignment] = useState('right');
  const [showFirst, setShowFirst] = useState(true);
  const x: Record<string, string> = { left: '8px', right: '8px', center: '36px', inside: '8px', outside: '8px' };
  const alignments: Array<[string, string]> = [['right','Right'],['left','Left'],['center','Center'],['inside','Inside'],['outside','Outside']];
  return <div className="row">
    <div className="col">
      <label htmlFor="react-wordx-page-position">{tr(L, 'Position:')}</label><select id="react-wordx-page-position" value={position} onChange={(event) => setPosition(event.currentTarget.value)}><option value="bottom">{tr(L, 'Bottom of page (Footer)')}</option><option value="top">{tr(L, 'Top of page (Header)')}</option></select>
      <label htmlFor="react-wordx-page-alignment">{tr(L, 'Alignment:')}</label><select id="react-wordx-page-alignment" value={alignment} onChange={(event) => setAlignment(event.currentTarget.value)}>{alignments.map(([value,label])=><option key={value} value={value}>{tr(L,label)}</option>)}</select>
      <label><input type="checkbox" checked={showFirst} onChange={(event) => setShowFirst(event.currentTarget.checked)} />{tr(L, 'Show number on first page')}</label>
    </div>
    <div className="col"><div className="cd-cap">{tr(L, 'Preview')}</div><div className="ps-page" style={{ width: 80, height: 104, margin: '4px auto' }}><span data-page-number-preview="" style={{ position: 'absolute', ...(position === 'top' ? { top: 6 } : { bottom: 6 }), ...(alignment === 'center' ? { left: 36 } : alignment === 'left' ? { left: 8 } : { right: x[alignment] }), width: 8, height: 8, border: '1px solid #000', background: '#fff' }} /></div></div>
  </div>;
}

function PageNumberFormatDialogBody({ L }: { L: WordxRuntime }) {
  const options = L.dlg?.reactPageNumberFormatCatalog?.() || [];
  const initial = L.dlg?.reactPageNumberFormatState?.();
  const [format, setFormat] = useState(initial?.fmt || 'decimal');
  const [chapter, setChapter] = useState(initial?.chapter || false);
  const [mode, setMode] = useState(initial?.mode || 'cont');
  const [start, setStart] = useState(initial?.start || 1);
  return <div className="col">
    <label htmlFor="react-wordx-page-number-format">{tr(L, 'Number format:')}</label><select id="react-wordx-page-number-format" value={format} onChange={(event) => setFormat(event.currentTarget.value)}>{options.map((option) => <option key={option.value} value={option.value}>{option.label}</option>)}</select>
    <label><input type="checkbox" checked={chapter} onChange={(event) => setChapter(event.currentTarget.checked)} />{tr(L, 'Include chapter number')}</label>
    <fieldset><legend>{tr(L, 'Page numbering')}</legend><label><input type="radio" name="react-wordx-page-number-mode" data-page-number-mode="cont" checked={mode === 'cont'} onChange={() => setMode('cont')} />{tr(L, 'Continue from previous section')}</label><label><input type="radio" name="react-wordx-page-number-mode" data-page-number-mode="start" checked={mode === 'start'} onChange={() => setMode('start')} />{tr(L, 'Start at:')} <input id="react-wordx-page-number-start" type="number" min={0} value={start} onChange={(event) => setStart(Number(event.currentTarget.value))} /></label></fieldset>
  </div>;
}

function CellOperationDialogBody({ L }: { L: WordxRuntime }) {
  const kind = L.dlg?.reactCellDialogKind || 'insert';
  const options = L.dlg?.reactCellDialogCatalog?.(kind) || [];
  const [selected, setSelected] = useState(kind === 'insert' ? 'down' : 'left');
  return <fieldset><legend>{tr(L, kind === 'insert' ? 'Insert cells' : 'Delete cells')}</legend>{options.map(([value, label]) => <label key={value} className="row" style={{ gap: 6 }}><input type="radio" name="react-wordx-cell-operation" data-cell-operation={value} checked={selected === value} onChange={() => setSelected(value)} /><span>{tr(L, label)}</span></label>)}</fieldset>;
}

function InsertTableDialogBody({ L }: { L: WordxRuntime }) {
  const initial = L.dlg?.reactInsertTableState?.() || { columns: 5, rows: 2 };
  const [columns, setColumns] = useState(initial.columns);
  const [rows, setRows] = useState(initial.rows);
  const [fit, setFit] = useState('fixed');
  const [width, setWidth] = useState('');
  const [remember, setRemember] = useState(false);
  const [style, setStyle] = useState('Table Grid');
  useEffect(() => {
    const onStyle = (event: Event) => setStyle((event as CustomEvent<{ name: string }>).detail.name);
    window.addEventListener('wordx-insert-table-style', onStyle);
    return () => window.removeEventListener('wordx-insert-table-style', onStyle);
  }, []);
  return <div className="col">
    <fieldset><legend>{tr(L, 'Table size')}</legend><label>{tr(L, 'Number of columns:')} <input id="react-wordx-table-columns" type="number" min={1} max={63} value={columns} onChange={(event) => setColumns(Number(event.currentTarget.value))} /></label><label>{tr(L, 'Number of rows:')} <input id="react-wordx-table-rows" type="number" min={1} max={32767} value={rows} onChange={(event) => setRows(Number(event.currentTarget.value))} /></label></fieldset>
    <fieldset><legend>{tr(L, 'AutoFit behavior')}</legend><label><input type="radio" name="react-wordx-table-fit" data-table-fit="fixed" checked={fit === 'fixed'} onChange={() => setFit('fixed')} />{tr(L, 'Fixed column width:')} <input id="react-wordx-table-width" type="text" value={width} onChange={(event) => setWidth(event.currentTarget.value)} /></label><label><input type="radio" name="react-wordx-table-fit" data-table-fit="contents" checked={fit === 'contents'} onChange={() => setFit('contents')} />{tr(L, 'AutoFit to contents')}</label><label><input type="radio" name="react-wordx-table-fit" data-table-fit="window" checked={fit === 'window'} onChange={() => setFit('window')} />{tr(L, 'AutoFit to window')}</label></fieldset>
    <div className="row"><span className="tp-note">{L.t?.('Table style: {0}', style) || `Table style: ${style}`}</span><button className="btn small" type="button" onClick={() => L.dlg?.reactInsertTableAction?.('autoFormat')}>{tr(L, 'AutoFormat...')}</button></div>
    <label><input type="checkbox" checked={remember} onChange={(event) => setRemember(event.currentTarget.checked)} />{tr(L, 'Remember dimensions for new tables')}</label>
  </div>;
}

function SplitCellsDialogBody({ L }: { L: WordxRuntime }) {
  const initial = L.dlg?.reactSplitCellsState?.() || { columns: 2, rows: 1, merge: false };
  const [columns, setColumns] = useState(initial.columns);
  const [rows, setRows] = useState(initial.rows);
  const [merge, setMerge] = useState(initial.merge);
  return <div className="col">
    <label>{tr(L, 'Number of columns:')} <input id="react-wordx-split-columns" type="number" min={1} max={63} value={columns} onChange={(event) => setColumns(Number(event.currentTarget.value))} /></label>
    <label>{tr(L, 'Number of rows:')} <input id="react-wordx-split-rows" type="number" min={1} max={100} value={rows} onChange={(event) => setRows(Number(event.currentTarget.value))} /></label>
    <label><input type="checkbox" checked={merge} onChange={(event) => setMerge(event.currentTarget.checked)} />{tr(L, 'Merge cells before split')}</label>
  </div>;
}

function TextToTableDialogBody({ L }: { L: WordxRuntime }) {
  const source = L.dlg?.reactTextToTableState?.() || { initialSeparator: 'para', columns: 1, paragraphs: [] };
  const [separator, setSeparator] = useState(source.initialSeparator === '\t' ? 'tabs' : source.initialSeparator === ',' ? 'commas' : 'paragraphs');
  const [custom, setCustom] = useState('-');
  const [columns, setColumns] = useState(source.columns);
  const [fit, setFit] = useState('fixed');
  const paragraphs = source.paragraphs;
  const columnCount = (kind: string) => Math.max(1, ...paragraphs.map((paragraph) => kind === 'tabs' ? paragraph.tabs + 1 : (paragraph.text.split(kind === 'commas' ? ',' : custom || '-').length)));
  useEffect(() => { setColumns(separator === 'paragraphs' ? 1 : columnCount(separator)); }, [separator, custom]);
  const rows = separator === 'paragraphs' ? Math.ceil(paragraphs.length / Math.max(1, columns)) : paragraphs.length;
  return <div className="col">
    <fieldset><legend>{tr(L, 'Table size')}</legend><label>{tr(L, 'Number of columns:')} <input id="react-wordx-text-table-columns" type="number" min={1} max={63} value={columns} onChange={(event) => setColumns(Number(event.currentTarget.value))} /></label><label>{tr(L, 'Number of rows:')} <output>{rows}</output></label></fieldset>
    <fieldset><legend>{tr(L, 'AutoFit behavior')}</legend><label><input type="radio" name="react-wordx-text-table-fit" data-text-fit="fixed" checked={fit === 'fixed'} onChange={() => setFit('fixed')} />{tr(L, 'Fixed column width')}</label><label><input type="radio" name="react-wordx-text-table-fit" data-text-fit="contents" checked={fit === 'contents'} onChange={() => setFit('contents')} />{tr(L, 'AutoFit to contents')}</label><label><input type="radio" name="react-wordx-text-table-fit" data-text-fit="window" checked={fit === 'window'} onChange={() => setFit('window')} />{tr(L, 'AutoFit to window')}</label></fieldset>
    <fieldset><legend>{tr(L, 'Separate text at')}</legend><label><input type="radio" name="react-wordx-text-separator" data-text-separator="paragraphs" checked={separator === 'paragraphs'} onChange={() => setSeparator('paragraphs')} />{tr(L, 'Paragraphs')}</label><label><input type="radio" name="react-wordx-text-separator" data-text-separator="commas" checked={separator === 'commas'} onChange={() => setSeparator('commas')} />{tr(L, 'Commas')}</label><label><input type="radio" name="react-wordx-text-separator" data-text-separator="tabs" checked={separator === 'tabs'} onChange={() => setSeparator('tabs')} />{tr(L, 'Tabs')}</label><label><input type="radio" name="react-wordx-text-separator" data-text-separator="other" checked={separator === 'other'} onChange={() => setSeparator('other')} />{tr(L, 'Other:')} <input id="react-wordx-text-separator" type="text" maxLength={1} value={custom} onChange={(event) => setCustom(event.currentTarget.value)} /></label></fieldset>
  </div>;
}

function TableToTextDialogBody({ L }: { L: WordxRuntime }) {
  const [separator, setSeparator] = useState('\t');
  const [custom, setCustom] = useState('-');
  return <fieldset><legend>{tr(L, 'Separate text with')}</legend><label><input type="radio" name="react-wordx-table-text-separator" data-table-text-separator="para" checked={separator === 'para'} onChange={() => setSeparator('para')} />{tr(L, 'Paragraph marks')}</label><label><input type="radio" name="react-wordx-table-text-separator" data-table-text-separator="tabs" checked={separator === '\t'} onChange={() => setSeparator('\t')} />{tr(L, 'Tabs')}</label><label><input type="radio" name="react-wordx-table-text-separator" data-table-text-separator="commas" checked={separator === ','} onChange={() => setSeparator(',')} />{tr(L, 'Commas')}</label><label><input type="radio" name="react-wordx-table-text-separator" data-table-text-separator="other" checked={separator === 'other'} onChange={() => setSeparator('other')} />{tr(L, 'Other:')} <input id="react-wordx-table-text-separator" type="text" maxLength={1} value={custom} onChange={(event) => setCustom(event.currentTarget.value)} /></label></fieldset>;
}

function FormulaDialogBody({ L }: { L: WordxRuntime }) {
  const catalog = L.dlg?.reactFormulaCatalog?.() || { guess: '=', functions: [], bookmarks: [], formats: [] };
  const [formula, setFormula] = useState(catalog.guess);
  const [format, setFormat] = useState('');
  const [selectedFunction, setSelectedFunction] = useState('');
  const [bookmark, setBookmark] = useState('');
  const input = useRef<HTMLInputElement>(null);
  const addFunction = (name: string) => {
    if (!name) return;
    const node = input.current;
    const start = node?.selectionStart ?? formula.length;
    const end = node?.selectionEnd ?? start;
    const next = formula.slice(0, start) + `${name}()` + formula.slice(end);
    setFormula(next);
    setSelectedFunction('');
    requestAnimationFrame(() => { node?.focus(); node?.setSelectionRange(start + name.length + 1, start + name.length + 1); });
  };
  const addBookmark = (name: string) => { if (name) setFormula((current) => current + name); setBookmark(''); input.current?.focus(); };
  return <div className="col">
    <label htmlFor="react-wordx-formula-input">{tr(L, 'Formula:')}</label><input ref={input} id="react-wordx-formula-input" type="text" value={formula} onChange={(event) => setFormula(event.currentTarget.value)} />
    <label htmlFor="react-wordx-formula-format">{tr(L, 'Number format:')}</label><input id="react-wordx-formula-format" type="text" list="react-wordx-formula-formats" value={format} onChange={(event) => setFormat(event.currentTarget.value)} /><datalist id="react-wordx-formula-formats">{catalog.formats.map((item) => <option key={item} value={item} />)}</datalist>
    <div className="row"><select aria-label={tr(L, 'Paste function')} value={selectedFunction} onChange={(event) => addFunction(event.currentTarget.value)}><option value="">{tr(L, 'Paste function')}</option>{catalog.functions.map((name) => <option key={name} value={name} translate="no">{name}</option>)}</select><select aria-label={tr(L, 'Paste bookmark')} value={bookmark} onChange={(event) => addBookmark(event.currentTarget.value)}><option value="">{tr(L, 'Paste bookmark')}</option>{catalog.bookmarks.map((name) => <option key={name} value={name} translate="no">{name}</option>)}</select></div>
  </div>;
}

function CrossReferenceDialogBody({ L }: { L: WordxRuntime }) {
  const [type, setType] = useState('Heading');
  const [reference, setReference] = useState('text');
  const [item, setItem] = useState('0');
  const [hyperlink, setHyperlink] = useState(true);
  const [aboveBelow, setAboveBelow] = useState(false);
  const catalog = L.dlg?.reactCrossRefCatalog?.(type) || { types: [], references: [], items: [] };
  useEffect(() => {
    setReference(catalog.references[0]?.value || '');
    setItem(catalog.items[0]?.value || '');
  }, [type]);
  return <div className="col">
    <div className="row"><label>{tr(L, 'Reference type:')} <select id="react-wordx-cross-reference-type" value={type} onChange={(event) => setType(event.currentTarget.value)}>{catalog.types.map((entry) => <option key={entry.value} value={entry.value}>{entry.label}</option>)}</select></label><label>{tr(L, 'Insert reference to:')} <select id="react-wordx-cross-reference-reference" value={reference} onChange={(event) => setReference(event.currentTarget.value)}>{catalog.references.map((entry) => <option key={entry.value} value={entry.value}>{tr(L, entry.label)}</option>)}</select></label></div>
    <div className="row"><label><input id="react-wordx-cross-reference-link" type="checkbox" checked={hyperlink} onChange={(event) => setHyperlink(event.currentTarget.checked)} />{tr(L, 'Insert as hyperlink')}</label><label><input id="react-wordx-cross-reference-above-below" type="checkbox" checked={aboveBelow} onChange={(event) => setAboveBelow(event.currentTarget.checked)} />{tr(L, 'Include above/below')}</label></div>
    <label htmlFor="react-wordx-cross-reference-items">{tr(L, 'For which item:')}</label><select id="react-wordx-cross-reference-items" size={8} value={item} onChange={(event) => setItem(event.currentTarget.value)}>{catalog.items.map((entry) => <option key={entry.value} value={entry.value}>{entry.label}</option>)}</select>
  </div>;
}

function MarkEntryDialogBody({ L }: { L: WordxRuntime }) {
  const initial = L.dlg?.reactMarkEntryState?.() || { main: '', sub: '', option: 'cur', xref: 'See ', bold: false, italic: false };
  const [main, setMain] = useState(initial.main);
  const [sub, setSub] = useState(initial.sub);
  const [option, setOption] = useState(initial.option);
  const [xref, setXref] = useState(initial.xref);
  const [bold, setBold] = useState(initial.bold);
  const [italic, setItalic] = useState(initial.italic);
  return <div className="col">
    <label htmlFor="react-wordx-mark-main">{tr(L, 'Main entry:')}</label><input id="react-wordx-mark-main" type="text" value={main} onChange={(event) => setMain(event.currentTarget.value)} />
    <label htmlFor="react-wordx-mark-sub">{tr(L, 'Subentry:')}</label><input id="react-wordx-mark-sub" type="text" value={sub} onChange={(event) => setSub(event.currentTarget.value)} />
    <fieldset><legend>{tr(L, 'Options')}</legend><label><input type="radio" name="react-wordx-mark-option" value="cur" checked={option === 'cur'} onChange={() => setOption('cur')} />{tr(L, 'Current page')}</label><label><input type="radio" name="react-wordx-mark-option" value="xref" checked={option === 'xref'} onChange={() => setOption('xref')} />{tr(L, 'Cross-reference:')} <input id="react-wordx-mark-xref" type="text" value={xref} onChange={(event) => setXref(event.currentTarget.value)} /></label></fieldset>
    <fieldset><legend>{tr(L, 'Page number format')}</legend><label><input id="react-wordx-mark-bold" type="checkbox" checked={bold} onChange={(event) => setBold(event.currentTarget.checked)} />{tr(L, 'Bold')}</label><label><input id="react-wordx-mark-italic" type="checkbox" checked={italic} onChange={(event) => setItalic(event.currentTarget.checked)} />{tr(L, 'Italic')}</label></fieldset>
  </div>;
}

type FieldDialogAdapter = {
  reactFieldState?(): { code: string; type: string; editing: boolean; format: string; datePicture: string; keep: boolean };
  reactFieldCatalog?(category?: string): { categories: Array<{ value: string; label: string }>; names: Array<{ value: string; label: string }>; formats: Array<{ value: string; label: string }>; datePictures: Array<{ value: string; label: string }> };
  reactFieldDescription?(name: string): string;
  reactFieldPreview?(code: string, format: string, datePicture: string, keep: boolean): string;
  reactFieldAction?(code: string, format: string, datePicture: string, keep: boolean): boolean | void;
};

function FieldDialogBody({ L }: { L: WordxRuntime }) {
  const adapter = L.dlg as FieldDialogAdapter | undefined;
  const initial = adapter?.reactFieldState?.() || { code: '', type: '', editing: false, format: '', datePicture: '', keep: true };
  const [category, setCategory] = useState('all');
  const [name, setName] = useState(initial.type);
  const [code, setCode] = useState(initial.code);
  const [format, setFormat] = useState(initial.format);
  const [datePicture, setDatePicture] = useState(initial.datePicture);
  const [keep, setKeep] = useState(initial.keep);
  const [description, setDescription] = useState(adapter?.reactFieldDescription?.(initial.type) || '');
  const catalog = adapter?.reactFieldCatalog?.(category) || { categories: [], names: [], formats: [], datePictures: [] };
  const preview = adapter?.reactFieldPreview?.(code, format, datePicture, keep) || '';
  return <div className="row">
    <div className="col" style={{ width: 190 }}>
      <label htmlFor="react-wordx-field-category">{tr(L, 'Categories:')}</label><select id="react-wordx-field-category" value={category} onChange={(event) => setCategory(event.currentTarget.value)}>{catalog.categories.map((item) => <option key={item.value} value={item.value}>{item.label}</option>)}</select>
      <label htmlFor="react-wordx-field-name">{tr(L, 'Field names:')}</label><select id="react-wordx-field-name" size={12} value={name} onChange={(event) => { const next = event.currentTarget.value; setName(next); setCode(next === '=' ? '= ' : next + ' '); setDescription(adapter?.reactFieldDescription?.(next) || ''); }}>{catalog.names.map((item) => <option key={item.value} value={item.value} translate="no">{item.label}</option>)}</select>
    </div>
    <div className="col" style={{ flex: 1, minWidth: 220 }}>
      <label htmlFor="react-wordx-field-code">{tr(L, 'Field codes:')}</label><input id="react-wordx-field-code" type="text" value={code} onChange={(event) => setCode(event.currentTarget.value)} />
      <div className="tp-note" style={{ minHeight: 28 }}>{description}</div>
      <fieldset><legend>{tr(L, 'Field options')}</legend><label>{tr(L, 'Format:')} <select id="react-wordx-field-format" value={format} onChange={(event) => setFormat(event.currentTarget.value)}>{catalog.formats.map((item) => <option key={item.value} value={item.value}>{item.label}</option>)}</select></label><label>{tr(L, 'Date picture:')} <select id="react-wordx-field-date-picture" value={datePicture} onChange={(event) => setDatePicture(event.currentTarget.value)}>{catalog.datePictures.map((item) => <option key={item.value} value={item.value}>{item.label}</option>)}</select></label></fieldset>
      <label>{tr(L, 'Result preview')}</label><div className="tp-box" style={{ minHeight: 22 }} translate="no">{preview}</div>
      <label><input id="react-wordx-field-keep" type="checkbox" checked={keep} onChange={(event) => setKeep(event.currentTarget.checked)} />{tr(L, 'Preserve formatting during updates')}</label>
    </div>
  </div>;
}

type FootnoteDialogAdapter = {
  reactFootnoteState?(): { kind: string; fnPos: string; enPos: string; fnFmt: string; enFmt: string; fnStart: number; enStart: number; fnRestart: string; enRestart: string; custom: string };
  reactFootnoteCatalog?(): { locations: Record<string, Array<[string, string]>>; formats: Array<[string, string]>; restarts: Array<[string, string]> };
  reactFootnoteAction?(action: string, values: { kind: string; fnPos: string; enPos: string; fnFmt: string; enFmt: string; fnStart: number; enStart: number; fnRestart: string; enRestart: string; custom: string }): void | boolean;
};

type DiagramDialogAdapter = {
  reactDiagramCatalog?(): Array<{ value: string; label: string }>;
  reactDiagramAction?(kind: string): void;
};

function DiagramPreview({ kind }: { kind: string }) {
  if (kind === 'pyramid') return <svg viewBox="0 0 56 44" aria-hidden="true"><path d="M28 2 50 42H6z" fill="#fc9" stroke="#333" /><path d="M19 18h18M12 30h32" stroke="#333" /></svg>;
  if (kind === 'venn') return <svg viewBox="0 0 56 44" aria-hidden="true"><circle cx="22" cy="18" r="12" fill="#9cf" fillOpacity=".6" stroke="#333" /><circle cx="34" cy="18" r="12" fill="#fc9" fillOpacity=".6" stroke="#333" /><circle cx="28" cy="28" r="12" fill="#cf9" fillOpacity=".6" stroke="#333" /></svg>;
  if (kind === 'target') return <svg viewBox="0 0 56 44" aria-hidden="true"><circle cx="24" cy="24" r="18" fill="#9cf" stroke="#333" /><circle cx="24" cy="24" r="12" fill="#fc9" stroke="#333" /><circle cx="24" cy="24" r="6" fill="#f99" stroke="#333" /></svg>;
  if (kind === 'org') return <svg viewBox="0 0 56 44" aria-hidden="true"><rect x="20" y="2" width="16" height="10" fill="#fdd" stroke="#333" /><path d="M28 12v6M10 18h36M10 18v4M28 18v4M46 18v4" stroke="#333" fill="none" /><rect x="2" y="22" width="16" height="10" fill="#dfd" stroke="#333" /><rect x="20" y="22" width="16" height="10" fill="#dfd" stroke="#333" /><rect x="38" y="22" width="16" height="10" fill="#dfd" stroke="#333" /></svg>;
  if (kind === 'radial') return <svg viewBox="0 0 56 44" aria-hidden="true"><path d="M28 22 28 6M28 22 44 32M28 22 12 32" stroke="#333" /><circle cx="28" cy="22" r="7" fill="#fc9" stroke="#333" /><circle cx="28" cy="6" r="5" fill="#9cf" stroke="#333" /><circle cx="44" cy="32" r="5" fill="#9cf" stroke="#333" /><circle cx="12" cy="32" r="5" fill="#9cf" stroke="#333" /></svg>;
  return <svg viewBox="0 0 56 44" aria-hidden="true"><circle cx="28" cy="22" r="16" fill="none" stroke="#888" strokeDasharray="3 2" /><circle cx="28" cy="6" r="5" fill="#9cf" stroke="#333" /><circle cx="42" cy="30" r="5" fill="#fc9" stroke="#333" /><circle cx="14" cy="30" r="5" fill="#cf9" stroke="#333" /></svg>;
}

function DiagramDialogBody({ L }: { L: WordxRuntime }) {
  const adapter = L.dlg as DiagramDialogAdapter | undefined;
  const catalog = adapter?.reactDiagramCatalog?.() || [];
  const [selected, setSelected] = useState('org');
  return <div className="col">
    <div>{tr(L, 'Select a diagram type:')}</div>
    <div className="preset-grid">{catalog.map((item) => <button type="button" key={item.value} data-wordx-diagram={item.value} aria-label={item.label} aria-pressed={selected === item.value} className={selected === item.value ? 'on' : ''} onClick={() => setSelected(item.value)} onDoubleClick={() => adapter?.reactDiagramAction?.(item.value)}><DiagramPreview kind={item.value} /><span>{item.label}</span></button>)}</div>
  </div>;
}

function FootnoteDialogBody({ L }: { L: WordxRuntime }) {
  const adapter = L.dlg as FootnoteDialogAdapter | undefined;
  const [state, setState] = useState(adapter?.reactFootnoteState?.() || { kind: 'fn', fnPos: 'pageBottom', enPos: 'docEnd', fnFmt: 'decimal', enFmt: 'lowerRoman', fnStart: 1, enStart: 1, fnRestart: 'continuous', enRestart: 'continuous', custom: '' });
  const catalog = adapter?.reactFootnoteCatalog?.() || { locations: {}, formats: [], restarts: [] };
  const update = <K extends keyof typeof state>(key: K, value: (typeof state)[K]) => setState((current) => ({ ...current, [key]: value }));
  const kind = state.kind === 'fn' ? 'fn' : 'en';
  const formatId = kind === 'fn' ? 'react-wordx-footnote-fnFmt' : 'react-wordx-footnote-enFmt';
  const startId = kind === 'fn' ? 'react-wordx-footnote-fnStart' : 'react-wordx-footnote-enStart';
  const restartId = kind === 'fn' ? 'react-wordx-footnote-fnRestart' : 'react-wordx-footnote-enRestart';
  const set = (action: string) => adapter?.reactFootnoteAction?.(action, state);
  return <div className="col">
    <fieldset><legend>{tr(L, 'Location')}</legend><label><input id="react-wordx-footnote-kind-fn" type="radio" name="react-wordx-footnote-kind" value="fn" checked={kind === 'fn'} onChange={() => update('kind', 'fn')} />{tr(L, 'Footnotes:')} <select id="react-wordx-footnote-fnPos" value={state.fnPos} onChange={(event) => update('fnPos', event.currentTarget.value)}>{(catalog.locations.fn || []).map(([value, label]) => <option key={value} value={value}>{tr(L, label)}</option>)}</select></label><label><input id="react-wordx-footnote-kind-en" type="radio" name="react-wordx-footnote-kind" value="en" checked={kind === 'en'} onChange={() => update('kind', 'en')} />{tr(L, 'Endnotes:')} <select id="react-wordx-footnote-enPos" value={state.enPos} onChange={(event) => update('enPos', event.currentTarget.value)}>{(catalog.locations.en || []).map(([value, label]) => <option key={value} value={value}>{tr(L, label)}</option>)}</select></label><button type="button" className="btn small" onClick={() => set('convert')}>{tr(L, 'Convert...')}</button></fieldset>
    <fieldset><legend>{tr(L, 'Format')}</legend><label>{tr(L, 'Number format:')} <select id={formatId} value={kind === 'fn' ? state.fnFmt : state.enFmt} onChange={(event) => update(kind === 'fn' ? 'fnFmt' : 'enFmt', event.currentTarget.value)}>{catalog.formats.map(([value, label]) => <option key={value} value={value}>{tr(L, label)}</option>)}</select></label><label>{tr(L, 'Custom mark:')} <input id="react-wordx-footnote-custom" type="text" value={state.custom} onChange={(event) => update('custom', event.currentTarget.value)} /></label><button type="button" className="btn small" onClick={() => set('symbol')}>{tr(L, 'Symbol...')}</button><label>{tr(L, 'Start at:')} <input id={startId} type="number" min={1} value={kind === 'fn' ? state.fnStart : state.enStart} onChange={(event) => update(kind === 'fn' ? 'fnStart' : 'enStart', Math.max(1, Number(event.currentTarget.value)))} /></label><label>{tr(L, 'Numbering:')} <select id={restartId} value={kind === 'fn' ? state.fnRestart : state.enRestart} onChange={(event) => update(kind === 'fn' ? 'fnRestart' : 'enRestart', event.currentTarget.value)}>{catalog.restarts.map(([value, label]) => <option key={value} value={value}>{tr(L, label)}</option>)}</select></label></fieldset>
  </div>;
}

function ConvertNotesDialogBody({ L }: { L: WordxRuntime }) {
  const [kind,setKind]=useState('fn');
  return <fieldset className="col" id="react-wordx-convert-notes-body"><label><input type="radio" name="react-wordx-convert-notes" value="fn" checked={kind==='fn'} onChange={()=>setKind('fn')}/>{tr(L,'Convert all footnotes to endnotes')}</label><label><input type="radio" name="react-wordx-convert-notes" value="en" checked={kind==='en'} onChange={()=>setKind('en')}/>{tr(L,'Convert all endnotes to footnotes')}</label></fieldset>;
}

type WordxStyleState={name:string;builtIn:boolean;isNew:boolean;type:string;basedOn:string;next:string;styles:Array<{id:string;name:string}>;fonts:string[];sizes:string[];font:string;size:number;bold:boolean;italic:boolean;underline:string;color:string;automaticColor:boolean;align:string;autoRedefine:boolean;preview:{font:string;size:number;bold:boolean;italic:boolean;underline:boolean;color:string;align:CSSProperties['textAlign']}};
function WordxStyleDialogBody({L}:{L:WordxRuntime}){
  const adapter=L.dlg as (typeof L.dlg & {reactStyleCatalog?():WordxStyleState;reactStyleAction?(action:string,value?:any):WordxStyleState|void})|undefined;
  const [state,setState]=useState<WordxStyleState>(adapter?.reactStyleCatalog?.()||{name:'Style',builtIn:false,isNew:true,type:'paragraph',basedOn:'Normal',next:'__self',styles:[],fonts:[],sizes:[],font:'Times New Roman',size:12,bold:false,italic:false,underline:'none',color:'#000000',automaticColor:true,align:'left',autoRedefine:false,preview:{font:'Times New Roman',size:12,bold:false,italic:false,underline:false,color:'#000000',align:'left'}});
  useEffect(()=>L.bus?.on('wordx-style-refresh',()=>{const next=adapter?.reactStyleCatalog?.();if(next)setState(next)}),[L,adapter]);
  const update=(value:Partial<WordxStyleState>)=>{const next=adapter?.reactStyleAction?.('change',value);if(next)setState(next)};
  const action=(name:string,value?:unknown)=>{const next=adapter?.reactStyleAction?.(name,value);if(next)setState(next as WordxStyleState)};
  return <div className="col" id="react-wordx-style-body"><fieldset className="col"><legend>{tr(L,'Properties')}</legend><label>{tr(L,'Name:')}<input data-style-field="name" value={state.name} disabled={state.builtIn&&!state.isNew} onChange={(e)=>update({name:e.currentTarget.value})}/></label><label>{tr(L,'Style type:')}<select data-style-field="type" value={state.type} disabled={!state.isNew} onChange={(e)=>update({type:e.currentTarget.value})}>{([['paragraph','Paragraph'],['character','Character'],['table','Table'],['numbering','List']] as const).map(([value,label])=><option key={value} value={value}>{tr(L,label)}</option>)}</select></label><label>{tr(L,'Style based on:')}<select data-style-field="basedOn" value={state.basedOn} onChange={(e)=>update({basedOn:e.currentTarget.value})}><option value="">{tr(L,'(no style)')}</option>{state.styles.map((item)=><option key={item.id} value={item.id}>{item.name}</option>)}</select></label>{state.type==='paragraph'&&<label>{tr(L,'Style for following paragraph:')}<select data-style-field="next" value={state.next} onChange={(e)=>update({next:e.currentTarget.value})}>{state.styles.map((item)=><option key={item.id} value={item.id}>{item.name}</option>)}{state.isNew&&<option value="__self">{tr(L,'(this style)')}</option>}</select></label>}</fieldset><fieldset className="col"><legend>{tr(L,'Formatting')}</legend><div className="row"><select aria-label={tr(L,'Font')} data-style-field="font" value={state.font} onChange={(e)=>update({font:e.currentTarget.value})}>{state.fonts.map((font)=><option key={font}>{font}</option>)}</select><select aria-label={tr(L,'Font size')} data-style-field="size" value={String(state.size)} onChange={(e)=>update({size:Number(e.currentTarget.value)})}>{state.sizes.map((size)=><option key={size}>{size}</option>)}</select><button type="button" aria-label={tr(L,'Bold')} aria-pressed={state.bold} className={`tb-btn${state.bold?' on':''}`} data-style-field="bold" onClick={()=>update({bold:!state.bold})}><b>B</b></button><button type="button" aria-label={tr(L,'Italic')} aria-pressed={state.italic} className={`tb-btn${state.italic?' on':''}`} data-style-field="italic" onClick={()=>update({italic:!state.italic})}><i>I</i></button><button type="button" aria-label={tr(L,'Underline')} aria-pressed={state.underline!=='none'} className={`tb-btn${state.underline!=='none'?' on':''}`} data-style-field="underline" onClick={()=>update({underline:state.underline==='none'?'single':'none'})}><u>U</u></button><button type="button" aria-label={tr(L,'Font color')} data-style-field="color" onClick={(e)=>action('color',e.currentTarget)}><span className="cpk-sw" style={{background:state.automaticColor?'transparent':state.color}}/></button>{([['left','Align Left'],['center','Center'],['right','Align Right'],['both','Justify']] as const).map(([value,label])=><button type="button" key={value} aria-label={tr(L,label)} aria-pressed={state.align===value} className={`tb-btn${state.align===value?' on':''}`} onClick={()=>update({align:value})}>{value==='both'?'☰':value==='left'?'≡':value==='center'?'≋':'≡'}</button>)}</div><div className="preview-box" style={{height:100,flexDirection:'column',alignItems:'stretch',justifyContent:'center',padding:'6px 10px'}}><i style={{height:3,background:'#ccc'}}/><div style={{fontFamily:state.preview.font,fontSize:Math.min(state.preview.size,24),fontWeight:state.preview.bold?'bold':undefined,fontStyle:state.preview.italic?'italic':undefined,textDecoration:state.preview.underline?'underline':undefined,color:state.preview.color,textAlign:state.preview.align,whiteSpace:'nowrap',overflow:'hidden'}}>Sample Text Sample Text Sample Text</div><i style={{height:3,background:'#ccc'}}/></div></fieldset><div className="row"><label><input type="checkbox" data-style-field="autoRedefine" checked={state.autoRedefine} onChange={(e)=>update({autoRedefine:e.currentTarget.checked})}/>{tr(L,'Automatically update')}</label><label><input type="checkbox" data-style-field="addToTemplate"/>{tr(L,'Add to template')}</label></div><div className="tp-btns">{tr(L,'Advanced formatting uses the existing Font, Paragraph, Tabs, Border and Language dialogs.')}</div></div>;
}

/** React-owned Wordx Help body, backed by the editor's existing localized topic and shortcut catalog. */
export function WordxTaskContentLayer({ host }: { host: HTMLElement }) {
  const [mount, setMount] = useState<HTMLElement | null>(null);
  const [mapMount, setMapMount] = useState<HTMLElement | null>(null);
  const [reviewMount, setReviewMount] = useState<HTMLElement | null>(null);
  const [readingMount, setReadingMount] = useState<HTMLElement | null>(null);
  const [taskHeadMount, setTaskHeadMount] = useState<HTMLElement | null>(null);
  const [statusMount, setStatusMount] = useState<HTMLElement | null>(null);
  const [gotoMount, setGotoMount] = useState<HTMLElement | null>(null);
  const [findInputMount, setFindInputMount] = useState<HTMLElement | null>(null);
  const [replaceInputMount, setReplaceInputMount] = useState<HTMLElement | null>(null);
  const [findAuxMount, setFindAuxMount] = useState<HTMLElement | null>(null);
  const [findOptionsMount, setFindOptionsMount] = useState<HTMLElement | null>(null);
  const [findFormatMount, setFindFormatMount] = useState<HTMLElement | null>(null);
  const [findFormatSummaryMount, setFindFormatSummaryMount] = useState<HTMLElement | null>(null);
  const [replaceFormatSummaryMount, setReplaceFormatSummaryMount] = useState<HTMLElement | null>(null);
  const [findMoreMount, setFindMoreMount] = useState<HTMLElement | null>(null);
  const [findActionsMount, setFindActionsMount] = useState<HTMLElement | null>(null);
  const [sortMount, setSortMount] = useState<HTMLElement | null>(null);
  const [mailMergeRecipientsMount, setMailMergeRecipientsMount] = useState<HTMLElement | null>(null);
  const [newAddressListMount, setNewAddressListMount] = useState<HTMLElement | null>(null);
  const [mergeFieldMount, setMergeFieldMount] = useState<HTMLElement | null>(null);
  const [reviewersMount, setReviewersMount] = useState<HTMLElement | null>(null);
  const [commentEditMount, setCommentEditMount] = useState<HTMLElement | null>(null);
  const [contentControlMount, setContentControlMount] = useState<HTMLElement | null>(null);
  const [spellCheckMount, setSpellCheckMount] = useState<HTMLElement | null>(null);
  const [tableAutoFormatMount, setTableAutoFormatMount] = useState<HTMLElement | null>(null);
  const [tabsDialogMount, setTabsDialogMount] = useState<HTMLElement | null>(null);
  const [paragraphDialogMount, setParagraphDialogMount] = useState<HTMLElement | null>(null);
  const [bordersDialogMount, setBordersDialogMount] = useState<HTMLElement | null>(null);
  const [tablePropsDialogMount, setTablePropsDialogMount] = useState<HTMLElement | null>(null);
  const [rulerCornerMount, setRulerCornerMount] = useState<HTMLElement | null>(null);
  const [rulerCanvasMount, setRulerCanvasMount] = useState<HTMLElement | null>(null);
  const [verticalRulerCanvasMount, setVerticalRulerCanvasMount] = useState<HTMLElement | null>(null);
  const [rulerEventsMount, setRulerEventsMount] = useState<HTMLElement | null>(null);
  const [verticalRulerEventsMount, setVerticalRulerEventsMount] = useState<HTMLElement | null>(null);
  const [rulerMarkerMount, setRulerMarkerMount] = useState<HTMLElement | null>(null);
  const [verticalRulerMarkerMount, setVerticalRulerMarkerMount] = useState<HTMLElement | null>(null);
  const [changeCaseMount, setChangeCaseMount] = useState<HTMLElement | null>(null);
  const [textDirectionMount, setTextDirectionMount] = useState<HTMLElement | null>(null);
  const [themeMount, setThemeMount] = useState<HTMLElement | null>(null);
  const [languageMount, setLanguageMount] = useState<HTMLElement | null>(null);
  const [dropCapMount, setDropCapMount] = useState<HTMLElement | null>(null);
  const [lineNumbersMount, setLineNumbersMount] = useState<HTMLElement | null>(null);
  const [columnsMount, setColumnsMount] = useState<HTMLElement | null>(null);
  const [pageSetupMount, setPageSetupMount] = useState<HTMLElement | null>(null);
  const [captionMount, setCaptionMount] = useState<HTMLElement | null>(null);
  const [captionNumberingMount, setCaptionNumberingMount] = useState<HTMLElement | null>(null);
  const [watermarkMount, setWatermarkMount] = useState<HTMLElement | null>(null);
  const [hyperlinkMount, setHyperlinkMount] = useState<HTMLElement | null>(null);
  const [symbolMount, setSymbolMount] = useState<HTMLElement | null>(null);
  const [bulletsMount, setBulletsMount] = useState<HTMLElement | null>(null);
  const [customizeListMount, setCustomizeListMount] = useState<HTMLElement | null>(null);
  const [shapeMenuMount, setShapeMenuMount] = useState<HTMLElement | null>(null);
  const [shadowMenuMount, setShadowMenuMount] = useState<HTMLElement | null>(null);
  const [wordartMount, setWordartMount] = useState<HTMLElement | null>(null);
  const [lettersArtTextMount, setLettersArtTextMount] = useState<HTMLElement | null>(null);
  const [wordCountMount, setWordCountMount] = useState<HTMLElement | null>(null);
  const [summarizeMount, setSummarizeMount] = useState<HTMLElement | null>(null);
  const [zoomMount, setZoomMount] = useState<HTMLElement | null>(null);
  const [hyphenationMount, setHyphenationMount] = useState<HTMLElement | null>(null);
  const [pasteSpecialMount, setPasteSpecialMount] = useState<HTMLElement | null>(null);
  const [macrosMount, setMacrosMount] = useState<HTMLElement | null>(null);
  const [optionsViewMount, setOptionsViewMount] = useState<HTMLElement | null>(null);
  const [optionsGeneralMount, setOptionsGeneralMount] = useState<HTMLElement | null>(null);
  const [optionsEditMount, setOptionsEditMount] = useState<HTMLElement | null>(null);
  const [optionsPrintMount, setOptionsPrintMount] = useState<HTMLElement | null>(null);
  const [optionsSaveMount, setOptionsSaveMount] = useState<HTMLElement | null>(null);
  const [optionsTrackMount, setOptionsTrackMount] = useState<HTMLElement | null>(null);
  const [optionsUserMount, setOptionsUserMount] = useState<HTMLElement | null>(null);
  const [optionsCompatMount, setOptionsCompatMount] = useState<HTMLElement | null>(null);
  const [optionsSecurityMount, setOptionsSecurityMount] = useState<HTMLElement | null>(null);
  const [confirmPasswordMount, setConfirmPasswordMount] = useState<HTMLElement | null>(null);
  const [optionsSpellingMount, setOptionsSpellingMount] = useState<HTMLElement | null>(null);
  const [customDictionaryMount, setCustomDictionaryMount] = useState<HTMLElement | null>(null);
  const [formatObjectColorsMount, setFormatObjectColorsMount] = useState<HTMLElement | null>(null);
  const [fillEffectsMount, setFillEffectsMount] = useState<HTMLElement | null>(null);
  const [formatObjectWebMount, setFormatObjectWebMount] = useState<HTMLElement | null>(null);
  const [formatObjectSizeMount, setFormatObjectSizeMount] = useState<HTMLElement | null>(null);
  const [formatObjectLayoutMount, setFormatObjectLayoutMount] = useState<HTMLElement | null>(null);
  const [formatObjectAdvancedMount, setFormatObjectAdvancedMount] = useState<HTMLElement | null>(null);
  const [formatObjectTextBoxMount, setFormatObjectTextBoxMount] = useState<HTMLElement | null>(null);
  const [formatObjectPictureMount, setFormatObjectPictureMount] = useState<HTMLElement | null>(null);
  const [autoCorrectExceptionsMount, setAutoCorrectExceptionsMount] = useState<HTMLElement | null>(null);
  const [fontMount, setFontMount] = useState<HTMLElement | null>(null);
  const [autoCorrectMainMount, setAutoCorrectMainMount] = useState<HTMLElement | null>(null);
  const [autoTextMount, setAutoTextMount] = useState<HTMLElement | null>(null);
  const [autoCorrectFormatMount, setAutoCorrectFormatMount] = useState<HTMLElement | null>(null);
  const [customizeMount, setCustomizeMount] = useState<HTMLElement | null>(null);
  const [drawingGridMount, setDrawingGridMount] = useState<HTMLElement | null>(null);
  const [autoFormatMount, setAutoFormatMount] = useState<HTMLElement | null>(null);
  const [autoFormatReviewMount, setAutoFormatReviewMount] = useState<HTMLElement | null>(null);
  const [envelopesMount, setEnvelopesMount] = useState<HTMLElement | null>(null);
  const [indexTablesMount, setIndexTablesMount] = useState<HTMLElement | null>(null);
  const [letterWizardMount, setLetterWizardMount] = useState<HTMLElement | null>(null);
  const [organizerMount, setOrganizerMount] = useState<HTMLElement | null>(null);
  const [templatesMount, setTemplatesMount] = useState<HTMLElement | null>(null);
  const [propertiesMount, setPropertiesMount] = useState<HTMLElement | null>(null);
  const [saveAsMount, setSaveAsMount] = useState<HTMLElement | null>(null);
  const [templatePickerMount, setTemplatePickerMount] = useState<HTMLElement | null>(null);
  const [printMount, setPrintMount] = useState<HTMLElement | null>(null);
  const [bookmarkMount, setBookmarkMount] = useState<HTMLElement | null>(null);
  const [objectMount, setObjectMount] = useState<HTMLElement | null>(null);
  const [equationMount, setEquationMount] = useState<HTMLElement | null>(null);
  const [breakMount, setBreakMount] = useState<HTMLElement | null>(null);
  const [dateTimeMount, setDateTimeMount] = useState<HTMLElement | null>(null);
  const [pageNumbersMount, setPageNumbersMount] = useState<HTMLElement | null>(null);
  const [pageNumberFormatMount, setPageNumberFormatMount] = useState<HTMLElement | null>(null);
  const [cellOperationMount, setCellOperationMount] = useState<HTMLElement | null>(null);
  const [insertTableMount, setInsertTableMount] = useState<HTMLElement | null>(null);
  const [splitCellsMount, setSplitCellsMount] = useState<HTMLElement | null>(null);
  const [textToTableMount, setTextToTableMount] = useState<HTMLElement | null>(null);
  const [tableToTextMount, setTableToTextMount] = useState<HTMLElement | null>(null);
  const [formulaMount, setFormulaMount] = useState<HTMLElement | null>(null);
  const [crossRefMount, setCrossRefMount] = useState<HTMLElement | null>(null);
  const [markEntryMount, setMarkEntryMount] = useState<HTMLElement | null>(null);
  const [fieldMount, setFieldMount] = useState<HTMLElement | null>(null);
  const [footnoteMount, setFootnoteMount] = useState<HTMLElement | null>(null);
  const [convertNotesMount, setConvertNotesMount] = useState<HTMLElement | null>(null);
  const [styleDialogMount, setStyleDialogMount] = useState<HTMLElement | null>(null);
  const [diagramMount, setDiagramMount] = useState<HTMLElement | null>(null);
  const [, refresh] = useState(0);
  const [dropHintActive, setDropHintActive] = useState(false);
  const L = runtime();
  const current = L?.panes?.task?.current;
  useEffect(() => {
    if (!L) return;
    setDropHintActive(!!L.app?.reactDropHintActive);
    return L.bus?.on('wordx-drop-hint', () => setDropHintActive(!!L.app?.reactDropHintActive));
  }, [L]);
  useEffect(() => { if (rulerCanvasMount && verticalRulerCanvasMount) L?.rulers?.draw?.(); }, [rulerCanvasMount, verticalRulerCanvasMount, L]);
  useEffect(() => L?.bus?.on('sel-done', () => { L.app?.painter?.release(); L.app?.highlighter?.release(); }), [L]);
  useEffect(() => {
    let target: HTMLElement | null = null;
    let off: (() => void) | undefined;
    let mapTarget: HTMLElement | null = null;
    let mapRoot: HTMLElement | null = null;
    let hiddenChildren = new Map<HTMLElement, string>();
    let reviewTarget: HTMLElement | null = null;
    let reviewRoot: HTMLElement | null = null;
    let hiddenReviewChildren = new Map<HTMLElement, string>();
    let readingTarget: HTMLElement | null = null;
    let readingRoot: HTMLElement | null = null;
    let taskHeadTarget: HTMLElement | null = null;
    let taskHeadLegacyDisplay = '';
    let taskHeadRoot: HTMLElement | null = null;
    let statusTarget: HTMLElement | null = null;
    let statusRoot: HTMLElement | null = null;
    const hiddenStatusChildren = new Map<HTMLElement, string>();
    let gotoTarget: HTMLElement | null = null;
    let findInputTarget: HTMLElement | null = null;
    let replaceInputTarget: HTMLElement | null = null;
    let findAuxTarget: HTMLElement | null = null;
    let findOptionsTarget: HTMLElement | null = null;
    let findFormatTarget: HTMLElement | null = null;
    let findFormatSummaryTarget: HTMLElement | null = null;
    let replaceFormatSummaryTarget: HTMLElement | null = null;
    let findMoreTarget: HTMLElement | null = null;
    let findActionsTarget: HTMLElement | null = null;
    let sortTarget: HTMLElement | null = null;
    let mailMergeRecipientsTarget: HTMLElement | null = null;
    let newAddressListTarget: HTMLElement | null = null;
    let mergeFieldTarget: HTMLElement | null = null;
    let reviewersTarget: HTMLElement | null = null;
    let commentEditTarget: HTMLElement | null = null;
    let contentControlTarget: HTMLElement | null = null;
    let spellCheckTarget: HTMLElement | null = null;
    let tableAutoFormatTarget: HTMLElement | null = null;
    let tabsDialogTarget: HTMLElement | null = null;
    let paragraphDialogTarget: HTMLElement | null = null;
    let bordersDialogTarget: HTMLElement | null = null;
    let tablePropsDialogTarget: HTMLElement | null = null;
    let rulerCornerTarget: HTMLElement | null = null;
    let rulerCanvasTarget: HTMLElement | null = null;
    let verticalRulerCanvasTarget: HTMLElement | null = null;
    let rulerEventsTarget: HTMLElement | null = null;
    let verticalRulerEventsTarget: HTMLElement | null = null;
    let rulerMarkerTarget: HTMLElement | null = null;
    let verticalRulerMarkerTarget: HTMLElement | null = null;
    let changeCaseTarget: HTMLElement | null = null;
    let textDirectionTarget: HTMLElement | null = null;
    let themeTarget: HTMLElement | null = null;
    let languageTarget: HTMLElement | null = null;
    let dropCapTarget: HTMLElement | null = null;
    let lineNumbersTarget: HTMLElement | null = null;
    let columnsTarget: HTMLElement | null = null;
    let pageSetupTarget: HTMLElement | null = null;
    let captionTarget: HTMLElement | null = null;
    let captionNumberingTarget: HTMLElement | null = null;
    let watermarkTarget: HTMLElement | null = null;
    let hyperlinkTarget: HTMLElement | null = null;
    let symbolTarget: HTMLElement | null = null;
    let bulletsTarget: HTMLElement | null = null;
    let customizeListTarget: HTMLElement | null = null;
    let shapeMenuTarget: HTMLElement | null = null;
    let shadowMenuTarget: HTMLElement | null = null;
    let wordartTarget: HTMLElement | null = null;
    let lettersArtTextTarget: HTMLElement | null = null;
    let wordCountTarget: HTMLElement | null = null;
    let summarizeTarget: HTMLElement | null = null;
    let zoomTarget: HTMLElement | null = null;
    let hyphenationTarget: HTMLElement | null = null;
    let pasteSpecialTarget: HTMLElement | null = null;
    let macrosTarget: HTMLElement | null = null;
    let optionsViewTarget: HTMLElement | null = null;
    let optionsGeneralTarget: HTMLElement | null = null;
    let optionsEditTarget: HTMLElement | null = null;
    let optionsPrintTarget: HTMLElement | null = null;
    let optionsSaveTarget: HTMLElement | null = null;
    let optionsTrackTarget: HTMLElement | null = null;
    let optionsUserTarget: HTMLElement | null = null;
    let optionsCompatTarget: HTMLElement | null = null;
    let optionsSecurityTarget: HTMLElement | null = null;
    let confirmPasswordTarget: HTMLElement | null = null;
    let optionsSpellingTarget: HTMLElement | null = null;
    let customDictionaryTarget: HTMLElement | null = null;
    let formatObjectWebTarget: HTMLElement | null = null;
    let formatObjectColorsTarget: HTMLElement | null = null;
    let fillEffectsTarget: HTMLElement | null = null;
    let formatObjectSizeTarget: HTMLElement | null = null;
    let formatObjectLayoutTarget: HTMLElement | null = null;
    let formatObjectAdvancedTarget: HTMLElement | null = null;
    let formatObjectTextBoxTarget: HTMLElement | null = null;
    let formatObjectPictureTarget: HTMLElement | null = null;
    let autoCorrectExceptionsTarget: HTMLElement | null = null;
    let fontTarget: HTMLElement | null = null;
    let autoCorrectMainTarget: HTMLElement | null = null;
    let autoTextTarget: HTMLElement | null = null;
    let autoCorrectFormatTarget: HTMLElement | null = null;
    let customizeTarget: HTMLElement | null = null;
    let drawingGridTarget: HTMLElement | null = null;
    let autoFormatTarget: HTMLElement | null = null;
    let autoFormatReviewTarget: HTMLElement | null = null;
    let envelopesTarget: HTMLElement | null = null;
    let indexTablesTarget: HTMLElement | null = null;
    let letterWizardTarget: HTMLElement | null = null;
    let organizerTarget: HTMLElement | null = null;
    let templatesTarget: HTMLElement | null = null;
    let propertiesTarget: HTMLElement | null = null;
    let saveAsTarget: HTMLElement | null = null;
    let templatePickerTarget: HTMLElement | null = null;
    let printTarget: HTMLElement | null = null;
    let bookmarkTarget: HTMLElement | null = null;
    let objectTarget: HTMLElement | null = null;
    let equationTarget: HTMLElement | null = null;
    let breakTarget: HTMLElement | null = null;
    let dateTimeTarget: HTMLElement | null = null;
    let pageNumbersTarget: HTMLElement | null = null;
    let pageNumberFormatTarget: HTMLElement | null = null;
    let cellOperationTarget: HTMLElement | null = null;
    let insertTableTarget: HTMLElement | null = null;
    let splitCellsTarget: HTMLElement | null = null;
    let textToTableTarget: HTMLElement | null = null;
    let tableToTextTarget: HTMLElement | null = null;
    let formulaTarget: HTMLElement | null = null;
    let crossRefTarget: HTMLElement | null = null;
    let markEntryTarget: HTMLElement | null = null;
    let fieldTarget: HTMLElement | null = null;
    let footnoteTarget: HTMLElement | null = null;
    let convertNotesTarget: HTMLElement | null = null;
    let styleDialogTarget: HTMLElement | null = null;
    let diagramTarget: HTMLElement | null = null;
    const restoreMap = () => {
      mapRoot?.remove(); mapRoot = null;
      hiddenChildren.forEach((display, child) => { child.style.display = display; }); hiddenChildren.clear();
      mapTarget?.classList.remove('vibeoffice-react-docmap'); mapTarget = null; setMapMount(null);
    };
    const restoreReview = () => {
      reviewRoot?.remove(); reviewRoot = null;
      hiddenReviewChildren.forEach((display, child) => { child.style.display = display; }); hiddenReviewChildren.clear();
      reviewTarget = null; setReviewMount(null);
    };
    const attach = () => {
      const engine = runtime();
      const next = host.querySelector<HTMLElement>('#react-wordx-task-body');
      if (next !== target) {
        target?.classList.remove('vibeoffice-react-owned');
        target?.style.removeProperty('display');
        target = next;
        if (target) {
          target.style.display = 'contents';
          target.classList.add('vibeoffice-react-owned');
        }
        setMount(target);
      }
      const nextMap = host.querySelector<HTMLElement>('#docmap');
      if (nextMap !== mapTarget) {
        restoreMap();
        mapTarget = nextMap;
        if (mapTarget) {
          mapTarget.classList.add('vibeoffice-react-docmap');
          for (const child of Array.from(mapTarget.children) as HTMLElement[]) { hiddenChildren.set(child, child.style.display); child.style.display = 'none'; }
          mapRoot = document.createElement('div'); mapRoot.id = 'react-wordx-docmap'; mapTarget.append(mapRoot); setMapMount(mapRoot);
        }
      }
      const nextReview = host.querySelector<HTMLElement>('#revpane');
      if (nextReview !== reviewTarget) {
        restoreReview();
        reviewTarget = nextReview;
        if (reviewTarget) {
          for (const child of Array.from(reviewTarget.children) as HTMLElement[]) { hiddenReviewChildren.set(child, child.style.display); child.style.display = 'none'; }
          reviewRoot = document.createElement('div'); reviewRoot.id = 'react-wordx-reviewpane'; reviewTarget.append(reviewRoot); setReviewMount(reviewRoot);
        }
      }
      const nextReading = host.querySelector<HTMLElement>('#toolbars');
      if (nextReading !== readingTarget) {
        readingRoot?.remove(); readingRoot = null;
        readingTarget = nextReading;
        if (readingTarget) {
          readingRoot = document.createElement('div'); readingRoot.id = 'react-wordx-reading-toolbar'; readingTarget.append(readingRoot); setReadingMount(readingRoot);
        } else setReadingMount(null);
      }
      const nextTaskHead = host.querySelector<HTMLElement>('#taskpane > .tp-head');
      if (nextTaskHead !== taskHeadTarget) {
        taskHeadRoot?.remove(); taskHeadRoot = null;
        if (taskHeadTarget) taskHeadTarget.style.display = taskHeadLegacyDisplay;
        taskHeadTarget = nextTaskHead;
        if (taskHeadTarget) {
          taskHeadLegacyDisplay = taskHeadTarget.style.display;
          taskHeadTarget.style.display = 'none';
          taskHeadRoot = document.createElement('div'); taskHeadRoot.id = 'react-wordx-task-head'; taskHeadRoot.className = 'react-task-head-root vibeoffice-react-owned'; host.querySelector('#taskpane')?.append(taskHeadRoot); setTaskHeadMount(taskHeadRoot);
        } else setTaskHeadMount(null);
      }
      const nextStatus = host.querySelector<HTMLElement>('#statusbar');
      if (nextStatus !== statusTarget) {
        statusRoot?.remove(); statusRoot = null;
        hiddenStatusChildren.forEach((display, child) => { child.style.display = display; }); hiddenStatusChildren.clear();
        statusTarget = nextStatus;
        if (statusTarget) {
          statusRoot = document.createElement('div'); statusRoot.id = 'react-wordx-statusbar'; statusTarget.insertBefore(statusRoot, statusTarget.querySelector('[data-react-status]')); setStatusMount(statusRoot);
        } else setStatusMount(null);
      }
      if (statusTarget) for (const child of Array.from(statusTarget.children) as HTMLElement[]) if (child !== statusRoot && !child.hasAttribute('data-react-status') && !hiddenStatusChildren.has(child)) { hiddenStatusChildren.set(child, child.style.display); child.style.display = 'none'; }
      const nextGoto = host.querySelector<HTMLElement>('#react-wordx-goto-root');
      if (nextGoto !== gotoTarget) { gotoTarget = nextGoto; setGotoMount(gotoTarget); }
      const nextFindInput = host.querySelector<HTMLElement>('#react-wordx-find-input-root');
      if (nextFindInput !== findInputTarget) { findInputTarget = nextFindInput; setFindInputMount(findInputTarget); }
      const nextReplaceInput = host.querySelector<HTMLElement>('#react-wordx-replace-input-root');
      if (nextReplaceInput !== replaceInputTarget) { replaceInputTarget = nextReplaceInput; setReplaceInputMount(replaceInputTarget); }
      const nextFindAux = host.querySelector<HTMLElement>('#react-wordx-find-aux-root');
      if (nextFindAux !== findAuxTarget) { findAuxTarget = nextFindAux; setFindAuxMount(findAuxTarget); }
      const nextFindOptions = host.querySelector<HTMLElement>('#react-wordx-find-options-root');
      if (nextFindOptions !== findOptionsTarget) { findOptionsTarget = nextFindOptions; setFindOptionsMount(findOptionsTarget); }
      const nextFindFormat = host.querySelector<HTMLElement>('#react-wordx-find-format-root');
      if (nextFindFormat !== findFormatTarget) { findFormatTarget = nextFindFormat; setFindFormatMount(findFormatTarget); }
      const nextFindFormatSummary = host.querySelector<HTMLElement>('#react-wordx-find-format-summary-root');
      if (nextFindFormatSummary !== findFormatSummaryTarget) { findFormatSummaryTarget = nextFindFormatSummary; setFindFormatSummaryMount(findFormatSummaryTarget); }
      const nextReplaceFormatSummary = host.querySelector<HTMLElement>('#react-wordx-replace-format-summary-root');
      if (nextReplaceFormatSummary !== replaceFormatSummaryTarget) { replaceFormatSummaryTarget = nextReplaceFormatSummary; setReplaceFormatSummaryMount(replaceFormatSummaryTarget); }
      const nextFindMore = host.querySelector<HTMLElement>('#react-wordx-find-more-root');
      if (nextFindMore !== findMoreTarget) { findMoreTarget = nextFindMore; setFindMoreMount(findMoreTarget); }
      const nextFindActions = host.querySelector<HTMLElement>('#react-wordx-find-actions-root');
      if (nextFindActions !== findActionsTarget) { findActionsTarget = nextFindActions; setFindActionsMount(findActionsTarget); }
      const nextSort = host.querySelector<HTMLElement>('#react-wordx-sort-root');
      if (nextSort !== sortTarget) { sortTarget = nextSort; setSortMount(sortTarget); }
      const nextMailMergeRecipients = host.querySelector<HTMLElement>('#react-wordx-mail-merge-recipients-root');
      if (nextMailMergeRecipients !== mailMergeRecipientsTarget) { mailMergeRecipientsTarget = nextMailMergeRecipients; setMailMergeRecipientsMount(mailMergeRecipientsTarget); }
      const nextNewAddressList = host.querySelector<HTMLElement>('#react-wordx-new-address-list-root');
      if (nextNewAddressList !== newAddressListTarget) { newAddressListTarget = nextNewAddressList; setNewAddressListMount(newAddressListTarget); }
      const nextMergeField = host.querySelector<HTMLElement>('#react-wordx-mail-merge-field-root');
      if (nextMergeField !== mergeFieldTarget) { mergeFieldTarget = nextMergeField; setMergeFieldMount(mergeFieldTarget); }
      const nextReviewers = host.querySelector<HTMLElement>('#react-wordx-reviewers-root');
      if (nextReviewers !== reviewersTarget) { reviewersTarget = nextReviewers; setReviewersMount(reviewersTarget); }
      const nextCommentEdit = host.querySelector<HTMLElement>('#react-wordx-comment-edit-root');
      if (nextCommentEdit !== commentEditTarget) { commentEditTarget = nextCommentEdit; setCommentEditMount(commentEditTarget); }
      const nextContentControl = host.querySelector<HTMLElement>('#react-wordx-content-control-root');
      if (nextContentControl !== contentControlTarget) { contentControlTarget = nextContentControl; setContentControlMount(contentControlTarget); }
      const nextSpellCheck = host.querySelector<HTMLElement>('#react-wordx-spell-check-root');
      if (nextSpellCheck !== spellCheckTarget) { spellCheckTarget = nextSpellCheck; setSpellCheckMount(spellCheckTarget); }
      const nextTableAutoFormat = host.querySelector<HTMLElement>('#react-wordx-table-autoformat-root');
      if (nextTableAutoFormat !== tableAutoFormatTarget) { tableAutoFormatTarget = nextTableAutoFormat; setTableAutoFormatMount(tableAutoFormatTarget); }
      const nextTabsDialog = host.querySelector<HTMLElement>('#react-wordx-tabs-root');
      if (nextTabsDialog !== tabsDialogTarget) { tabsDialogTarget = nextTabsDialog; setTabsDialogMount(tabsDialogTarget); }
      const nextParagraphDialog = host.querySelector<HTMLElement>('#react-wordx-paragraph-root');
      if (nextParagraphDialog !== paragraphDialogTarget) { paragraphDialogTarget = nextParagraphDialog; setParagraphDialogMount(paragraphDialogTarget); }
      const nextBordersDialog = host.querySelector<HTMLElement>('#react-wordx-borders-root');
      if (nextBordersDialog !== bordersDialogTarget) { bordersDialogTarget = nextBordersDialog; setBordersDialogMount(bordersDialogTarget); }
      const nextTablePropsDialog = host.querySelector<HTMLElement>('#react-wordx-table-properties-root');
      if (nextTablePropsDialog !== tablePropsDialogTarget) { tablePropsDialogTarget = nextTablePropsDialog; setTablePropsDialogMount(tablePropsDialogTarget); }
      const nextRulerCorner = host.querySelector<HTMLElement>('#rcorner');
      if (nextRulerCorner !== rulerCornerTarget) { rulerCornerTarget = nextRulerCorner; setRulerCornerMount(rulerCornerTarget); }
      const nextRulerCanvas = host.querySelector<HTMLElement>('#react-wordx-ruler-canvas-root');
      if (nextRulerCanvas !== rulerCanvasTarget) { rulerCanvasTarget = nextRulerCanvas; setRulerCanvasMount(rulerCanvasTarget); }
      const nextVerticalRulerCanvas = host.querySelector<HTMLElement>('#react-wordx-v-ruler-canvas-root');
      if (nextVerticalRulerCanvas !== verticalRulerCanvasTarget) { verticalRulerCanvasTarget = nextVerticalRulerCanvas; setVerticalRulerCanvasMount(verticalRulerCanvasTarget); }
      const nextRulerEvents = host.querySelector<HTMLElement>('#react-wordx-ruler-events-root');
      if (nextRulerEvents !== rulerEventsTarget) { rulerEventsTarget = nextRulerEvents; setRulerEventsMount(rulerEventsTarget); }
      const nextVerticalRulerEvents = host.querySelector<HTMLElement>('#react-wordx-v-ruler-events-root');
      if (nextVerticalRulerEvents !== verticalRulerEventsTarget) { verticalRulerEventsTarget = nextVerticalRulerEvents; setVerticalRulerEventsMount(verticalRulerEventsTarget); }
      const nextRulerMarkers = host.querySelector<HTMLElement>('#react-wordx-ruler-markers-root');
      if (nextRulerMarkers !== rulerMarkerTarget) { rulerMarkerTarget = nextRulerMarkers; setRulerMarkerMount(rulerMarkerTarget); }
      const nextVerticalRulerMarkers = host.querySelector<HTMLElement>('#react-wordx-v-ruler-markers-root');
      if (nextVerticalRulerMarkers !== verticalRulerMarkerTarget) { verticalRulerMarkerTarget = nextVerticalRulerMarkers; setVerticalRulerMarkerMount(verticalRulerMarkerTarget); }
      const nextChangeCase = host.querySelector<HTMLElement>('#react-wordx-change-case-root');
      if (nextChangeCase !== changeCaseTarget) { changeCaseTarget = nextChangeCase; setChangeCaseMount(changeCaseTarget); }
      const nextTextDirection = host.querySelector<HTMLElement>('#react-wordx-text-direction-root');
      if (nextTextDirection !== textDirectionTarget) { textDirectionTarget = nextTextDirection; setTextDirectionMount(textDirectionTarget); }
      const nextTheme = host.querySelector<HTMLElement>('#react-wordx-theme-root');
      if (nextTheme !== themeTarget) { themeTarget = nextTheme; setThemeMount(themeTarget); }
      const nextLanguage = host.querySelector<HTMLElement>('#react-wordx-language-root');
      if (nextLanguage !== languageTarget) { languageTarget = nextLanguage; setLanguageMount(languageTarget); }
      const nextDropCap = host.querySelector<HTMLElement>('#react-wordx-drop-cap-root');
      if (nextDropCap !== dropCapTarget) { dropCapTarget = nextDropCap; setDropCapMount(dropCapTarget); }
      const nextLineNumbers = host.querySelector<HTMLElement>('#react-wordx-line-numbers-root');
      if (nextLineNumbers !== lineNumbersTarget) { lineNumbersTarget = nextLineNumbers; setLineNumbersMount(lineNumbersTarget); }
      const nextColumns = host.querySelector<HTMLElement>('#react-wordx-columns-root');
      if (nextColumns !== columnsTarget) { columnsTarget = nextColumns; setColumnsMount(columnsTarget); }
      const nextPageSetup = host.querySelector<HTMLElement>('#react-wordx-page-setup-root');
      if (nextPageSetup !== pageSetupTarget) { pageSetupTarget = nextPageSetup; setPageSetupMount(pageSetupTarget); }
      const nextCaption = host.querySelector<HTMLElement>('#react-wordx-caption-root');
      if (nextCaption !== captionTarget) { captionTarget = nextCaption; setCaptionMount(captionTarget); }
      const nextCaptionNumbering = host.querySelector<HTMLElement>('#react-wordx-caption-numbering-root');
      if (nextCaptionNumbering !== captionNumberingTarget) { captionNumberingTarget = nextCaptionNumbering; setCaptionNumberingMount(captionNumberingTarget); }
      const nextWatermark = host.querySelector<HTMLElement>('#react-wordx-watermark-root');
      if (nextWatermark !== watermarkTarget) { watermarkTarget = nextWatermark; setWatermarkMount(watermarkTarget); }
      const nextHyperlink = host.querySelector<HTMLElement>('#react-wordx-hyperlink-root');
      if (nextHyperlink !== hyperlinkTarget) { hyperlinkTarget = nextHyperlink; setHyperlinkMount(hyperlinkTarget); }
      const nextSymbol = host.querySelector<HTMLElement>('#react-wordx-symbol-root');
      if (nextSymbol !== symbolTarget) { symbolTarget = nextSymbol; setSymbolMount(symbolTarget); }
      const nextBullets = host.querySelector<HTMLElement>('#react-wordx-bullets-root');
      if (nextBullets !== bulletsTarget) { bulletsTarget = nextBullets; setBulletsMount(bulletsTarget); }
      const nextCustomizeList = host.querySelector<HTMLElement>('#react-wordx-customize-list-root');
      if (nextCustomizeList !== customizeListTarget) { customizeListTarget = nextCustomizeList; setCustomizeListMount(customizeListTarget); }
      const nextShapeMenu = host.querySelector<HTMLElement>('#react-wordx-shape-menu-root');
      if (nextShapeMenu !== shapeMenuTarget) { shapeMenuTarget = nextShapeMenu; setShapeMenuMount(shapeMenuTarget); }
      const nextShadowMenu = host.querySelector<HTMLElement>('#react-wordx-shadow-menu-root');
      if (nextShadowMenu !== shadowMenuTarget) { shadowMenuTarget = nextShadowMenu; setShadowMenuMount(shadowMenuTarget); }
      const nextWordart = host.querySelector<HTMLElement>('#react-wordx-wordart-root');
      if (nextWordart !== wordartTarget) { wordartTarget = nextWordart; setWordartMount(wordartTarget); }
      const nextLettersArtText = host.querySelector<HTMLElement>('#react-wordx-lettersart-text-root');
      if (nextLettersArtText !== lettersArtTextTarget) { lettersArtTextTarget = nextLettersArtText; setLettersArtTextMount(lettersArtTextTarget); }
      const nextWordCount = host.querySelector<HTMLElement>('#react-wordx-word-count-root');
      if (nextWordCount !== wordCountTarget) { wordCountTarget = nextWordCount; setWordCountMount(wordCountTarget); }
      const nextSummarize = host.querySelector<HTMLElement>('#react-wordx-summarize-root');
      if (nextSummarize !== summarizeTarget) { summarizeTarget = nextSummarize; setSummarizeMount(summarizeTarget); }
      const nextZoom = host.querySelector<HTMLElement>('#react-wordx-zoom-root');
      if (nextZoom !== zoomTarget) { zoomTarget = nextZoom; setZoomMount(zoomTarget); }
      const nextHyphenation = host.querySelector<HTMLElement>('#react-wordx-hyphenation-root');
      if (nextHyphenation !== hyphenationTarget) { hyphenationTarget = nextHyphenation; setHyphenationMount(hyphenationTarget); }
      const nextPasteSpecial = host.querySelector<HTMLElement>('#react-wordx-paste-special-root');
      if (nextPasteSpecial !== pasteSpecialTarget) { pasteSpecialTarget = nextPasteSpecial; setPasteSpecialMount(pasteSpecialTarget); }
      const nextMacros = host.querySelector<HTMLElement>('#react-wordx-macros-root');
      if (nextMacros !== macrosTarget) { macrosTarget = nextMacros; setMacrosMount(macrosTarget); }
      const nextOptionsView = host.querySelector<HTMLElement>('#react-wordx-options-view-root');
      if (nextOptionsView !== optionsViewTarget) { optionsViewTarget = nextOptionsView; setOptionsViewMount(optionsViewTarget); }
      const nextOptionsGeneral = host.querySelector<HTMLElement>('#react-wordx-options-general-root');
      if (nextOptionsGeneral !== optionsGeneralTarget) { optionsGeneralTarget = nextOptionsGeneral; setOptionsGeneralMount(optionsGeneralTarget); }
      const nextOptionsEdit = host.querySelector<HTMLElement>('#react-wordx-options-edit-root');
      if (nextOptionsEdit !== optionsEditTarget) { optionsEditTarget = nextOptionsEdit; setOptionsEditMount(optionsEditTarget); }
      const nextOptionsPrint = host.querySelector<HTMLElement>('#react-wordx-options-print-root');
      if (nextOptionsPrint !== optionsPrintTarget) { optionsPrintTarget = nextOptionsPrint; setOptionsPrintMount(optionsPrintTarget); }
      const nextOptionsSave = host.querySelector<HTMLElement>('#react-wordx-options-save-root');
      if (nextOptionsSave !== optionsSaveTarget) { optionsSaveTarget = nextOptionsSave; setOptionsSaveMount(optionsSaveTarget); }
      const nextOptionsTrack = host.querySelector<HTMLElement>('#react-wordx-options-track-root');
      if (nextOptionsTrack !== optionsTrackTarget) { optionsTrackTarget = nextOptionsTrack; setOptionsTrackMount(optionsTrackTarget); }
      const nextOptionsUser = host.querySelector<HTMLElement>('#react-wordx-options-user-root');
      if (nextOptionsUser !== optionsUserTarget) { optionsUserTarget = nextOptionsUser; setOptionsUserMount(optionsUserTarget); }
      const nextOptionsCompat = host.querySelector<HTMLElement>('#react-wordx-options-compat-root');
      if (nextOptionsCompat !== optionsCompatTarget) { optionsCompatTarget = nextOptionsCompat; setOptionsCompatMount(optionsCompatTarget); }
      const nextOptionsSecurity = host.querySelector<HTMLElement>('#react-wordx-options-security-root');
      if (nextOptionsSecurity !== optionsSecurityTarget) { optionsSecurityTarget = nextOptionsSecurity; setOptionsSecurityMount(optionsSecurityTarget); }
      const nextConfirmPassword = host.querySelector<HTMLElement>('#react-wordx-confirm-password-root');
      if (nextConfirmPassword !== confirmPasswordTarget) { confirmPasswordTarget = nextConfirmPassword; setConfirmPasswordMount(confirmPasswordTarget); }
      const nextOptionsSpelling = host.querySelector<HTMLElement>('#react-wordx-options-spelling-root');
      if (nextOptionsSpelling !== optionsSpellingTarget) { optionsSpellingTarget = nextOptionsSpelling; setOptionsSpellingMount(optionsSpellingTarget); }
      const nextCustomDictionary = host.querySelector<HTMLElement>('#react-wordx-custom-dictionary-root');
      if (nextCustomDictionary !== customDictionaryTarget) { customDictionaryTarget = nextCustomDictionary; setCustomDictionaryMount(customDictionaryTarget); }
      const nextFormatObjectWeb = host.querySelector<HTMLElement>('#react-wordx-format-object-web-root');
      if (nextFormatObjectWeb !== formatObjectWebTarget) { formatObjectWebTarget = nextFormatObjectWeb; setFormatObjectWebMount(formatObjectWebTarget); }
      const nextFormatObjectColors = host.querySelector<HTMLElement>('#react-wordx-format-object-colors-root');
      if (nextFormatObjectColors !== formatObjectColorsTarget) { formatObjectColorsTarget = nextFormatObjectColors; setFormatObjectColorsMount(formatObjectColorsTarget); }
      const nextFillEffects = host.querySelector<HTMLElement>('#react-wordx-fill-effects-root');
      if (nextFillEffects !== fillEffectsTarget) { fillEffectsTarget = nextFillEffects; setFillEffectsMount(fillEffectsTarget); }
      const nextFormatObjectSize = host.querySelector<HTMLElement>('#react-wordx-format-object-size-root');
      if (nextFormatObjectSize !== formatObjectSizeTarget) { formatObjectSizeTarget = nextFormatObjectSize; setFormatObjectSizeMount(formatObjectSizeTarget); }
      const nextFormatObjectLayout = host.querySelector<HTMLElement>('#react-wordx-format-object-layout-root');
      if (nextFormatObjectLayout !== formatObjectLayoutTarget) { formatObjectLayoutTarget = nextFormatObjectLayout; setFormatObjectLayoutMount(formatObjectLayoutTarget); }
      const nextFormatObjectAdvanced = host.querySelector<HTMLElement>('#react-wordx-format-object-advanced-root');
      if (nextFormatObjectAdvanced !== formatObjectAdvancedTarget) { formatObjectAdvancedTarget = nextFormatObjectAdvanced; setFormatObjectAdvancedMount(formatObjectAdvancedTarget); }
      const nextFormatObjectTextBox = host.querySelector<HTMLElement>('#react-wordx-format-object-textbox-root');
      if (nextFormatObjectTextBox !== formatObjectTextBoxTarget) { formatObjectTextBoxTarget = nextFormatObjectTextBox; setFormatObjectTextBoxMount(formatObjectTextBoxTarget); }
      const nextFormatObjectPicture = host.querySelector<HTMLElement>('#react-wordx-format-object-picture-root');
      if (nextFormatObjectPicture !== formatObjectPictureTarget) { formatObjectPictureTarget = nextFormatObjectPicture; setFormatObjectPictureMount(formatObjectPictureTarget); }
      const nextAutoCorrectExceptions = host.querySelector<HTMLElement>('#react-wordx-autocorrect-exceptions-root');
      if (nextAutoCorrectExceptions !== autoCorrectExceptionsTarget) { autoCorrectExceptionsTarget = nextAutoCorrectExceptions; setAutoCorrectExceptionsMount(autoCorrectExceptionsTarget); }
      const nextFont = host.querySelector<HTMLElement>('#react-wordx-font-root');
      if (nextFont !== fontTarget) { fontTarget = nextFont; setFontMount(fontTarget); }
      const nextAutoCorrectMain = host.querySelector<HTMLElement>('#react-wordx-autocorrect-main-root');
      if (nextAutoCorrectMain !== autoCorrectMainTarget) { autoCorrectMainTarget = nextAutoCorrectMain; setAutoCorrectMainMount(autoCorrectMainTarget); }
      const nextAutoText = host.querySelector<HTMLElement>('#react-wordx-autotext-root');
      if (nextAutoText !== autoTextTarget) { autoTextTarget = nextAutoText; setAutoTextMount(autoTextTarget); }
      const nextAutoCorrectFormat = host.querySelector<HTMLElement>('#react-wordx-autocorrect-format-root');
      if (nextAutoCorrectFormat !== autoCorrectFormatTarget) { autoCorrectFormatTarget = nextAutoCorrectFormat; setAutoCorrectFormatMount(autoCorrectFormatTarget); }
      const nextCustomize = host.querySelector<HTMLElement>('#react-wordx-customize-root');
      if (nextCustomize !== customizeTarget) { customizeTarget = nextCustomize; setCustomizeMount(customizeTarget); }
      const nextDrawingGrid = host.querySelector<HTMLElement>('#react-wordx-drawing-grid-root');
      if (nextDrawingGrid !== drawingGridTarget) { drawingGridTarget = nextDrawingGrid; setDrawingGridMount(drawingGridTarget); }
      const nextAutoFormat = host.querySelector<HTMLElement>('#react-wordx-autoformat-root');
      if (nextAutoFormat !== autoFormatTarget) { autoFormatTarget = nextAutoFormat; setAutoFormatMount(autoFormatTarget); }
      const nextAutoFormatReview = host.querySelector<HTMLElement>('#react-wordx-autoformat-review-root');
      if (nextAutoFormatReview !== autoFormatReviewTarget) { autoFormatReviewTarget = nextAutoFormatReview; setAutoFormatReviewMount(autoFormatReviewTarget); }
      const nextEnvelopes = host.querySelector<HTMLElement>('#react-wordx-envelopes-root');
      if (nextEnvelopes !== envelopesTarget) { envelopesTarget = nextEnvelopes; setEnvelopesMount(envelopesTarget); }
      const nextIndexTables = host.querySelector<HTMLElement>('#react-wordx-index-tables-root');
      if (nextIndexTables !== indexTablesTarget) { indexTablesTarget = nextIndexTables; setIndexTablesMount(indexTablesTarget); }
      const nextLetterWizard = host.querySelector<HTMLElement>('#react-wordx-letter-wizard-root');
      if (nextLetterWizard !== letterWizardTarget) { letterWizardTarget = nextLetterWizard; setLetterWizardMount(letterWizardTarget); }
      const nextOrganizer = host.querySelector<HTMLElement>('#react-wordx-organizer-root');
      if (nextOrganizer !== organizerTarget) { organizerTarget = nextOrganizer; setOrganizerMount(organizerTarget); }
      const nextTemplates = host.querySelector<HTMLElement>('#react-wordx-templates-root');
      if (nextTemplates !== templatesTarget) { templatesTarget = nextTemplates; setTemplatesMount(templatesTarget); }
      const nextProperties = host.querySelector<HTMLElement>('#react-wordx-properties-root');
      if (nextProperties !== propertiesTarget) { propertiesTarget = nextProperties; setPropertiesMount(propertiesTarget); }
      const nextSaveAs = host.querySelector<HTMLElement>('#react-wordx-save-as-root');
      if (nextSaveAs !== saveAsTarget) { saveAsTarget = nextSaveAs; setSaveAsMount(saveAsTarget); }
      const nextTemplatePicker = host.querySelector<HTMLElement>('#react-wordx-template-picker-root');
      if (nextTemplatePicker !== templatePickerTarget) { templatePickerTarget = nextTemplatePicker; setTemplatePickerMount(templatePickerTarget); }
      const nextPrint = host.querySelector<HTMLElement>('#react-wordx-print-root');
      if (nextPrint !== printTarget) { printTarget = nextPrint; setPrintMount(printTarget); }
      const nextBookmark = host.querySelector<HTMLElement>('#react-wordx-bookmark-root');
      if (nextBookmark !== bookmarkTarget) { bookmarkTarget = nextBookmark; setBookmarkMount(bookmarkTarget); }
      const nextObject = host.querySelector<HTMLElement>('#react-wordx-object-root');
      if (nextObject !== objectTarget) { objectTarget = nextObject; setObjectMount(objectTarget); }
      const nextEquation = host.querySelector<HTMLElement>('#react-wordx-equation-root');
      if (nextEquation !== equationTarget) { equationTarget = nextEquation; setEquationMount(equationTarget); }
      const nextBreak = host.querySelector<HTMLElement>('#react-wordx-break-root');
      if (nextBreak !== breakTarget) { breakTarget = nextBreak; setBreakMount(breakTarget); }
      const nextDateTime = host.querySelector<HTMLElement>('#react-wordx-date-time-root');
      if (nextDateTime !== dateTimeTarget) { dateTimeTarget = nextDateTime; setDateTimeMount(dateTimeTarget); }
      const nextPageNumbers = host.querySelector<HTMLElement>('#react-wordx-page-numbers-root');
      if (nextPageNumbers !== pageNumbersTarget) { pageNumbersTarget = nextPageNumbers; setPageNumbersMount(pageNumbersTarget); }
      const nextPageNumberFormat = host.querySelector<HTMLElement>('#react-wordx-page-number-format-root');
      if (nextPageNumberFormat !== pageNumberFormatTarget) { pageNumberFormatTarget = nextPageNumberFormat; setPageNumberFormatMount(pageNumberFormatTarget); }
      const nextCellOperation = host.querySelector<HTMLElement>('#react-wordx-cell-operation-root');
      if (nextCellOperation !== cellOperationTarget) { cellOperationTarget = nextCellOperation; setCellOperationMount(cellOperationTarget); }
      const nextInsertTable = host.querySelector<HTMLElement>('#react-wordx-insert-table-root');
      if (nextInsertTable !== insertTableTarget) { insertTableTarget = nextInsertTable; setInsertTableMount(insertTableTarget); }
      const nextSplitCells = host.querySelector<HTMLElement>('#react-wordx-split-cells-root');
      if (nextSplitCells !== splitCellsTarget) { splitCellsTarget = nextSplitCells; setSplitCellsMount(splitCellsTarget); }
      const nextTextToTable = host.querySelector<HTMLElement>('#react-wordx-text-to-table-root');
      if (nextTextToTable !== textToTableTarget) { textToTableTarget = nextTextToTable; setTextToTableMount(textToTableTarget); }
      const nextTableToText = host.querySelector<HTMLElement>('#react-wordx-table-to-text-root');
      if (nextTableToText !== tableToTextTarget) { tableToTextTarget = nextTableToText; setTableToTextMount(tableToTextTarget); }
      const nextFormula = host.querySelector<HTMLElement>('#react-wordx-formula-root');
      if (nextFormula !== formulaTarget) { formulaTarget = nextFormula; setFormulaMount(formulaTarget); }
      const nextCrossRef = host.querySelector<HTMLElement>('#react-wordx-cross-reference-root');
      if (nextCrossRef !== crossRefTarget) { crossRefTarget = nextCrossRef; setCrossRefMount(crossRefTarget); }
      const nextMarkEntry = host.querySelector<HTMLElement>('#react-wordx-mark-entry-root');
      if (nextMarkEntry !== markEntryTarget) { markEntryTarget = nextMarkEntry; setMarkEntryMount(markEntryTarget); }
      const nextField = host.querySelector<HTMLElement>('#react-wordx-field-root');
      if (nextField !== fieldTarget) { fieldTarget = nextField; setFieldMount(fieldTarget); }
      const nextFootnote = host.querySelector<HTMLElement>('#react-wordx-footnote-root');
      if (nextFootnote !== footnoteTarget) { footnoteTarget = nextFootnote; setFootnoteMount(footnoteTarget); }
      const nextConvertNotes = host.querySelector<HTMLElement>('#react-wordx-convert-notes-root');
      if (nextConvertNotes !== convertNotesTarget) { convertNotesTarget = nextConvertNotes; setConvertNotesMount(convertNotesTarget); }
      const nextStyleDialog = host.querySelector<HTMLElement>('#react-wordx-style-root');
      if (nextStyleDialog !== styleDialogTarget) { styleDialogTarget = nextStyleDialog; setStyleDialogMount(styleDialogTarget); }
      const nextDiagram = host.querySelector<HTMLElement>('#react-wordx-diagram-root');
      if (nextDiagram !== diagramTarget) { diagramTarget = nextDiagram; setDiagramMount(diagramTarget); }
      if (!off && engine?.bus) off = engine.bus.on('wordx-task-render', () => refresh((value) => value + 1));
    };
    const observer = new MutationObserver(attach);
    observer.observe(host, { childList: true, subtree: true });
    attach();
    return () => { observer.disconnect(); off?.(); target?.classList.remove('vibeoffice-react-owned'); target?.style.removeProperty('display'); restoreMap(); restoreReview(); readingRoot?.remove(); readingRoot = null; taskHeadRoot?.remove(); if (taskHeadTarget) taskHeadTarget.style.display = taskHeadLegacyDisplay; statusRoot?.remove(); hiddenStatusChildren.forEach((display, child) => { child.style.display = display; }); hiddenStatusChildren.clear(); };
  }, [host, L]);
  useLayoutEffect(() => {
    const body = host.querySelector<HTMLElement>('#taskpane > .tp-body');
    body?.classList.toggle('react-pane-adapter', current === 'getting-started' || current === 'help' || current === 'search' || current === 'clipart' || current === 'new' || current === 'clipboard' || current === 'styles' || current === 'reveal' || current === 'protect' || current === 'research' || current === 'mailmerge');
    return () => body?.classList.remove('react-pane-adapter');
  }, [host, current]);
  if (!L) return null;
  const dropHint = dropHintActive ? createPortal(<div className="drop-hint" data-react-owned="wordx-drop-hint">{tr(L, 'Drop a document to open it, or pictures to insert them')}</div>, host) : null;
  const docMap = <>{mapMount ? createPortal(<DocumentMap key="wordx-document-map" L={L} />, mapMount) : null}{mailMergeRecipientsMount ? createPortal(<MailMergeRecipientsDialogBody key="wordx-mail-merge-recipients-dialog" L={L} />, mailMergeRecipientsMount) : null}{newAddressListMount ? createPortal(<NewAddressListDialogBody key="wordx-new-address-list-dialog" L={L} />, newAddressListMount) : null}{mergeFieldMount ? createPortal(<InsertMergeFieldDialogBody key="wordx-mail-merge-field-dialog" L={L} />, mergeFieldMount) : null}{reviewersMount ? createPortal(<ReviewersDialogBody key="wordx-reviewers-dialog" L={L} />, reviewersMount) : null}{commentEditMount ? createPortal(<CommentEditDialogBody key="wordx-comment-edit-dialog" L={L} />, commentEditMount) : null}{contentControlMount ? createPortal(<ContentControlDialogBody key="wordx-content-control-dialog" L={L} />, contentControlMount) : null}{spellCheckMount ? createPortal(<SpellCheckDialogBody key="wordx-spell-check-dialog" L={L} />, spellCheckMount) : null}</>;
  const review = reviewMount ? createPortal(<ReviewingPane key="wordx-reviewing-pane" L={L} />, reviewMount) : null;
  const reading = readingMount ? createPortal(<ReadingToolbar key="wordx-reading-toolbar" L={L} />, readingMount) : null;
  const preview = readingMount ? createPortal(<PrintPreviewToolbar key="wordx-print-preview-toolbar" L={L} />, readingMount) : null;
  const taskHead = taskHeadMount ? createPortal(<TaskPaneHeader key="wordx-task-pane-header" L={L} />, taskHeadMount) : null;
  const status = statusMount ? createPortal(<WordxStatusBar key="wordx-statusbar" L={L} />, statusMount) : null;
  const goto = gotoMount ? createPortal(<GoToPane key="wordx-goto-pane" L={L} />, gotoMount) : null;
  const findInput = findInputMount ? createPortal(<FindReplaceInput key="wordx-find-input" L={L} replace={false} />, findInputMount) : null;
  const replaceInput = replaceInputMount ? createPortal(<FindReplaceInput key="wordx-replace-input" L={L} replace />, replaceInputMount) : null;
  const findAux = findAuxMount ? createPortal(<FindAuxiliaryControls key="wordx-find-aux" L={L} />, findAuxMount) : null;
  const findOptions = findOptionsMount ? createPortal(<FindOptionsControls key="wordx-find-options" L={L} />, findOptionsMount) : null;
  const findFormat = findFormatMount ? createPortal(<FindFormatControls key="wordx-find-format" L={L} />, findFormatMount) : null;
  const findFormatSummary = findFormatSummaryMount ? createPortal(<FindFormatSummary key="wordx-find-format-summary" L={L} replace={false} />, findFormatSummaryMount) : null;
  const replaceFormatSummary = replaceFormatSummaryMount ? createPortal(<FindFormatSummary key="wordx-replace-format-summary" L={L} replace />, replaceFormatSummaryMount) : null;
  const findMore = findMoreMount ? createPortal(<FindMoreToggle key="wordx-find-more" L={L} />, findMoreMount) : null;
  const findActions = findActionsMount ? createPortal(<FindFooterActions key="wordx-find-actions" L={L} />, findActionsMount) : null;
  const sortDialog = sortMount ? createPortal(<SortDialogBody key="wordx-sort-dialog" L={L} />, sortMount) : null;
  const tableAutoFormatDialog = tableAutoFormatMount ? createPortal(<TableAutoFormatDialogBody key="wordx-table-autoformat-dialog" L={L} />, tableAutoFormatMount) : null;
  const tabsDialog = tabsDialogMount ? createPortal(<TabsDialogBody key="wordx-tabs-dialog" L={L} />, tabsDialogMount) : null;
  const paragraphDialog = paragraphDialogMount ? createPortal(<ParagraphDialogBody key="wordx-paragraph-dialog" L={L} />, paragraphDialogMount) : null;
  const bordersDialog = bordersDialogMount ? createPortal(<BordersDialogBody key="wordx-borders-dialog" L={L} />, bordersDialogMount) : null;
  const tablePropsDialog = tablePropsDialogMount ? createPortal(<TablePropertiesDialogBody key="wordx-table-properties-dialog" L={L} />, tablePropsDialogMount) : null;
  const rulerTabSelector = rulerCornerMount && L.rulers?.reactTabState ? createPortal(<RulerTabSelector key="wordx-ruler-tab-selector" L={L} />, rulerCornerMount) : null;
  const rulerCanvas = rulerCanvasMount ? createPortal(<canvas className="r-canvas" />, rulerCanvasMount) : null;
  const verticalRulerCanvas = verticalRulerCanvasMount ? createPortal(<canvas className="r-canvas" />, verticalRulerCanvasMount) : null;
  const rulerEvents = rulerEventsMount && L.rulers?.reactRulerDown ? createPortal(<RulerEventSurface key="wordx-ruler-events" L={L} />, rulerEventsMount) : null;
  const verticalRulerEvents = verticalRulerEventsMount && L.rulers?.reactVerticalDoubleClick ? createPortal(<VerticalRulerEventSurface key="wordx-vertical-ruler-events" L={L} />, verticalRulerEventsMount) : null;
  const rulerMarkers = rulerMarkerMount && L.rulers?.reactMarkers ? createPortal(<RulerMarkerLayer key="wordx-ruler-markers" L={L} />, rulerMarkerMount) : null;
  const verticalRulerMarkers = verticalRulerMarkerMount && L.rulers?.reactVerticalMarkers ? createPortal(<VerticalRulerMarkerLayer key="wordx-vertical-ruler-markers" L={L} />, verticalRulerMarkerMount) : null;
  const changeCaseDialog = changeCaseMount ? createPortal(<ChangeCaseDialogBody key="wordx-change-case-dialog" L={L} />, changeCaseMount) : null;
  const textDirectionDialog = textDirectionMount ? createPortal(<TextDirectionDialogBody key="wordx-text-direction-dialog" L={L} />, textDirectionMount) : null;
  const themeDialog = themeMount ? createPortal(<ThemeDialogBody key="wordx-theme-dialog" L={L} />, themeMount) : null;
  const languageDialog = languageMount ? createPortal(<LanguageDialogBody key="wordx-language-dialog" L={L} />, languageMount) : null;
  const dropCapDialog = dropCapMount ? createPortal(<DropCapDialogBody key="wordx-drop-cap-dialog" L={L} />, dropCapMount) : null;
  const lineNumbersDialog = lineNumbersMount ? createPortal(<LineNumbersDialogBody key="wordx-line-numbers-dialog" L={L} />, lineNumbersMount) : null;
  const columnsDialog = columnsMount ? createPortal(<ColumnsDialogBody key="wordx-columns-dialog" L={L} />, columnsMount) : null;
  const pageSetupDialog = pageSetupMount ? createPortal(<PageSetupDialogBody key="wordx-page-setup-dialog" L={L} />, pageSetupMount) : null;
  const captionDialog = captionMount ? createPortal(<CaptionDialogBody key="wordx-caption-dialog" L={L} />, captionMount) : null;
  const captionNumberingDialog = captionNumberingMount ? createPortal(<CaptionNumberingDialogBody key="wordx-caption-numbering-dialog" L={L} />, captionNumberingMount) : null;
  const watermarkDialog = watermarkMount ? createPortal(<WatermarkDialogBody key="wordx-watermark-dialog" L={L} />, watermarkMount) : null;
  const hyperlinkDialog = hyperlinkMount ? createPortal(<HyperlinkDialogBody key="wordx-hyperlink-dialog" L={L} />, hyperlinkMount) : null;
  const symbolDialog = symbolMount ? createPortal(<SymbolDialogBody key="wordx-symbol-dialog" L={L} />, symbolMount) : null;
  const bulletsDialog = <>{bulletsMount ? createPortal(<BulletsDialogBody key="wordx-bullets-dialog" L={L} />, bulletsMount) : null}{customizeListMount ? createPortal(<CustomizeListDialogBody key="wordx-customize-list-dialog" L={L} />, customizeListMount) : null}</>;
  const wordartDialog = <>{shapeMenuMount ? createPortal(<AutoShapeMenu key="wordx-shape-menu" L={L} />, shapeMenuMount) : null}{shadowMenuMount ? createPortal(<ObjectShadowMenu key="wordx-shadow-menu" L={L} />, shadowMenuMount) : null}{wordartMount ? createPortal(<WordartDialogBody key="wordx-wordart-dialog" L={L} />, wordartMount) : null}{lettersArtTextMount ? createPortal(<LettersArtTextDialogBody key="wordx-lettersart-text-dialog" L={L} />, lettersArtTextMount) : null}</>;
  const wordCountDialog = wordCountMount ? createPortal(<WordCountDialogBody key="wordx-word-count-dialog" L={L} />, wordCountMount) : null;
  const summarizeDialog = summarizeMount ? createPortal(<AutoSummarizeDialogBody key="wordx-auto-summarize-dialog" L={L} />, summarizeMount) : null;
  const zoomDialog = zoomMount ? createPortal(<ZoomDialogBody key="wordx-zoom-dialog" L={L} />, zoomMount) : null;
  const hyphenationDialog = hyphenationMount ? createPortal(<HyphenationDialogBody key="wordx-hyphenation-dialog" L={L} />, hyphenationMount) : null;
  const pasteSpecialDialog = pasteSpecialMount ? createPortal(<PasteSpecialDialogBody key="wordx-paste-special-dialog" L={L} />, pasteSpecialMount) : null;
  const macrosDialog = macrosMount ? createPortal(<MacrosDialogBody key="wordx-macros-dialog" L={L} />, macrosMount) : null;
  const optionsViewDialog = optionsViewMount ? createPortal(<OptionsViewDialogBody key="wordx-options-view-dialog" L={L} />, optionsViewMount) : null;
  const optionsGeneralDialog = optionsGeneralMount ? createPortal(<OptionsGeneralDialogBody key="wordx-options-general-dialog" L={L} />, optionsGeneralMount) : null;
  const optionsEditDialog = optionsEditMount ? createPortal(<OptionsEditDialogBody key="wordx-options-edit-dialog" L={L} />, optionsEditMount) : null;
  const optionsPrintDialog = optionsPrintMount ? createPortal(<OptionsPrintDialogBody key="wordx-options-print-dialog" L={L} />, optionsPrintMount) : null;
  const optionsSaveDialog = optionsSaveMount ? createPortal(<OptionsSaveDialogBody key="wordx-options-save-dialog" L={L} />, optionsSaveMount) : null;
  const optionsTrackDialog = optionsTrackMount ? createPortal(<OptionsTrackDialogBody key="wordx-options-track-dialog" L={L} />, optionsTrackMount) : null;
  const optionsUserDialog = optionsUserMount ? createPortal(<OptionsUserDialogBody key="wordx-options-user-dialog" L={L} />, optionsUserMount) : null;
  const optionsCompatDialog = optionsCompatMount ? createPortal(<OptionsCompatDialogBody key="wordx-options-compat-dialog" L={L} />, optionsCompatMount) : null;
  const optionsSecurityDialog = optionsSecurityMount ? createPortal(<OptionsSecurityDialogBody key="wordx-options-security-dialog" L={L} />, optionsSecurityMount) : null;
  const confirmPasswordDialog = confirmPasswordMount ? createPortal(<ConfirmPasswordDialogBody key="wordx-confirm-password-dialog" L={L} />, confirmPasswordMount) : null;
  const optionsSpellingDialog = optionsSpellingMount ? createPortal(<OptionsSpellingDialogBody key="wordx-options-spelling-dialog" L={L} />, optionsSpellingMount) : null;
  const customDictionaryDialog = customDictionaryMount ? createPortal(<CustomDictionaryDialogBody key="wordx-custom-dictionary-dialog" L={L} />, customDictionaryMount) : null;
  const formatObjectColorsDialog = formatObjectColorsMount ? createPortal(<FormatObjectColorsDialogBody key="wordx-format-object-colors-dialog" L={L} />, formatObjectColorsMount) : null;
  const fillEffectsDialog = fillEffectsMount ? createPortal(<FillEffectsDialogBody key="wordx-fill-effects-dialog" L={L} />, fillEffectsMount) : null;
  const formatObjectWebDialog = formatObjectWebMount ? createPortal(<FormatObjectWebDialogBody key="wordx-format-object-web-dialog" L={L} />, formatObjectWebMount) : null;
  const formatObjectSizeDialog = formatObjectSizeMount ? createPortal(<FormatObjectSizeDialogBody key="wordx-format-object-size-dialog" L={L} />, formatObjectSizeMount) : null;
  const formatObjectLayoutDialog = formatObjectLayoutMount ? createPortal(<FormatObjectLayoutDialogBody key="wordx-format-object-layout-dialog" L={L} />, formatObjectLayoutMount) : null;
  const formatObjectAdvancedDialog = formatObjectAdvancedMount ? createPortal(<FormatObjectAdvancedDialogBody key="wordx-format-object-advanced-dialog" L={L} />, formatObjectAdvancedMount) : null;
  const formatObjectTextBoxDialog = formatObjectTextBoxMount ? createPortal(<FormatObjectTextBoxDialogBody key="wordx-format-object-textbox-dialog" L={L} />, formatObjectTextBoxMount) : null;
  const formatObjectPictureDialog = formatObjectPictureMount ? createPortal(<FormatObjectPictureDialogBody key="wordx-format-object-picture-dialog" L={L} />, formatObjectPictureMount) : null;
  const autoCorrectExceptionsDialog = <>{fontMount ? createPortal(<FontDialogBody key="wordx-font-dialog" L={L} />, fontMount) : null}{autoCorrectExceptionsMount ? createPortal(<AutoCorrectExceptionsDialogBody key="wordx-autocorrect-exceptions-dialog" L={L} />, autoCorrectExceptionsMount) : null}{autoCorrectMainMount ? createPortal(<AutoCorrectMainDialogBody key="wordx-autocorrect-main-dialog" L={L} />, autoCorrectMainMount) : null}{autoTextMount ? createPortal(<AutoTextDialogBody key="wordx-autotext-dialog" L={L} />, autoTextMount) : null}{autoCorrectFormatMount ? createPortal(<AutoCorrectFormatDialogBody key="wordx-autocorrect-format-dialog" L={L} />, autoCorrectFormatMount) : null}</>;
  const customizeDialog = customizeMount ? createPortal(<CustomizeDialogBody key="wordx-customize-dialog" L={L} />, customizeMount) : null;
  const drawingGridDialog = drawingGridMount ? createPortal(<DrawingGridDialogBody key="wordx-drawing-grid-dialog" L={L} />, drawingGridMount) : null;
  const autoFormatDialog = autoFormatMount ? createPortal(<AutoFormatDialogBody key="wordx-autoformat-dialog" L={L} />, autoFormatMount) : null;
  const autoFormatReviewDialog = autoFormatReviewMount ? createPortal(<AutoFormatReviewDialogBody key="wordx-autoformat-review-dialog" L={L} />, autoFormatReviewMount) : null;
  const envelopesDialog = envelopesMount ? createPortal(<EnvelopesDialogBody key="wordx-envelopes-dialog" L={L} />, envelopesMount) : null;
  const indexTablesDialog = indexTablesMount ? createPortal(<IndexTablesDialogBody key="wordx-index-tables-dialog" L={L} />, indexTablesMount) : null;
  const letterWizardDialog = letterWizardMount ? createPortal(<LetterWizardDialogBody key="wordx-letter-wizard-dialog" L={L} />, letterWizardMount) : null;
  const organizerDialog = organizerMount ? createPortal(<OrganizerDialogBody key="wordx-organizer-dialog" L={L} />, organizerMount) : null;
  const templatesDialog = templatesMount ? createPortal(<TemplatesDialogBody key="wordx-templates-dialog" L={L} />, templatesMount) : null;
  const propertiesDialog = propertiesMount ? createPortal(<PropertiesDialogBody key="wordx-properties-dialog" L={L} />, propertiesMount) : null;
  const saveAsDialog = saveAsMount ? createPortal(<SaveAsDialogBody key="wordx-save-as-dialog" L={L} />, saveAsMount) : null;
  const templatePickerDialog = templatePickerMount ? createPortal(<TemplatePickerDialogBody key="wordx-template-picker-dialog" L={L} />, templatePickerMount) : null;
  const printDialog = printMount ? createPortal(<PrintDialogBody key="wordx-print-dialog" L={L} />, printMount) : null;
  const bookmark = bookmarkMount ? createPortal(<BookmarkDialogBody key="wordx-bookmark-dialog" L={L} />, bookmarkMount) : null;
  const object = objectMount ? createPortal(<InsertObjectDialogBody key="wordx-object-dialog" L={L} />, objectMount) : null;
  const equation = equationMount ? createPortal(<EquationDialogBody key="wordx-equation-dialog" L={L} />, equationMount) : null;
  const breakDialog = breakMount ? createPortal(<BreakDialogBody key="wordx-break-dialog" L={L} />, breakMount) : null;
  const dateTimeDialog = dateTimeMount ? createPortal(<DateTimeDialogBody key="wordx-date-time-dialog" L={L} />, dateTimeMount) : null;
  const pageNumbersDialog = pageNumbersMount ? createPortal(<PageNumbersDialogBody key="wordx-page-numbers-dialog" L={L} />, pageNumbersMount) : null;
  const pageNumberFormatDialog = pageNumberFormatMount ? createPortal(<PageNumberFormatDialogBody key="wordx-page-number-format-dialog" L={L} />, pageNumberFormatMount) : null;
  const cellOperationDialog = cellOperationMount ? createPortal(<CellOperationDialogBody key={`wordx-cell-operation-${L.dlg?.reactCellDialogKind}`} L={L} />, cellOperationMount) : null;
  const insertTableDialog = insertTableMount ? createPortal(<InsertTableDialogBody key="wordx-insert-table-dialog" L={L} />, insertTableMount) : null;
  const splitCellsDialog = splitCellsMount ? createPortal(<SplitCellsDialogBody key="wordx-split-cells-dialog" L={L} />, splitCellsMount) : null;
  const textToTableDialog = textToTableMount ? createPortal(<TextToTableDialogBody key="wordx-text-to-table-dialog" L={L} />, textToTableMount) : null;
  const tableToTextDialog = tableToTextMount ? createPortal(<TableToTextDialogBody key="wordx-table-to-text-dialog" L={L} />, tableToTextMount) : null;
  const formulaDialog = formulaMount ? createPortal(<FormulaDialogBody key="wordx-formula-dialog" L={L} />, formulaMount) : null;
  const crossRefDialog = crossRefMount ? createPortal(<CrossReferenceDialogBody key="wordx-cross-reference-dialog" L={L} />, crossRefMount) : null;
  const markEntryDialog = markEntryMount ? createPortal(<MarkEntryDialogBody key="wordx-mark-entry-dialog" L={L} />, markEntryMount) : null;
  const fieldDialog = fieldMount ? createPortal(<FieldDialogBody key="wordx-field-dialog" L={L} />, fieldMount) : null;
  const footnoteDialog = footnoteMount ? createPortal(<FootnoteDialogBody key="wordx-footnote-dialog" L={L} />, footnoteMount) : null;
  const convertNotesDialog = convertNotesMount ? createPortal(<ConvertNotesDialogBody key="wordx-convert-notes-dialog" L={L} />, convertNotesMount) : null;
  const styleDialog = styleDialogMount ? createPortal(<WordxStyleDialogBody key="wordx-style-dialog" L={L} />, styleDialogMount) : null;
  const diagramDialog = diagramMount ? createPortal(<DiagramDialogBody key="wordx-diagram-dialog" L={L} />, diagramMount) : null;
  if (!mount) return <><OutlineInteractionBridge L={L} />{dropHint}{docMap}{review}{reading}{preview}{taskHead}{status}{goto}{findInput}{replaceInput}{findAux}{findOptions}{findFormat}{findFormatSummary}{replaceFormatSummary}{findMore}{findActions}{sortDialog}{tableAutoFormatDialog}{tabsDialog}{paragraphDialog}{bordersDialog}{tablePropsDialog}{rulerTabSelector}{rulerCanvas}{verticalRulerCanvas}{rulerEvents}{verticalRulerEvents}{rulerMarkers}{verticalRulerMarkers}{changeCaseDialog}{textDirectionDialog}{themeDialog}{languageDialog}{dropCapDialog}{lineNumbersDialog}{columnsDialog}{pageSetupDialog}{captionDialog}{captionNumberingDialog}{watermarkDialog}{hyperlinkDialog}{symbolDialog}{bulletsDialog}{wordartDialog}{wordCountDialog}{summarizeDialog}{zoomDialog}{hyphenationDialog}{pasteSpecialDialog}{macrosDialog}{autoCorrectExceptionsDialog}{customizeDialog}{drawingGridDialog}{autoFormatDialog}{envelopesDialog}{indexTablesDialog}{letterWizardDialog}{organizerDialog}{templatesDialog}{propertiesDialog}{saveAsDialog}{templatePickerDialog}{printDialog}{bookmark}{object}{equation}{breakDialog}{dateTimeDialog}{pageNumbersDialog}{pageNumberFormatDialog}{cellOperationDialog}{insertTableDialog}{splitCellsDialog}{textToTableDialog}{tableToTextDialog}{formulaDialog}{crossRefDialog}{markEntryDialog}{fieldDialog}{footnoteDialog}{convertNotesDialog}{styleDialog}{diagramDialog}</>;
  let pane = null;
  if (current === 'getting-started') pane = createPortal(<GettingStartedPane key="wordx-getting-started" L={L} />, mount);
  else if (current === 'styles') pane = createPortal(<StylesPane key="wordx-styles" L={L} />, mount);
  else if (current === 'reveal') pane = createPortal(<RevealFormattingPane key="wordx-reveal" L={L} />, mount);
  else if (current === 'protect') pane = createPortal(<ProtectDocumentPane key="wordx-protect" L={L} />, mount);
  else if (current === 'research') pane = createPortal(<ResearchPane key="wordx-research" L={L} />, mount);
  else if (current === 'mailmerge') pane = createPortal(<MailMergePane key="wordx-mailmerge" L={L} />, mount);
  else if (current === 'search') pane = createPortal(<SearchPane key="wordx-search" L={L} />, mount);
  else if (current === 'clipart') pane = createPortal(<ClipArtPane key="wordx-clipart" L={L} />, mount);
  else if (current === 'new') pane = createPortal(<NewDocumentPane key="wordx-new" L={L} />, mount);
  else if (current === 'clipboard') pane = createPortal(<ClipboardPane key="wordx-clipboard" L={L} />, mount);
  else if (current === 'help') pane = createPortal(<HelpPane key="wordx-help" L={L} />, mount);
  return <><OutlineInteractionBridge L={L} />{dropHint}{pane}{docMap}{review}{reading}{preview}{taskHead}{status}{goto}{findInput}{replaceInput}{findAux}{findOptions}{findFormat}{findFormatSummary}{replaceFormatSummary}{findMore}{findActions}{sortDialog}{tableAutoFormatDialog}{tabsDialog}{paragraphDialog}{bordersDialog}{tablePropsDialog}{rulerTabSelector}{rulerCanvas}{verticalRulerCanvas}{rulerEvents}{verticalRulerEvents}{rulerMarkers}{verticalRulerMarkers}{changeCaseDialog}{textDirectionDialog}{themeDialog}{languageDialog}{dropCapDialog}{lineNumbersDialog}{columnsDialog}{pageSetupDialog}{captionDialog}{captionNumberingDialog}{watermarkDialog}{hyperlinkDialog}{symbolDialog}{bulletsDialog}{wordartDialog}{wordCountDialog}{summarizeDialog}{zoomDialog}{hyphenationDialog}{pasteSpecialDialog}{macrosDialog}{optionsViewDialog}{optionsGeneralDialog}{optionsEditDialog}{optionsPrintDialog}{optionsSaveDialog}{optionsTrackDialog}{optionsUserDialog}{optionsCompatDialog}{optionsSecurityDialog}{confirmPasswordDialog}{optionsSpellingDialog}{customDictionaryDialog}{formatObjectColorsDialog}{fillEffectsDialog}{formatObjectWebDialog}{formatObjectSizeDialog}{formatObjectLayoutDialog}{formatObjectAdvancedDialog}{formatObjectTextBoxDialog}{formatObjectPictureDialog}{autoCorrectExceptionsDialog}{customizeDialog}{drawingGridDialog}{autoFormatDialog}{autoFormatReviewDialog}{envelopesDialog}{indexTablesDialog}{letterWizardDialog}{organizerDialog}{templatesDialog}{propertiesDialog}{saveAsDialog}{templatePickerDialog}{printDialog}{bookmark}{object}{equation}{breakDialog}{dateTimeDialog}{pageNumbersDialog}{pageNumberFormatDialog}{cellOperationDialog}{insertTableDialog}{splitCellsDialog}{textToTableDialog}{tableToTextDialog}{formulaDialog}{crossRefDialog}{markEntryDialog}{fieldDialog}{footnoteDialog}{convertNotesDialog}{styleDialog}{diagramDialog}</>;
}
