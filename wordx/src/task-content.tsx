import { useEffect, useLayoutEffect, useRef, useState, type CSSProperties, type KeyboardEvent } from 'react';
import { createPortal } from 'react-dom';

type HelpEntry = [string, string];
interface ClipArtItem { id: string; name: string; svg: string }
interface ClipboardItem { index: number; text: string; table: boolean }
interface WordxRuntime {
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
    preview?: { mag?: boolean };
    reactTaskShellState?(): { current: string; title: string; backDisabled: boolean; forwardDisabled: boolean; items: Array<{ id: string; title: string; current: boolean }> };
    reactTaskShellAction?(action: string, anchor?: HTMLButtonElement): void;
  };
  layout?: { zoom?: number };
  find?: { reactGoToCatalog?(): Array<{ id: string; label: string; prompt: string }>; reactGoToAction?(kind: string, value: string, dir: number): boolean; reactCloseDialog?(): void };
  dlg?: { reactDropCapCatalog?(): { position: string; positions: Array<[string, string]>; fonts: string[]; font: string; lines: number; distanceText: string }; reactDropCapAction?(action: string, value?: string): void; reactLanguageCatalog?(): { languages: Array<{ value: string; label: string }>; language: string; noProof: boolean; detect: boolean }; reactLanguageAction?(action: string, value?: string): void; reactBookmarkCatalog?(): { name: string; sort: string; hidden: boolean; names: string[] }; reactBookmarkAction?(action: string, value?: string): boolean | void; reactObjectCatalog?(): Array<{ value: string; label: string }>; reactObjectAction?(action: string, value?: string): void; reactEquationAction?(action: string, value?: string): void; reactBreakCatalog?(): Array<{ value: string; label: string }>; reactBreakAction?(value: string): void; reactDateTimeCatalog?(): Array<{ value: string; label: string }>; reactDateTimeState?(): { format: string; autoUpdate: boolean }; reactPageNumberAction?(position: string, alignment: string, showFirst: boolean): void; reactPageNumberFormatState?(): { fmt: string; chapter: boolean; mode: string; start: number }; reactPageNumberFormatCatalog?(): Array<{ value: string; label: string }>; reactPageNumberFormatAction?(format: string, chapter: boolean, mode: string, start: number): void; reactCellDialogCatalog?(kind?: string): Array<[string, string]>; reactCellDialogKind?: string; reactInsertTableState?(): { columns: number; rows: number }; reactInsertTableAction?(action: string, options?: { columns: number; rows: number; fit: string; width: string; remember: boolean }): void; reactSplitCellsState?(): { columns: number; rows: number; merge: boolean }; reactTextToTableState?(): { initialSeparator: string; columns: number; paragraphs: Array<{ text: string; tabs: number }> }; reactTextToTableAction?(options: { sep: string; cols: number; autofit: string }): void; reactFormulaCatalog?(): { guess: string; functions: string[]; bookmarks: string[]; formats: string[] }; reactFormulaAction?(formula: string, format: string): void; reactCrossRefCatalog?(type?: string): { types: Array<{ value: string; label: string }>; references: Array<{ value: string; label: string }>; items: Array<{ value: string; label: string }> }; reactCrossRefAction?(type: string, index: number, reference: string, hyperlink: boolean, aboveBelow: boolean): void; reactMarkEntryState?(): { main: string; sub: string; option: string; xref: string; bold: boolean; italic: boolean }; reactMarkEntryAction?(all: boolean, values: { main: string; sub: string; option: string; xref: string; bold: boolean; italic: boolean }): void };
  app?: { fileName?: string; view?: string; prevView?: string; toggleDocMap?(): void; toggleThumbs?(): void; setView?(view: string): void; reactStatusState?(): Array<{ id: string; text: string; className: string; tip: string; html: string }>; reactStatusAction?(action: string): void };
  review?: { togglePane(on: boolean): void; reactPaneCatalog?(): Array<{ index: number; kind: string; head: string; body: string; color: string }>; reactPaneSelect?(index: number, edit?: boolean): void };
  icons?: { get(name: string): string; app?(size: number): string };
  ui?: { keyText(key: string): string };
  t?(text: string, ...args: unknown[]): string;
  bus?: { on(event: string, listener: () => void): () => void };
}
function runtime(): WordxRuntime | undefined { return (window as Window & { L?: WordxRuntime }).L; }
function tr(L: WordxRuntime, text: string): string { return L.t?.(text) || text; }

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
  useEffect(() => L.bus?.on('wordx-task-render', () => refresh((value) => value + 1)), [L]);
  const state = L.panes?.reactTaskShellState?.();
  if (!state) return null;
  const button = (action: string, icon: string, label: string, disabled = false) => <button className="tp-nav" type="button" aria-label={tr(L, label)} data-tip={tr(L, label)} disabled={disabled} onClick={(event) => L.panes?.reactTaskShellAction?.(action, event.currentTarget)}><span dangerouslySetInnerHTML={{ __html: L.icons?.get(icon) || '' }} /></button>;
  return <div className="tp-head">
    {button('back', 'back', 'Back', state.backDisabled)}{button('forward', 'forward', 'Forward', state.forwardDisabled)}{button('home', 'home', 'Home')}
    <button className="tp-title" type="button" aria-haspopup="menu" aria-label={tr(L, state.title)} onClick={(event) => L.panes?.reactTaskShellAction?.('menu', event.currentTarget)}><span>{tr(L, state.title)}</span><span className="dd-arrow" /></button>
    {button('close', 'close', 'Close task pane')}
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
    <input ref={input} id={id} type="text" list={replace ? 'fr-rhist' : 'fr-fhist'} value={value} style={{ flex: 1, minWidth: 0 }} onFocus={() => adapter?.reactFindAction?.('focus', replace ? 'repl' : 'find')} onChange={(event) => { setValue(event.currentTarget.value); adapter?.reactFindAction?.('input'); }} />
  </div>;
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
  const [findMoreMount, setFindMoreMount] = useState<HTMLElement | null>(null);
  const [sortMount, setSortMount] = useState<HTMLElement | null>(null);
  const [tableAutoFormatMount, setTableAutoFormatMount] = useState<HTMLElement | null>(null);
  const [tabsDialogMount, setTabsDialogMount] = useState<HTMLElement | null>(null);
  const [changeCaseMount, setChangeCaseMount] = useState<HTMLElement | null>(null);
  const [textDirectionMount, setTextDirectionMount] = useState<HTMLElement | null>(null);
  const [themeMount, setThemeMount] = useState<HTMLElement | null>(null);
  const [languageMount, setLanguageMount] = useState<HTMLElement | null>(null);
  const [dropCapMount, setDropCapMount] = useState<HTMLElement | null>(null);
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
  const [diagramMount, setDiagramMount] = useState<HTMLElement | null>(null);
  const [, refresh] = useState(0);
  const L = runtime();
  const current = L?.panes?.task?.current;
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
    let findMoreTarget: HTMLElement | null = null;
    let sortTarget: HTMLElement | null = null;
    let tableAutoFormatTarget: HTMLElement | null = null;
    let tabsDialogTarget: HTMLElement | null = null;
    let changeCaseTarget: HTMLElement | null = null;
    let textDirectionTarget: HTMLElement | null = null;
    let themeTarget: HTMLElement | null = null;
    let languageTarget: HTMLElement | null = null;
    let dropCapTarget: HTMLElement | null = null;
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
      const nextFindMore = host.querySelector<HTMLElement>('#react-wordx-find-more-root');
      if (nextFindMore !== findMoreTarget) { findMoreTarget = nextFindMore; setFindMoreMount(findMoreTarget); }
      const nextSort = host.querySelector<HTMLElement>('#react-wordx-sort-root');
      if (nextSort !== sortTarget) { sortTarget = nextSort; setSortMount(sortTarget); }
      const nextTableAutoFormat = host.querySelector<HTMLElement>('#react-wordx-table-autoformat-root');
      if (nextTableAutoFormat !== tableAutoFormatTarget) { tableAutoFormatTarget = nextTableAutoFormat; setTableAutoFormatMount(tableAutoFormatTarget); }
      const nextTabsDialog = host.querySelector<HTMLElement>('#react-wordx-tabs-root');
      if (nextTabsDialog !== tabsDialogTarget) { tabsDialogTarget = nextTabsDialog; setTabsDialogMount(tabsDialogTarget); }
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
  const docMap = mapMount ? createPortal(<DocumentMap key="wordx-document-map" L={L} />, mapMount) : null;
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
  const findMore = findMoreMount ? createPortal(<FindMoreToggle key="wordx-find-more" L={L} />, findMoreMount) : null;
  const sortDialog = sortMount ? createPortal(<SortDialogBody key="wordx-sort-dialog" L={L} />, sortMount) : null;
  const tableAutoFormatDialog = tableAutoFormatMount ? createPortal(<TableAutoFormatDialogBody key="wordx-table-autoformat-dialog" L={L} />, tableAutoFormatMount) : null;
  const tabsDialog = tabsDialogMount ? createPortal(<TabsDialogBody key="wordx-tabs-dialog" L={L} />, tabsDialogMount) : null;
  const changeCaseDialog = changeCaseMount ? createPortal(<ChangeCaseDialogBody key="wordx-change-case-dialog" L={L} />, changeCaseMount) : null;
  const textDirectionDialog = textDirectionMount ? createPortal(<TextDirectionDialogBody key="wordx-text-direction-dialog" L={L} />, textDirectionMount) : null;
  const themeDialog = themeMount ? createPortal(<ThemeDialogBody key="wordx-theme-dialog" L={L} />, themeMount) : null;
  const languageDialog = languageMount ? createPortal(<LanguageDialogBody key="wordx-language-dialog" L={L} />, languageMount) : null;
  const dropCapDialog = dropCapMount ? createPortal(<DropCapDialogBody key="wordx-drop-cap-dialog" L={L} />, dropCapMount) : null;
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
  const diagramDialog = diagramMount ? createPortal(<DiagramDialogBody key="wordx-diagram-dialog" L={L} />, diagramMount) : null;
  if (!mount) return <>{docMap}{review}{reading}{preview}{taskHead}{status}{goto}{findInput}{replaceInput}{findAux}{findOptions}{findFormat}{findMore}{sortDialog}{tableAutoFormatDialog}{tabsDialog}{changeCaseDialog}{textDirectionDialog}{themeDialog}{languageDialog}{dropCapDialog}{bookmark}{object}{equation}{breakDialog}{dateTimeDialog}{pageNumbersDialog}{pageNumberFormatDialog}{cellOperationDialog}{insertTableDialog}{splitCellsDialog}{textToTableDialog}{tableToTextDialog}{formulaDialog}{crossRefDialog}{markEntryDialog}{fieldDialog}{footnoteDialog}{diagramDialog}</>;
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
  return <>{pane}{docMap}{review}{reading}{preview}{taskHead}{status}{goto}{findInput}{replaceInput}{findAux}{findOptions}{findFormat}{findMore}{sortDialog}{tableAutoFormatDialog}{tabsDialog}{changeCaseDialog}{textDirectionDialog}{themeDialog}{languageDialog}{dropCapDialog}{bookmark}{object}{equation}{breakDialog}{dateTimeDialog}{pageNumbersDialog}{pageNumberFormatDialog}{cellOperationDialog}{insertTableDialog}{splitCellsDialog}{textToTableDialog}{tableToTextDialog}{formulaDialog}{crossRefDialog}{markEntryDialog}{fieldDialog}{footnoteDialog}{diagramDialog}</>;
}
