import { Children, createElement, useEffect, useLayoutEffect, useRef, useState, type MouseEvent, type ReactNode } from 'react';
import { createPortal } from 'react-dom';
import type { PptxRibbonSpec, RibbonContext, RibbonItem } from './ribbon-spec.js';

interface TaskContentRuntime {
  app?: { newPresentation(): void; loadSample(): void; autoContent(): void; insertClipArt(item: ClipArtItem): void; applyColorScheme(scheme: ColorScheme, all: boolean): void; editColorScheme(): void; opts?: { autoPreview?: boolean; layoutPaneOnNew?: boolean }; saveOpts?(): void; applyLayoutToSel?(key: string, reapply?: boolean): void; insertSlide?(key: string): void; clipHistory?: Array<{ kind: string; label: string }>; pasteItem?(item: unknown): void };
  diagram?: { reactTextPaneCatalog?(): DiagramTextCatalog; reactTextFocus?(id: string): void; reactTextChange?(id: string, text: string): void; reactTextKey?(event: KeyboardEvent, id: string, input: HTMLInputElement): void; reactTextSet?(key: 'layout' | 'colors' | 'style', value: string): void; reactTextReset?(): void; reactTextConvert?(): void };
  
  clipart?: { search(query: string): Array<{ id: string; name: string; svg: string }>; url(item: { id: string; name: string; svg: string }): string };
  panes?: { helpQuery?: string; reactDesignCatalog?(): DesignCatalog; reactApplyDesign?(key: string, all: boolean): void; reactApplyUsedDesign?(id: string): void; reactBrowseDesign?(): void; reactDesignMenu?(key: string, x: number, y: number): void; reactColorCatalog?(): { current: ColorScheme; schemes: ColorScheme[] }; reactApplyColorScheme?(scheme: ColorScheme, all: boolean): void; reactColorMenu?(scheme: ColorScheme, x: number, y: number): void; reactEditColorScheme?(): void; reactAnimationSchemes?(): AnimationSchemeChoice[]; reactApplyAnimationScheme?(definition: AnimationDefinition, all: boolean): void; reactPlayAnimationPreview?(): void; reactLayoutCatalog?(): LayoutCatalog; reactApplyLayout?(key: string): void; reactLayoutMenu?(key: string, x: number, y: number): void; reactSetLayoutPaneOnNew?(value: boolean): void; reactTransitionCatalog?(): TransitionCatalog; reactChooseTransition?(choice: TransitionChoice): void; reactSetTransitionSpeed?(speed: string): void; reactSetTransitionClick?(enabled: boolean): void; reactSetTransitionAfter?(seconds: number | null): void; reactTransitionToAll?(): void; reactPreviewTransition?(): void; reactSetAutoPreview?(enabled: boolean): void; reactCustomAnimationCatalog?(): CustomAnimationCatalog; reactAddAnimationMenu?(x: number, y: number): void; reactSelectAnimation?(id: string): void; reactAnimationMenu?(id: string, x: number, y: number): void; reactRemoveAnimation?(id: string): void; reactUpdateAnimation?(key: string, value: string | number): void; reactMoveAnimation?(delta: number): void; reactAnimationOptions?(id: string): void; reactPlayAnimation?(id?: string): void; reactSetAnimationAutoPreview?(enabled: boolean): void; reactClipboardCatalog?(): Array<{ index: number; kind: string; label: string; icon: string }>; reactPasteClipboard?(index: number): void; reactPasteAllClipboard?(): void; reactClearClipboard?(): void; task?: { current?: string; registry?: Record<string, { title: string }>; show(id: string): void; helpData?: { topics: Array<[string, string]>; keys: Array<[string, string]> } } };
  ribbonSpec?(): PptxRibbonSpec;
  ui?: { exec(command: string, arg?: unknown, event?: Event): unknown; cmds?: Record<string, { label?: string; tip?: string; icon?: string; key?: string; enabled?: () => boolean; checked?: () => boolean }>; stripAmp?(value: string): string; openMenu?(items: unknown[], rect: DOMRect | { left: number; bottom: number }): unknown; keyText?(key: string): string };
  icons?: { get(name: string, size?: number): string; app?(size?: number): string };
  t?(message: string, ...args: unknown[]): string;
  bus?: { on(event: string, listener: () => void): () => void };
}

interface ClipArtItem { id: string; name: string; keys?: string; svg: string }
interface DesignChoice { id?: string; key?: string; name: string; preview: string }
interface DesignCatalog { currentId?: string; currentName?: string; used: DesignChoice[]; available: DesignChoice[] }
interface ColorScheme { [key: string]: string }
type AnimationDefinition = { tr?: Record<string, unknown>; title?: Array<string | null>; body?: Array<string | null> } | null;
interface AnimationSchemeChoice { name: string; index: number; group: boolean; definition: AnimationDefinition }
interface LayoutCatalog { currentKey: string | null; groups: Array<{ name: string; layouts: Array<{ key: string; name: string; preview: string }> }>; showWhenInserting: boolean }
interface TransitionChoice { name: string; index: number; type: string; dir?: string; thruBlk?: boolean; spokes?: number; orient?: string }
interface TransitionCatalog { available: boolean; transitions: TransitionChoice[]; currentName: string; speed: string; click: boolean; after: number | null; autoPreview: boolean }
interface AnimationEntry { id: string; sid: string; start: string; dur: number; cls: string; eff: string; dir: string; by: string; step: string; color: string; name: string; label: string; tooltip: string }
interface CustomAnimationCatalog { available: boolean; animations: AnimationEntry[]; selectedId: string | null; hasSelection: boolean; selectedName?: string; selectedStart?: string; selectedDirection?: string; selectedSpeed?: number; speedDisabled?: boolean; starts: Array<{ value: string; name: string }>; directions: Array<{ value: string; name: string }>; speeds: Array<{ name: string; value: number }>; directionDisabled?: boolean; autoPreview: boolean }
interface DiagramTextChoice { id: string; name: string; preview: string; selected: boolean }
interface DiagramTextCatalog { available: boolean; focusId?: string | null; rows: Array<{ id: string; text: string; level: number }>; layouts: DiagramTextChoice[]; colors: DiagramTextChoice[]; styles: Array<{ language: string; title: string; choices: DiagramTextChoice[] }> }

function runtime(): TaskContentRuntime | undefined { return (window as Window & { L?: TaskContentRuntime }).L; }
function translate(L: TaskContentRuntime, text: string, ...args: unknown[]): string { return L.t?.(text, ...args) || text; }

/** React-owned PPTX task-pane bodies backed by the existing document and command engines. */
export function PptxTaskContentLayer({ host }: { host: HTMLElement }) {
  const [mount, setMount] = useState<HTMLElement | null>(null);
  const [revision, refresh] = useState(0);
  const L = runtime();
  const task = L?.panes?.task;
  const current = task?.current;
  const taskTitle = current ? task?.registry?.[current]?.title : undefined;
  const toolContext = L?.ribbonSpec?.().contextual.find((context) => context.set === taskTitle);
  useEffect(() => {
    let slot: HTMLElement | null = null;
    let off: (() => void) | undefined;
    const attach = () => {
      const engine = runtime();
      const target = host.querySelector<HTMLElement>('#react-task-body-root');
      if (!engine || !target) return;
      if (!off && engine.bus) off = engine.bus.on('pptx-task-render', () => refresh((value) => value + 1));
      if (target === slot) return;
      slot = target;
      target.classList.add('vibeoffice-react-owned');
      setMount(target);
    };
    const observer = new MutationObserver(attach);
    observer.observe(host, { childList: true, subtree: true });
    attach();
    return () => { observer.disconnect(); off?.(); slot?.classList.remove('vibeoffice-react-owned'); };
  }, [host]);
  useLayoutEffect(() => {
    const body = host.querySelector<HTMLElement>('#taskpane > .tp-body');
    body?.classList.toggle('react-pane-adapter', !!toolContext || current === 'getting-started' || current === 'new' || current === 'help' || current === 'clipart' || current === 'design' || current === 'colors' || current === 'animSchemes' || current === 'layout' || current === 'transition' || current === 'customAnim' || current === 'clipboard' || current === 'diagramText');
    return () => body?.classList.remove('react-pane-adapter');
  }, [host, current, toolContext?.id]);
  if (!mount || !L || !task || (!toolContext && !['getting-started', 'new', 'help', 'clipart', 'design', 'colors', 'animSchemes', 'layout', 'transition', 'customAnim', 'clipboard', 'diagramText'].includes(current || ''))) return null;
  const goToSearch = (query: string) => { if (L.panes) L.panes.helpQuery = query; task.show('help'); };
  const link = (icon: string, label: string, run: () => void) => <button className="tp-link" type="button" onClick={run}>
    {icon && <span dangerouslySetInnerHTML={{ __html: L.icons?.get(icon) || '' }} />}<span>{translate(L, label)}</span>
  </button>;
  const section = (title: string, ...children: ReactNode[]) => createElement('section', { className: 'tp-sec' }, createElement('div', { className: 'tp-h' }, translate(L, title)), ...Children.toArray(children));
  const appIcon = L.icons?.app?.(32) || '';
  const editionName = (window as Window & { VO?: { editionName?(app: string): string } }).VO?.editionName?.('PPTX') || 'PPTX';
  if (toolContext) return <ContextualToolsPane L={L} mount={mount} context={toolContext} />;
  if (current === 'diagramText') return <IntelliArtTextPane L={L} mount={mount} revision={revision} />;
  if (current === 'help') return <HelpPane L={L} task={task} mount={mount} />;
  if (current === 'clipart') return <ClipArtPane L={L} mount={mount} />;
  if (current === 'design') return <DesignPane L={L} mount={mount} />;
  if (current === 'colors') return <ColorSchemesPane L={L} mount={mount} />;
  if (current === 'animSchemes') return <AnimationSchemesPane L={L} mount={mount} />;
  if (current === 'layout') return <SlideLayoutPane L={L} mount={mount} />;
  if (current === 'transition') return <SlideTransitionPane L={L} mount={mount} />;
  if (current === 'customAnim') return <CustomAnimationPane L={L} mount={mount} />;
  if (current === 'clipboard') return <ClipboardPane L={L} mount={mount} />;
  if (current === 'new') return createPortal(<div className="tp-body">
    {section('New', link('new', 'Blank presentation', () => L.app?.newPresentation()), link('design', 'From design template', () => { L.app?.newPresentation(); task.show('design'); }), link('properties', 'From AutoContent wizard...', () => L.app?.autoContent()), link('open', 'From existing presentation...', () => L.ui?.exec('open')), link('photoAlbum', 'Photo album...', () => L.ui?.exec('photoAlbum')))}
    {section('Templates', link('slideshow', 'Sample: café expansion plan', () => L.app?.loadSample()))}
  </div>, mount);
  return createPortal(<div className="tp-body">
    <div className="tp-row"><span dangerouslySetInnerHTML={{ __html: appIcon }} /><b translate="no" style={{ color: 'var(--pane-ink)' }}>{editionName}</b></div>
    {section('Search for:', <div className="tp-row">
      <input id="gs-search" type="text" placeholder={translate(L, 'Search help')} style={{ flex: 1, minWidth: 0 }} onKeyDown={(event) => { if (event.key === 'Enter') goToSearch(event.currentTarget.value); }} />
      <button className="btn small" type="button" onClick={() => goToSearch(host.querySelector<HTMLInputElement>('#gs-search')?.value || '')}>{translate(L, 'Go')}</button>
    </div>)}
    {section('Open', link('open', 'Open a presentation (.pptx)...', () => L.ui?.exec('open')), link('new', 'Create a new presentation...', () => task.show('new')), link('slideshow', 'View the sample presentation', () => L.app?.loadSample()))}
    {section('Learn', link('help', 'Keyboard shortcuts and help', () => task.show('help')), link('design', 'Apply a design template', () => task.show('design')), link('animation', 'Animate your slides', () => task.show('customAnim')))}
  </div>, mount);
}

function ClipArtPane({ L, mount }: { L: TaskContentRuntime; mount: HTMLElement }) {
  const [draft, setDraft] = useState('');
  const [query, setQuery] = useState('');
  const clips = L.clipart?.search(query) || [];
  const insert = (clip: ClipArtItem) => L.app?.insertClipArt(clip);
  return createPortal(<div className="tp-body">
    <section className="tp-sec"><div className="tp-h">{translate(L, 'Search for:')}</div><div className="tp-row">
      <input id="ca-q" type="text" placeholder={translate(L, 'e.g. people, growth, time')} value={draft} style={{ flex: 1, minWidth: 0 }} onChange={(event) => setDraft(event.currentTarget.value)} onKeyDown={(event) => { if (event.key === 'Enter') setQuery(draft); }} />
      <button className="btn small" type="button" onClick={() => setQuery(draft)}>{translate(L, 'Go')}</button>
    </div></section>
    <section className="tp-sec"><div className="tp-h">{translate(L, 'Results')}</div><div className="clip-grid">
      {clips.map((clip) => <button className="clip-it" type="button" data-tip={clip.name} key={clip.id} onClick={() => insert(clip)}><img src={L.clipart?.url(clip)} alt={translate(L, clip.name)} /><span>{translate(L, clip.name)}</span></button>)}
      {!clips.length && <div className="tp-note">{translate(L, 'No clips found.')}</div>}
    </div></section>
    <div className="tp-note">{translate(L, 'Click a clip to insert it on the current slide.')}</div>
  </div>, mount);
}

function DesignPane({ L, mount }: { L: TaskContentRuntime; mount: HTMLElement }) {
  const catalog = L.panes?.reactDesignCatalog?.() || { used: [], available: [] };
  const showMenu = (key: string, event: MouseEvent) => { event.preventDefault(); L.panes?.reactDesignMenu?.(key, event.clientX, event.clientY); };
  const card = (design: DesignChoice, used: boolean) => <button className={`dsn-it${(used ? design.id === catalog.currentId : design.name === catalog.currentName) ? ' on' : ''}`} type="button" data-tip={design.name} data-design-key={design.key || design.id} aria-label={design.name} key={`${used ? 'used' : 'available'}-${design.id || design.key}`} onClick={() => used ? design.id && L.panes?.reactApplyUsedDesign?.(design.id) : design.key && L.panes?.reactApplyDesign?.(design.key, true)} onContextMenu={(event) => !used && design.key && showMenu(design.key, event)}>
    <div className="dsn-preview" aria-hidden="true" dangerouslySetInnerHTML={{ __html: design.preview }} /><span className="dsn-name">{design.name}</span>
    {!used && <span className="dsn-dd" role="button" tabIndex={0} aria-label={translate(L, 'Apply design options')} onClick={(event) => { event.stopPropagation(); design.key && showMenu(design.key, event); }} onKeyDown={(event) => { if (event.key === 'Enter' || event.key === ' ') { event.preventDefault(); event.stopPropagation(); const rect=event.currentTarget.getBoundingClientRect(); design.key && L.panes?.reactDesignMenu?.(design.key, rect.left, rect.bottom); } }}><span className="dd-arrow" style={{ margin: 0 }} /></span>}
  </button>;
  return createPortal(<div className="tp-body">
    <div className="col">
      <button className="tp-link" type="button" onClick={() => L.panes?.task?.show('design')}><span>{translate(L, 'Design Templates')}</span></button>
      <button className="tp-link" type="button" onClick={() => L.panes?.task?.show('colors')}><span>{translate(L, 'Color Schemes')}</span></button>
      <button className="tp-link" type="button" onClick={() => L.panes?.task?.show('animSchemes')}><span>{translate(L, 'Animation Schemes')}</span></button>
    </div>
    {!!catalog.used.length && <section className="tp-sec"><div className="tp-h">{translate(L, 'Used in This Presentation')}</div><div className="dsn-grid">{catalog.used.map((design) => card(design, true))}</div></section>}
    <section className="tp-sec"><div className="tp-h">{translate(L, 'Available For Use')}</div><div className="dsn-grid">{catalog.available.map((design) => card(design, false))}</div></section>
    <button className="tp-link" type="button" onClick={() => L.panes?.reactBrowseDesign?.()}><span>{translate(L, 'Browse... (use the design of another .pptx/.potx)')}</span></button>
  </div>, mount);
}

function ColorSchemesPane({ L, mount }: { L: TaskContentRuntime; mount: HTMLElement }) {
  const data = L.panes?.reactColorCatalog?.();
  const schemes = [data?.current, ...(data?.schemes || [])].filter((scheme): scheme is ColorScheme => !!scheme);
  const slots = ['lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'hlink', 'folHlink'];
  const showMenu = (scheme: ColorScheme, event: MouseEvent) => { event.preventDefault(); L.panes?.reactColorMenu?.(scheme, event.clientX, event.clientY); };
  return createPortal(<div className="tp-body">
    <div className="col">
      <button className="tp-link" type="button" onClick={() => L.panes?.task?.show('design')}><span>{translate(L, 'Design Templates')}</span></button>
      <button className="tp-link" type="button" onClick={() => L.panes?.task?.show('colors')}><span>{translate(L, 'Color Schemes')}</span></button>
      <button className="tp-link" type="button" onClick={() => L.panes?.task?.show('animSchemes')}><span>{translate(L, 'Animation Schemes')}</span></button>
    </div>
    <section className="tp-sec"><div className="tp-h">{translate(L, 'Apply a color scheme:')}</div><div className="scheme-grid">
      {schemes.map((scheme, index) => {
        const label = index === 0 ? translate(L, 'Current scheme') : L.t?.('Color scheme {0}', index) || `Color scheme ${index}`;
        const name = index === 0 ? translate(L, 'Current color scheme') : translate(L, 'Apply this color scheme');
        return <button className="scheme-it" type="button" aria-label={label} data-tip={name} data-scheme-index={index} key={`${index}-${scheme.accent1 || ''}`} onClick={() => L.panes?.reactApplyColorScheme?.(scheme, true)} onContextMenu={(event) => showMenu(scheme, event)}>
          {slots.map((slot) => <span key={slot} style={{ background: scheme[slot] }} />)}
        </button>;
      })}
    </div></section>
    <button className="tp-link" type="button" onClick={() => L.panes?.reactEditColorScheme?.()}><span>{translate(L, 'Edit Color Schemes...')}</span></button>
  </div>, mount);
}

function AnimationSchemesPane({ L, mount }: { L: TaskContentRuntime; mount: HTMLElement }) {
  const schemes = L.panes?.reactAnimationSchemes?.() || [];
  const [selected, setSelected] = useState<number | null>(null);
  const [autoPreview, setAutoPreview] = useState(!!L.app?.opts?.autoPreview);
  const selectedScheme = schemes.find((scheme) => scheme.index === selected);
  const apply = (all: boolean) => L.panes?.reactApplyAnimationScheme?.(selectedScheme?.definition || null, all);
  const play = () => L.panes?.reactPlayAnimationPreview?.();
  return createPortal(<div className="tp-body">
    <div className="col">
      <button className="tp-link" type="button" onClick={() => L.panes?.task?.show('design')}><span>{translate(L, 'Design Templates')}</span></button>
      <button className="tp-link" type="button" onClick={() => L.panes?.task?.show('colors')}><span>{translate(L, 'Color Schemes')}</span></button>
      <button className="tp-link" type="button" onClick={() => L.panes?.task?.show('animSchemes')}><span>{translate(L, 'Animation Schemes')}</span></button>
    </div>
    <section className="tp-sec"><div className="tp-h">{translate(L, 'Apply to selected slides:')}</div><div className="list-box trans-list" style={{ maxHeight: 300 }} role="listbox" aria-label={translate(L, 'Animation Schemes')}>
      {schemes.map((scheme) => scheme.group
        ? <div className="li-group" key={`group-${scheme.index}`}>{translate(L, scheme.name)}</div>
        : <div className={`li${selected === scheme.index ? ' on' : ''}`} key={`scheme-${scheme.index}`} role="option" aria-selected={selected === scheme.index} tabIndex={0} data-scheme-index={scheme.index} onClick={() => { setSelected(scheme.index); L.panes?.reactApplyAnimationScheme?.(scheme.definition, false); if (autoPreview) play(); }} onKeyDown={(event) => { if (event.key === 'Enter' || event.key === ' ') { event.preventDefault(); setSelected(scheme.index); L.panes?.reactApplyAnimationScheme?.(scheme.definition, false); if (autoPreview) play(); } }}>{translate(L, scheme.name)}</div>)}
    </div></section>
    <div className="tp-btns"><button className="btn tp-btn" type="button" onClick={() => apply(true)}>{translate(L, 'Apply to All Slides')}</button></div>
    <div className="tp-btns"><button className="btn small" type="button" onClick={play}>▶ {translate(L, 'Play')}</button><button className="btn small" type="button" onClick={() => L.ui?.exec('showFromCurrent')}>{translate(L, 'Slide Show')}</button></div>
    <label className="ck"><input type="checkbox" checked={autoPreview} onChange={(event) => { const checked=event.currentTarget.checked; setAutoPreview(checked); if (L.app?.opts) L.app.opts.autoPreview=checked; L.app?.saveOpts?.(); }} /><span>{translate(L, 'AutoPreview')}</span></label>
  </div>, mount);
}

function SlideLayoutPane({ L, mount }: { L: TaskContentRuntime; mount: HTMLElement }) {
  const catalog = L.panes?.reactLayoutCatalog?.() || { currentKey: null, groups: [], showWhenInserting: false };
  const [showWhenInserting, setShowWhenInserting] = useState(catalog.showWhenInserting);
  const showMenu = (key: string, event: MouseEvent) => { event.preventDefault(); L.panes?.reactLayoutMenu?.(key, event.clientX, event.clientY); };
  return createPortal(<div className="tp-body">
    <div className="tp-note">{translate(L, 'Apply slide layout:')}</div>
    {catalog.groups.map((group) => <section className="tp-sec" key={group.name}><div className="tp-h">{translate(L, group.name)}</div><div className="lay-grid">
      {group.layouts.map((layout) => <button className={`lay-it${layout.key === catalog.currentKey ? ' on' : ''}`} type="button" data-key={layout.key} data-tip={layout.name} aria-label={translate(L, layout.name)} aria-current={layout.key === catalog.currentKey ? 'true' : undefined} key={layout.key} dangerouslySetInnerHTML={{ __html: layout.preview }} onClick={() => L.panes?.reactApplyLayout?.(layout.key)} onContextMenu={(event) => showMenu(layout.key, event)} />)}
    </div></section>)}
    <label className="ck"><input type="checkbox" checked={showWhenInserting} onChange={(event) => { const checked=event.currentTarget.checked; setShowWhenInserting(checked); L.panes?.reactSetLayoutPaneOnNew?.(checked); }} /><span>{translate(L, 'Show when inserting new slides')}</span></label>
  </div>, mount);
}

function SlideTransitionPane({ L, mount }: { L: TaskContentRuntime; mount: HTMLElement }) {
  const [, refresh] = useState(0);
  const catalog = L.panes?.reactTransitionCatalog?.() || { available: false, transitions: [], currentName: '', speed: 'fast', click: true, after: null, autoPreview: true };
  const [selectedName, setSelectedName] = useState(catalog.currentName);
  const [speed, setSpeed] = useState(catalog.speed);
  const [click, setClick] = useState(catalog.click);
  const [after, setAfter] = useState(catalog.after);
  const [automatic, setAutomatic] = useState(catalog.after != null);
  const [autoPreview, setAutoPreview] = useState(catalog.autoPreview);
  useEffect(() => L.bus?.on('slides-changed', () => refresh((value) => value + 1)), [L]);
  useEffect(() => {
    setSelectedName(catalog.currentName);
    setSpeed(catalog.speed);
    setClick(catalog.click);
    setAfter(catalog.after);
    setAutomatic(catalog.after != null);
  }, [catalog.currentName, catalog.speed, catalog.click, catalog.after]);
  if (!catalog.available) return createPortal(<div className="tp-body"><div className="tp-note">{translate(L, 'Transitions are set in Normal or Slide Sorter view.')}</div></div>, mount);
  const commitAfter = (value: number | null) => { setAfter(value); L.panes?.reactSetTransitionAfter?.(automatic ? value : null); };
  return createPortal(<div className="tp-body">
    <section className="tp-sec"><div className="tp-h">{translate(L, 'Apply to selected slides:')}</div><div className="list-box trans-list" style={{ height: 220 }} role="listbox" aria-label={translate(L, 'Slide Transitions')}>
      {catalog.transitions.map((transition) => <div className={`li${transition.name === selectedName ? ' on' : ''}`} key={`${transition.index}-${transition.name}`} role="option" aria-selected={transition.name === selectedName} tabIndex={0} data-transition-index={transition.index} onClick={() => { setSelectedName(transition.name); L.panes?.reactChooseTransition?.(transition); }} onKeyDown={(event) => { if (event.key === 'Enter' || event.key === ' ') { event.preventDefault(); setSelectedName(transition.name); L.panes?.reactChooseTransition?.(transition); } }}>{translate(L, transition.name)}</div>)}
    </div></section>
    <section className="tp-sec"><div className="tp-h">{translate(L, 'Modify transition')}</div>
      <div className="tp-row"><label htmlFor="react-transition-speed">{translate(L, 'Speed:')}</label><select id="react-transition-speed" value={speed} onChange={(event) => { const value=event.currentTarget.value; setSpeed(value); L.panes?.reactSetTransitionSpeed?.(value); }}><option value="slow">{translate(L, 'Slow')}</option><option value="med">{translate(L, 'Medium')}</option><option value="fast">{translate(L, 'Fast')}</option></select></div>
      <div className="tp-row"><label htmlFor="react-transition-sound">{translate(L, 'Sound:')}</label><select id="react-transition-sound" disabled value=""><option value="">{translate(L, '[No Sound]')}</option></select></div>
    </section>
    <section className="tp-sec"><div className="tp-h">{translate(L, 'Advance slide')}</div>
      <label className="ck"><input type="checkbox" checked={click} onChange={(event) => { const value=event.currentTarget.checked; setClick(value); L.panes?.reactSetTransitionClick?.(value); }} /><span>{translate(L, 'On mouse click')}</span></label>
      <div className="tp-row"><label htmlFor="react-transition-after">{translate(L, 'Automatically after')}</label><input id="react-transition-after" type="number" min="0" max="3600" step="1" value={after ?? 0} disabled={!automatic} onChange={(event) => { const value=Number(event.currentTarget.value); commitAfter(Number.isFinite(value) ? value : 0); }} /><span> s</span><input aria-label={translate(L, 'Enable automatic advance')} type="checkbox" checked={automatic} onChange={(event) => { const enabled=event.currentTarget.checked; setAutomatic(enabled); L.panes?.reactSetTransitionAfter?.(enabled ? after ?? 0 : null); }} /></div>
    </section>
    <div className="tp-btns"><button className="btn tp-btn" type="button" onClick={() => L.panes?.reactTransitionToAll?.()}>{translate(L, 'Apply to All Slides')}</button></div>
    <div className="tp-btns"><button className="btn small" type="button" onClick={() => L.panes?.reactPreviewTransition?.()}>▶ {translate(L, 'Play')}</button><button className="btn small" type="button" onClick={() => L.ui?.exec('showFromCurrent')}>{translate(L, 'Slide Show')}</button></div>
    <label className="ck"><input type="checkbox" checked={autoPreview} onChange={(event) => { const enabled=event.currentTarget.checked; setAutoPreview(enabled); L.panes?.reactSetAutoPreview?.(enabled); }} /><span>{translate(L, 'AutoPreview')}</span></label>
  </div>, mount);
}

function CustomAnimationPane({ L, mount }: { L: TaskContentRuntime; mount: HTMLElement }) {
  const [, refresh] = useState(0);
  const catalog = L.panes?.reactCustomAnimationCatalog?.() || { available: false, animations: [], selectedId: null, hasSelection: false, starts: [], directions: [], speeds: [], autoPreview: false };
  const [autoPreview, setAutoPreview] = useState(catalog.autoPreview);
  useEffect(() => L.bus?.on('pptx-task-render', () => refresh((value) => value + 1)), [L]);
  useEffect(() => setAutoPreview(catalog.autoPreview), [catalog.autoPreview]);
  if (!catalog.available) return createPortal(<div className="tp-body"><div className="tp-note">{translate(L, 'Custom animation is available in Normal view.')}</div></div>, mount);
  const selected = catalog.animations.find((animation) => animation.id === catalog.selectedId);
  const chooseStart = (value: string) => L.panes?.reactUpdateAnimation?.('start', value);
  return createPortal(<div className="tp-body">
    <div className="tp-row"><button className="btn small" type="button" aria-label={translate(L, 'Add Effect')} onClick={(event) => { const rect=event.currentTarget.getBoundingClientRect(); L.panes?.reactAddAnimationMenu?.(rect.left, rect.bottom); }}>{translate(L, 'Add Effect')}</button><button className="btn small" type="button" disabled={!selected} onClick={() => selected && L.panes?.reactRemoveAnimation?.(selected.id)}>{translate(L, 'Remove')}</button></div>
    <section className="tp-sec"><div className="tp-h">{translate(L, selected ? 'Modify: {0}' : 'Modify effect', selected?.name || '')}</div>
      <div className="tp-row"><label htmlFor="react-animation-start">{translate(L, 'Start:')}</label><select id="react-animation-start" value={selected?.start || 'click'} disabled={!selected} onChange={(event) => chooseStart(event.currentTarget.value)}>{catalog.starts.map((option) => <option key={option.value} value={option.value}>{translate(L, option.name)}</option>)}</select></div>
      <div className="tp-row"><label htmlFor="react-animation-direction">{translate(L, 'Direction:')}</label><select id="react-animation-direction" value={selected?.dir || ''} disabled={!selected || catalog.directionDisabled} onChange={(event) => L.panes?.reactUpdateAnimation?.('dir', event.currentTarget.value)}>{(catalog.directions.length ? catalog.directions : [{ value: '', name: '' }]).map((option) => <option key={option.value} value={option.value}>{translate(L, option.name)}</option>)}</select></div>
      <div className="tp-row"><label htmlFor="react-animation-speed">{translate(L, 'Speed:')}</label><select id="react-animation-speed" value={catalog.selectedSpeed || 1000} disabled={catalog.speedDisabled} onChange={(event) => L.panes?.reactUpdateAnimation?.('dur', Number(event.currentTarget.value))}>{catalog.speeds.map((option) => <option key={option.value} value={option.value}>{translate(L, option.name)}</option>)}</select></div>
    </section>
    <div className="list-box" style={{ minHeight: 150 }} role="listbox" aria-label={translate(L, 'Custom Animation')}>
      {catalog.animations.map((animation) => <div className={`li${animation.id === catalog.selectedId ? ' on' : ''}`} key={animation.id} role="option" aria-selected={animation.id === catalog.selectedId} tabIndex={0} data-animation-id={animation.id} title={animation.tooltip} onClick={() => L.panes?.reactSelectAnimation?.(animation.id)} onDoubleClick={() => L.panes?.reactAnimationOptions?.(animation.id)} onContextMenu={(event) => { event.preventDefault(); L.panes?.reactAnimationMenu?.(animation.id, event.clientX, event.clientY); }}>
        <span className="li-num">{animation.step}</span><span className="li-ico" style={{ color: animation.color }}>★</span><span className="li-n">{animation.name}{animation.by === 'para' ? ' ¶' : ''} — {animation.label}</span>
      </div>)}
      {!catalog.animations.length && <div className="tp-note" style={{ padding: 8 }}>{translate(L, 'Select an element of the slide, then click "Add Effect" to add animation.')}</div>}
    </div>
    <div className="tp-row"><span>{translate(L, 'Re-Order')}</span><button className="btn small" type="button" aria-label={translate(L, 'Move up')} disabled={!selected || catalog.animations[0]?.id === selected.id} onClick={() => L.panes?.reactMoveAnimation?.(-1)}>↑</button><button className="btn small" type="button" aria-label={translate(L, 'Move down')} disabled={!selected || catalog.animations.at(-1)?.id === selected.id} onClick={() => L.panes?.reactMoveAnimation?.(1)}>↓</button></div>
    <div className="tp-btns"><button className="btn small" type="button" onClick={() => L.panes?.reactPlayAnimation?.()}>▶ {translate(L, 'Play')}</button><button className="btn small" type="button" onClick={() => L.ui?.exec('showFromCurrent')}>{translate(L, 'Slide Show')}</button></div>
    <label className="ck"><input type="checkbox" checked={autoPreview} onChange={(event) => { const enabled=event.currentTarget.checked; setAutoPreview(enabled); L.panes?.reactSetAnimationAutoPreview?.(enabled); }} /><span>{translate(L, 'AutoPreview')}</span></label>
  </div>, mount);
}

function ClipboardPane({ L, mount }: { L: TaskContentRuntime; mount: HTMLElement }) {
  const [, refresh] = useState(0);
  useEffect(() => L.bus?.on('pptx-task-render', () => refresh((value) => value + 1)), [L]);
  const items = L.panes?.reactClipboardCatalog?.() || [];
  return createPortal(<div className="tp-body">
    <div className="tp-btns"><button className="btn small" type="button" onClick={() => L.panes?.reactPasteAllClipboard?.()}>{translate(L, 'Paste All')}</button><button className="btn small" type="button" onClick={() => L.panes?.reactClearClipboard?.()}>{translate(L, 'Clear All')}</button></div>
    <div className="tp-note">{translate(L, '{0} of 24 — Click an item to paste.', items.length)}</div>
    {items.map((item) => <button className="cb-item" type="button" key={`${item.index}-${item.label}`} onClick={() => L.panes?.reactPasteClipboard?.(item.index)}>
      <span className="cb-ico" dangerouslySetInnerHTML={{ __html: L.icons?.get(item.icon) || '' }} /><span className="cb-t" translate="no">{item.label}</span>
    </button>)}
  </div>, mount);
}

function HelpPane({ L, task, mount }: { L: TaskContentRuntime; task: NonNullable<TaskContentRuntime['panes']>['task']; mount: HTMLElement }) {
  const [query, setQuery] = useState(L.panes?.helpQuery || '');
  const data = task?.helpData;
  const words = query.toLowerCase().split(/\s+/).filter(Boolean);
  const matches = (text: string) => words.every((word) => text.toLowerCase().includes(word));
  const topics = (data?.topics || []).map(([title, description]) => [translate(L, title), translate(L, description)] as const).filter(([title, description]) => !words.length || matches(`${title} ${description}`));
  const keys = (data?.keys || []).map(([key, description]) => [key.split(' / ').map((part) => / [a-z]/.test(part) ? translate(L, part) : L.ui?.keyText?.(part) || part).join(' / '), translate(L, description)] as const).filter(([key, description]) => !words.length || matches(`${key} ${description}`));
  return createPortal(<div className="tp-body">
    <div className="tp-row"><input id="hp-q" type="text" value={query} placeholder={translate(L, 'Search')} style={{ flex: 1, minWidth: 0 }} onChange={(event) => { setQuery(event.currentTarget.value); if (L.panes) L.panes.helpQuery = event.currentTarget.value; }} /></div>
    {topics.map(([title, description]) => <div className="help-topic" key={title}><h4>{title}</h4><p>{description}</p></div>)}
    {!!keys.length && <div className="help-topic"><h4>{translate(L, 'Keyboard shortcuts')}</h4><table className="kbd-table"><tbody>{keys.map(([key, description]) => <tr key={key}><td translate="no">{key}</td><td>{description}</td></tr>)}</tbody></table></div>}
    {!topics.length && !keys.length && <div className="tp-note">{translate(L, 'No topics match. Try a different word.')}</div>}
  </div>, mount);
}

function ContextualToolsPane({ L, mount, context }: { L: TaskContentRuntime; mount: HTMLElement; context: RibbonContext }) {
  return createPortal(<div className="tp-body tp-tools-panel">
    {context.tabs.flatMap((tab) => tab.groups).map((group, groupIndex) => <section className="tp-sec" key={`${context.id}-${groupIndex}`}>
      <div className="tp-h">{translate(L, group.label)}</div>
      <div className="tp-tools">
        {group.items.filter((item) => item !== '|').map((item, index) => <ContextualRibbonItem key={`${groupIndex}-${index}`} L={L} item={item as RibbonItem} size="medium" />)}
        {group.launcher && <ContextualRibbonItem L={L} item={{ cmd: group.launcher }} size="medium" />}
      </div>
    </section>)}
  </div>, mount);
}

function ContextualRibbonItem({ L, item: value, size }: { L: TaskContentRuntime; item: RibbonItem; size: 'large' | 'medium' | 'small' }) {
  const item = typeof value === 'string' ? { cmd: value } : value;
  if (item.row) return <div className="rb-row">{item.row.map((child, index) => child === '|' ? <span className="rb-gap" key={`gap-${index}`} /> : <ContextualRibbonItem L={L} item={child} size="small" key={`item-${index}`} />)}</div>;
  if (item.make) return <ContextualControlSlot make={item.make} size={item.size || size} />;

  const commandId = item.cmd || item.split;
  const command = commandId ? L.ui?.cmds?.[commandId] : undefined;
  const actualSize = item.size || size;
  const rawLabel = actualSize === 'small' ? '' : item.label ?? (item.drop || (commandId ? command?.label || commandId : ''));
  const label = rawLabel ? (L.ui?.stripAmp?.(rawLabel) || rawLabel.replace('&', '')) : '';
  const iconName = item.icon || (commandId ? L.ribbonSpec?.().icons[commandId] || command?.icon : undefined);
  const icon = iconName ? L.icons?.get(iconName, actualSize === 'large' ? 32 : undefined) || '' : '';
  const tip = L.ui?.stripAmp?.(command?.tip || command?.label || item.tip || item.drop || label || commandId || '') || item.tip || label;
  const openMenu = (anchor: HTMLElement) => {
    if (!item.menu) return;
    const rect = anchor.getBoundingClientRect();
    item.menu({ left: rect.left, bottom: rect.bottom }, anchor);
  };
  const iconLabel = <>{icon && <span className="tb-icon" aria-hidden="true" dangerouslySetInnerHTML={{ __html: icon }} />}{label && <span className="tb-lbl">{translate(L, label)}</span>}</>;

  if (item.split && commandId) return <div className={`tb-split rb-${actualSize}`}>
    <button type="button" className={`tb-btn${label ? ' with-label' : ''}`} data-cmd={commandId} data-tip={tip} aria-label={L.ui?.stripAmp?.(command?.label || label || commandId) || label || commandId} aria-pressed={command?.checked ? !!command.checked() : undefined} disabled={command?.enabled ? !command.enabled() : false} onPointerDown={(event) => event.preventDefault()} onClick={(event) => L.ui?.exec(commandId, item.arg, event.nativeEvent)}>{iconLabel}</button>
    <button type="button" className="tb-dd" aria-label={translate(L, 'More {0} options', label || commandId)} aria-haspopup="menu" data-tip={translate(L, 'More {0} options', label || commandId)} onPointerDown={(event) => event.preventDefault()} onClick={(event) => openMenu(event.currentTarget.parentElement as HTMLElement)}><span className="dd-arrow" /></button>
  </div>;
  if (item.drop) return <button type="button" className={`tb-btn tb-drop rb-${actualSize}${label ? ' with-label' : ''}`} data-tip={tip} aria-label={tip} aria-haspopup="menu" onPointerDown={(event) => event.preventDefault()} onClick={(event) => openMenu(event.currentTarget)}>{iconLabel}<span className="dd-arrow" /></button>;
  if (!commandId) return null;
  const enabled = (() => { try { return command?.enabled ? !!command.enabled() : true; } catch { return false; } })();
  const checked = (() => { try { return !!command?.checked?.(); } catch { return false; } })();
  return <button type="button" className={`tb-btn rb-${actualSize}${label ? ' with-label' : ''}${checked ? ' on' : ''}${enabled ? '' : ' dis'}`} data-cmd={commandId} data-tip={tip} aria-label={L.ui?.stripAmp?.(command?.label || label || commandId) || label || commandId} aria-pressed={command?.checked ? checked : undefined} disabled={!enabled} onPointerDown={(event) => event.preventDefault()} onClick={(event) => L.ui?.exec(commandId, item.arg, event.nativeEvent)}>{iconLabel}</button>;
}

function ContextualControlSlot({ make, size }: { make: () => HTMLElement; size: 'large' | 'medium' | 'small' }) {
  const [host] = useState(() => document.createElement('span'));
  useEffect(() => {
    const control = make();
    control.classList.add(`rb-${size}`);
    host.replaceChildren(control);
    return () => { control.remove(); host.replaceChildren(); };
  }, [host, make, size]);
  return <span className="vibeoffice-ribbon-control" style={{ display: 'contents' }} ref={(element) => { if (element && element.lastElementChild !== host) element.append(host); }} />;
}

function IntelliArtTextPane({ L, mount, revision }: { L: TaskContentRuntime; mount: HTMLElement; revision: number }) {
  const catalog = L.diagram?.reactTextPaneCatalog?.() || { available: false, rows: [], layouts: [], colors: [], styles: [] };
  const [drafts, setDrafts] = useState<Record<string, string>>({});
  const inputs = useRef(new Map<string, HTMLInputElement>());
  const signature = JSON.stringify(catalog.rows);
  useEffect(() => {
    setDrafts(Object.fromEntries(catalog.rows.map((row) => [row.id, row.text])));
  }, [revision, signature]);
  const operation = (id: string, label: string) => <button className="tb-btn rb-small" type="button" data-cmd={id} aria-label={translate(L, label)} onPointerDown={(event) => event.preventDefault()} onClick={() => L.ui?.exec(id)}>{translate(L, label)}</button>;
  const choices = (items: DiagramTextChoice[], kind: 'layout' | 'colors' | 'style') => <div className={`sa-pick sa-${kind}`}>
    {items.map((item) => <button className={`sa-cell${item.selected ? ' on' : ''}`} type="button" key={item.id} data-choice={item.id} aria-label={item.name} aria-pressed={item.selected} data-tip={item.name} dangerouslySetInnerHTML={{ __html: item.preview }} onClick={() => L.diagram?.reactTextSet?.(kind, item.id)} />)}
  </div>;
  if (!catalog.available) return createPortal(<div className="tp-body">
    <div className="tp-note">{translate(L, 'Select an IntelliArt graphic to type its text here.')}</div>
    <div className="tp-row"><button className="btn" type="button" onClick={() => L.ui?.exec('insertDiagram')}>{translate(L, 'Insert IntelliArt...')}</button></div>
  </div>, mount);
  const toolbar: Array<[string, string]> = [
    ['diagramAddShape', 'Add Shape After'], ['diagramPromote', 'Promote'], ['diagramDemote', 'Demote'], ['diagramMoveUp', 'Move Up'], ['diagramMoveDown', 'Move Down'],
  ];
  const structurallyNavigating = (event: KeyboardEvent) => event.key === 'Enter' || event.key === 'Tab' || event.key === 'ArrowUp' || event.key === 'ArrowDown' || event.key === 'Escape' || (event.key === 'Backspace');
  return createPortal(<div className="tp-body">
    <div className="tp-h">{translate(L, 'Type your text here')}</div>
    <div className="sa-tools">{toolbar.map(([id, fallback]) => operation(id, L.ui?.cmds?.[id]?.label || fallback))}</div>
    <div className="sa-pane" role="tree" aria-label={translate(L, 'IntelliArt text')}>
      {catalog.rows.map((row) => <div className="sa-line" role="treeitem" key={row.id}>
        <span className="sa-bul" aria-hidden="true" style={{ marginLeft: row.level * 16 }}>•</span>
        <input ref={(element) => { if (element) inputs.current.set(row.id, element); else inputs.current.delete(row.id); }} type="text" className="sa-row" value={drafts[row.id] ?? row.text} placeholder="[Text]" aria-label={L.t?.('Item, level {0}', row.level + 1) || `Item, level ${row.level + 1}`} data-item={row.id} style={{ marginLeft: row.level * 16 }} onFocus={() => L.diagram?.reactTextFocus?.(row.id)} onChange={(event) => { const value = event.currentTarget.value; setDrafts((previous) => ({ ...previous, [row.id]: value })); L.diagram?.reactTextChange?.(row.id, value); }} onKeyDown={(event) => { if (structurallyNavigating(event.nativeEvent)) { L.diagram?.reactTextKey?.(event.nativeEvent, row.id, event.currentTarget); const nextId = L.diagram?.reactTextPaneCatalog?.().focusId; if (nextId) requestAnimationFrame(() => { const next = inputs.current.get(nextId); next?.focus(); next?.setSelectionRange(next.value.length, next.value.length); }); } }} />
      </div>)}
    </div>
    <div className="tp-note">{translate(L, 'Enter adds an item, Tab and Shift+Tab move it in and out a level, Alt+Shift+↑/↓ move it.')}</div>
    {!!catalog.layouts.length && <><div className="tp-h">{translate(L, 'Layout')}</div>{choices(catalog.layouts, 'layout')}<div className="tp-row"><button className="btn small" type="button" onClick={() => L.ui?.exec('diagramLayout')}>{translate(L, 'More Layouts...')}</button></div></>}
    {!!catalog.colors.length && <><div className="tp-h">{translate(L, 'Colors')}</div>{choices(catalog.colors, 'colors')}</>}
    {!!catalog.styles.length && <><div className="tp-h">{translate(L, 'IntelliArt Styles')}</div>{catalog.styles.map((group) => <div key={group.language}><div className="sa-sub">{translate(L, group.title)}</div>{choices(group.choices, 'style')}</div>)}</>}
    <div className="tp-row"><button className="btn small" type="button" onClick={() => L.diagram?.reactTextReset?.()}>{translate(L, 'Reset Graphic')}</button><button className="btn small" type="button" onClick={() => L.diagram?.reactTextConvert?.()}>{translate(L, 'Convert to Shapes')}</button></div>
  </div>, mount);
}
