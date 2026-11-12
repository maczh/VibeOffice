import { useEffect, useRef, useState, type PointerEvent as ReactPointerEvent, type MouseEvent as ReactMouseEvent } from 'react';
import { createPortal } from 'react-dom';

interface SlideLike { id?: string; hidden?: boolean; trans?: { type?: string; after?: number }; anims?: unknown[] }
interface PptxPaneRuntime {
  pres?: { slides: SlideLike[]; firstNum?: number };
  ed?: { idx: number; view: string };
  app?: { slideSel: Set<number>; reactMounted?: boolean; opts?: { outlinePlain?: boolean; spell?: boolean }; contextThumb(event: MouseEvent): void; setView(view: string): void; toggleLeft(show: boolean): void; toggleTask(show: boolean): void };
  render?: { thumb(presentation: NonNullable<PptxPaneRuntime['pres']>, slide: SlideLike, width: number, options: { index: number }): HTMLElement };
  panes?: {
    left?: { tab: string; body?: HTMLElement; tabs?: HTMLElement; list?: HTMLElement | null; outline?: HTMLElement | null; setTab(tab: string): void; reactOutlineRows?(): Array<{ kind: string; slideId?: string; level: number; text: string }>; onOutlineInput?(): void; onOutlinePaste?(event: ClipboardEvent): void; onOutlineKey?(event: KeyboardEvent): void; syncOutlineSlide?(): void; styleOutline?(): void; commitOutline?(keepFocus?: boolean): void; onThumbDown?(event: PointerEvent, index: number): void; onThumbContext?(event: MouseEvent, index: number): void; onThumbDoubleClick?(): void };
    sorter?: { el?: HTMLElement; reactMounted?: boolean; onDown?(event: PointerEvent, index: number, item: HTMLElement): void; mark(): void };
    task?: { stack: string[]; pos: number; current?: string; registry: Record<string, { id: string; title: string; hidden?: boolean }>; show(id: string): void; go(position: number): void };
  };
  ui?: { openMenu(items: Array<{ label: string; run(): void; checked(): boolean }>, position: { left: number; bottom: number }): void };
  show?: { previewTransition(frame: Element | null, presentation: NonNullable<PptxPaneRuntime['pres']>, slide: SlideLike, previous?: SlideLike): void };
  anim?: { transName(transition: NonNullable<SlideLike['trans']>): string };
  bus?: { on(event: string, listener: () => void): () => void; emit(event: string): void };
  icons?: { get(name: string, size?: number): string };
  t?(message: string, ...args: unknown[]): string;
}

function runtime(): PptxPaneRuntime | undefined { return (window as Window & { L?: PptxPaneRuntime }).L; }

/** React-owned slide thumbnail list. The document engine still supplies rendered slide artwork and editing commands. */
export function PptxSlidesLayer({ host }: { host: HTMLElement }) {
  const [mount, setMount] = useState<HTMLElement | null>(null);
  const [tabsMount, setTabsMount] = useState<HTMLElement | null>(null);
  const [taskMount, setTaskMount] = useState<HTMLElement | null>(null);
  const [sorterMount, setSorterMount] = useState<HTMLElement | null>(null);
  const [revision, setRevision] = useState(0);
  const outlineMutated = useRef(false);
  useEffect(() => {
    let slot: HTMLElement | null = null;
    let tabsSlot: HTMLElement | null = null;
    let taskSlot: HTMLElement | null = null;
    let sorterSlot: HTMLElement | null = null;
    let offRender: (() => void) | undefined;
    let offMark: (() => void) | undefined;
    let offTask: (() => void) | undefined;
    let offSorterRender: (() => void) | undefined;
    let offSorterMark: (() => void) | undefined;
    let subscribed = false;
    const attach = () => {
      const target = host.querySelector<HTMLElement>('#react-slides-root');
      const tabsTarget = host.querySelector<HTMLElement>('#react-slides-tabs-root');
      const taskTarget = host.querySelector<HTMLElement>('#react-task-head-root');
      const sorterTarget = host.querySelector<HTMLElement>('.sorter');
      const engine = runtime();
      if (!engine) return;
      if (!subscribed && engine.bus) {
        const refresh = () => { if (!outlineMutated.current) setRevision((value) => value + 1); };
        offRender = engine.bus.on('pptx-slides-render', refresh);
        offMark = engine.bus.on('pptx-slides-mark', refresh);
        offTask = engine.bus.on('pptx-task-render', refresh);
        offSorterRender = engine.bus.on('pptx-sorter-render', refresh);
        offSorterMark = engine.bus.on('pptx-sorter-mark', refresh);
        subscribed = true;
      }
      if (target && target !== slot) {
        slot = target;
        target.classList.add('vibeoffice-react-owned');
        if (engine.app) engine.app.reactMounted = true;
        setMount(target);
        engine.bus?.emit('pptx-slides-render');
      }
      if (tabsTarget && tabsTarget !== tabsSlot) {
        tabsSlot = tabsTarget;
        tabsTarget.classList.add('vibeoffice-react-owned');
        host.querySelector<HTMLElement>('#leftpane > .lp-tabs')?.classList.add('react-pane-adapter');
        setTabsMount(tabsTarget);
      }
      if (taskTarget && taskTarget !== taskSlot) {
        taskSlot = taskTarget;
        taskTarget.classList.add('vibeoffice-react-owned');
        host.querySelector<HTMLElement>('#taskpane > .tp-head')?.classList.add('react-pane-adapter');
        setTaskMount(taskTarget);
      }
      if (sorterTarget && sorterTarget !== sorterSlot) {
        sorterSlot?.classList.remove('vibeoffice-react-owned');
        sorterSlot = sorterTarget;
        sorterTarget.classList.add('vibeoffice-react-owned');
        const sorter = engine.panes?.sorter;
        if (sorter) sorter.reactMounted = true;
        setSorterMount(sorterTarget);
        engine.bus?.emit('pptx-sorter-render');
      }
    };
    const observer = new MutationObserver(attach);
    observer.observe(host, { childList: true, subtree: true });
    attach();
    return () => {
      observer.disconnect(); offRender?.(); offMark?.(); offTask?.(); offSorterRender?.(); offSorterMark?.();
      slot?.classList.remove('vibeoffice-react-owned');
      tabsSlot?.classList.remove('vibeoffice-react-owned');
      host.querySelector<HTMLElement>('#leftpane > .lp-tabs')?.classList.remove('react-pane-adapter');
      taskSlot?.classList.remove('vibeoffice-react-owned');
      sorterSlot?.classList.remove('vibeoffice-react-owned');
      host.querySelector<HTMLElement>('#taskpane > .tp-head')?.classList.remove('react-pane-adapter');
      const app = runtime()?.app;
      if (app) app.reactMounted = false;
      const sorter = runtime()?.panes?.sorter;
      if (sorter) sorter.reactMounted = false;
    };
  }, [host]);
  const L = runtime();
  const pane = L?.panes?.left;
  const task = L?.panes?.task;
  const sorter = L?.panes?.sorter;
  const presentation = L?.pres;
  const slides = presentation?.slides;
  void revision;
  if (!L || !pane) return null;
  const taskId = task ? task.stack[task.pos] : undefined;
  const taskTitle = taskId ? task?.registry[taskId]?.title : undefined;
  const taskHeader = taskMount && task && createPortal(<div className="tp-head">
    <button className="tp-nav" type="button" aria-label={L.t?.('Back') || 'Back'} disabled={task.pos <= 0} dangerouslySetInnerHTML={{ __html: L.icons?.get('back') || '' }} onClick={() => task.go(task.pos - 1)} />
    <button className="tp-nav" type="button" aria-label={L.t?.('Forward') || 'Forward'} disabled={task.pos >= task.stack.length - 1} dangerouslySetInnerHTML={{ __html: L.icons?.get('forward') || '' }} onClick={() => task.go(task.pos + 1)} />
    <button className="tp-nav" type="button" aria-label={L.t?.('Home') || 'Home'} dangerouslySetInnerHTML={{ __html: L.icons?.get('home') || '' }} onClick={() => task.show('getting-started')} />
    <button className="tp-title" type="button" aria-haspopup="menu" onClick={(event) => {
      const rect = event.currentTarget.getBoundingClientRect();
      L.ui?.openMenu(Object.values(task.registry).filter((item) => !item.hidden).map((item) => ({ label: item.title, run: () => task.show(item.id), checked: () => task.current === item.id })), { left: rect.left, bottom: rect.bottom });
    }}><span>{taskTitle ? L.t?.(taskTitle) || taskTitle : ''}</span><span className="dd-arrow" /></button>
    <button className="tp-nav" type="button" aria-label={L.t?.('Close task pane') || 'Close task pane'} dangerouslySetInnerHTML={{ __html: L.icons?.get('close', 12) || '' }} onClick={() => L.app?.toggleTask(false)} />
  </div>, taskMount);
  const tabs = tabsMount && createPortal(<div className="lp-tabs" role="tablist" aria-label="Presentation panes">
    <button className={`lp-tab${pane.tab === 'outline' ? ' on' : ''}`} type="button" role="tab" aria-selected={pane.tab === 'outline'} onClick={() => { if (pane.tab !== 'outline') outlineMutated.current = false; pane.setTab('outline'); }}>{L.t?.('Outline|pane') || 'Outline'}</button>
    <button className={`lp-tab${pane.tab === 'slides' ? ' on' : ''}`} type="button" role="tab" aria-selected={pane.tab === 'slides'} onClick={() => { outlineMutated.current = false; pane.setTab('slides'); }}>{L.t?.('Slides') || 'Slides'}</button>
    <button className="lp-close" type="button" aria-label={L.t?.('Close pane') || 'Close pane'} dangerouslySetInnerHTML={{ __html: L.icons?.get('close', 12) || '' }} onClick={() => L.app?.toggleLeft(false)} />
  </div>, tabsMount);
  const renderSlide = L.render?.thumb;
  const sorterView = sorterMount && sorter && presentation && slides && renderSlide && createPortal(<div className="sorter-grid" role="listbox" aria-label={L.t?.('Slide Sorter') || 'Slide Sorter'}>
    {slides.map((slide, index) => {
      const width = Math.max(90, Math.round(sorterMount.clientWidth ? Math.min(260, (sorterMount.clientWidth - 60) / 4) : 180));
      const markup = renderSlide(presentation, slide, width, { index }).outerHTML;
      return <div className={`so-it${L.app?.slideSel?.has(index) ? ' sel' : ''}${slide.hidden ? ' hiddenslide' : ''}`} data-i={index} key={slide.id || index} role="option" aria-selected={L.app?.slideSel?.has(index) || false} onPointerDown={(event: ReactPointerEvent<HTMLDivElement>) => sorter.onDown?.(event.nativeEvent, index, event.currentTarget)} onDoubleClick={() => { if (L.ed) L.ed.idx = index; L.app?.setView('normal'); }} onContextMenu={(event: ReactMouseEvent<HTMLDivElement>) => { event.preventDefault(); if (!L.app?.slideSel.has(index)) { if (L.app) L.app.slideSel = new Set([index]); if (L.ed) L.ed.idx = index; sorter.mark(); } L.app?.contextThumb(event.nativeEvent); }}>
        <button className="so-frame" type="button" aria-label={L.t?.('Slide {0}', index + 1) || `Slide ${index + 1}`} dangerouslySetInnerHTML={{ __html: markup }} />
        <div className="so-foot">{slide.trans?.type && slide.trans.type !== 'none' && <button type="button" className="th-star" title={L.anim?.transName(slide.trans)} aria-label={L.t?.('Preview transition') || 'Preview transition'} onPointerDown={(event) => event.stopPropagation()} onClick={(event) => { event.stopPropagation(); const frame = event.currentTarget.closest('.so-it')?.querySelector('.so-frame') || null; L.show?.previewTransition(frame, presentation, slide, slides[index - 1]); }}><svg width="12" height="12" viewBox="0 0 12 12" aria-hidden="true"><path d="M6 .8l1.5 3.2 3.4.4-2.5 2.3.7 3.4L6 8.4 2.9 10.1l.7-3.4L1.1 4.4l3.4-.4z" fill="none" stroke="#333" /></svg></button>}{slide.trans?.after != null && <span>{formatTransitionDelay(slide.trans.after)}</span>}<span className={`so-num${slide.hidden ? ' hid' : ''}`}>{String(index + (presentation.firstNum || 1))}</span></div>
      </div>;
    })}
  </div>, sorterMount);
  const outlineRows = pane.tab === 'outline' ? pane.reactOutlineRows?.() : undefined;
  const outlineView = mount && outlineRows && presentation && createPortal(<div className={`outline${L.app?.opts?.outlinePlain ? ' plain' : ''}`} contentEditable suppressContentEditableWarning spellCheck={!!L.app?.opts?.spell} role="textbox" aria-multiline="true" aria-label={L.t?.('Outline|pane') || 'Outline'} ref={(element) => { if (element) { pane.outline = element; pane.styleOutline?.(); } else if (pane.outline?.id === 'react-outline-editor') pane.outline = null; }} id="react-outline-editor" onInput={() => { outlineMutated.current = true; pane.onOutlineInput?.(); }} onPaste={(event) => pane.onOutlinePaste?.(event.nativeEvent)} onBlur={() => pane.commitOutline?.()} onKeyUp={() => pane.syncOutlineSlide?.()} onClick={() => pane.syncOutlineSlide?.()} onKeyDown={(event) => pane.onOutlineKey?.(event.nativeEvent)}>
    {outlineRows.map((row, index) => {
      const title = row.kind === 'title';
      const titleNumber = (presentation.firstNum || 1) + outlineRows.slice(0, index).filter((item) => item.kind === 'title').length;
      const style = title ? { fontWeight: 'bold' as const, marginTop: 6, paddingLeft: 36, position: 'relative' as const, minHeight: '1.35em' } : { paddingLeft: 48 + row.level * 18, position: 'relative' as const, minHeight: '1.35em', fontWeight: 'normal' as const };
      return <div className={`ol-row ${title ? 'title' : 'body'}`} data-slide={title ? row.slideId : undefined} data-lvl={title ? undefined : row.level} data-num={title ? titleNumber : undefined} style={style} key={`${row.slideId || 'slide'}-${row.kind}-${index}`}>{row.text || <br />}</div>;
    })}
  </div>, mount);
  if (pane.tab === 'outline' && L.ed?.view !== 'master') return <>{taskHeader}{tabs}{outlineView}{sorterView}</>;
  if (!mount || !presentation || !slides || pane.tab !== 'slides' || L.ed?.view === 'master') return <>{taskHeader}{tabs}{sorterView}</>;
  const body = pane.body;
  const width = Math.max(60, (body?.clientWidth || 190) - 48);
  if (!renderSlide) return <>{taskHeader}{tabs}{sorterView}</>;
  return <>{taskHeader}{tabs}{sorterView}{createPortal(<div className="thumbs" role="listbox" aria-label="Slides" ref={(element) => { pane.list = element; }}>
    {slides.map((slide, index) => {
      const markup = renderSlide(presentation, slide, width, { index }).outerHTML;
      const selected = L.app?.slideSel?.has(index) || false;
      const current = L.ed?.idx === index;
      const hasAnimation = (slide.trans?.type && slide.trans.type !== 'none') || !!slide.anims?.length;
      return <div className={`th${current ? ' cur' : ''}${selected ? ' sel' : ''}${slide.hidden ? ' hiddenslide' : ''}`} data-i={index} key={slide.id || index} role="option" aria-selected={current} onPointerDown={(event: ReactPointerEvent<HTMLDivElement>) => pane.onThumbDown?.(event.nativeEvent, index)} onContextMenu={(event: ReactMouseEvent<HTMLDivElement>) => pane.onThumbContext?.(event.nativeEvent, index)} onDoubleClick={() => pane.onThumbDoubleClick?.()}>
        <div className="th-meta"><span className="th-num">{String(index + (presentation.firstNum || 1))}</span>{hasAnimation && <button type="button" className="th-star" data-tip="Play animations" aria-label="Play animations" onClick={(event) => { event.stopPropagation(); const frame = event.currentTarget.closest('.th')?.querySelector('.th-frame') || null; L.show?.previewTransition(frame, presentation, slide, slides[index - 1]); }}><svg width="12" height="12" viewBox="0 0 12 12" aria-hidden="true"><path d="M6 .8l1.5 3.2 3.4.4-2.5 2.3.7 3.4L6 8.4 2.9 10.1l.7-3.4L1.1 4.4l3.4-.4z" fill="none" stroke="#333" /></svg></button>}</div>
        <button className="th-frame" type="button" aria-label={L.t?.('Slide {0}', index + 1) || `Slide ${index + 1}`} dangerouslySetInnerHTML={{ __html: markup }} />
      </div>;
    })}
  </div>, mount)}</>;
}

function formatTransitionDelay(milliseconds: number): string {
  const seconds = Math.round(milliseconds / 1000);
  return `${String(Math.floor(seconds / 60)).padStart(2, '0')}:${String(seconds % 60).padStart(2, '0')}`;
}
