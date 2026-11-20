import { createElement, useEffect, useMemo, useRef, useState, type CSSProperties, type PointerEvent as ReactPointerEvent, type MouseEvent as ReactMouseEvent, type KeyboardEvent as ReactKeyboardEvent, type ReactNode } from 'react';
import { createPortal } from 'react-dom';

interface SlideLike { id?: string; hidden?: boolean; trans?: { type?: string; after?: number }; anims?: unknown[] }
interface PptxPaneRuntime {
  pres?: { slides: SlideLike[]; firstNum?: number };
  ed?: { idx: number; view: string; masterKind?: string; editing?(): boolean; goto?(index: number): void };
  te?: { end?(): void };
  app?: { slideSel: Set<number>; focusArea?: string | null; reactMounted?: boolean; opts?: { outlinePlain?: boolean; spell?: boolean }; contextThumb(event: MouseEvent): void; moveSlides?(indices: number[], target: number): void; setView(view: string): void; toggleLeft(show: boolean): void; toggleTask(show: boolean): void };
  render?: { thumb(presentation: NonNullable<PptxPaneRuntime['pres']>, slide: SlideLike, width: number, options: { index: number }): HTMLElement };
  panes?: {
    left?: { tab: string; body?: HTMLElement; tabs?: HTMLElement; list?: HTMLElement | null; outline?: HTMLElement | null; setTab(tab: string): void; render?(): void; reactOutlineRows?(): Array<{ kind: string; slideId?: string; level: number; text: string }>; reactMasterCatalog?(): Array<{ kind: string; label: string; preview: HTMLElement; selected: boolean }> | null; reactSelectMaster?(kind: string): void; reactSelectThumb?(event: PointerEvent, index: number): boolean | undefined; onOutlineInput?(): void; onOutlinePaste?(event: ClipboardEvent): void; onOutlineKey?(event: KeyboardEvent): void; syncOutlineSlide?(): void; styleOutline?(): void; commitOutline?(keepFocus?: boolean): void; onThumbDown?(event: PointerEvent, index: number): void; onThumbContext?(event: MouseEvent, index: number): void; onThumbDoubleClick?(): void };
    sorter?: { el?: HTMLElement; reactMounted?: boolean; attachReactSurface?(root: HTMLElement): () => void; reactSelectThumb?(event: PointerEvent, index: number): boolean | undefined; onDown?(event: PointerEvent, index: number, item: HTMLElement): void; mark(): void };
    task?: { stack: string[]; pos: number; current?: string; registry: Record<string, { id: string; title: string; hidden?: boolean }>; show(id: string): void; go(position: number): void };
  };
  ui?: { openMenu(items: Array<{ label: string; run(): void; checked(): boolean }>, position: { left: number; bottom: number }): void };
  media?: { get(id: string): { type?: string } | null };
  mimeToExt?(mime: string): string;
  show?: { previewTransition(frame: Element | null, presentation: NonNullable<PptxPaneRuntime['pres']>, slide: SlideLike, previous?: SlideLike): void };
  anim?: { transName(transition: NonNullable<SlideLike['trans']>): string };
  bus?: { on(event: string, listener: () => void): () => void; emit(event: string): void };
  icons?: { get(name: string, size?: number): string };
  t?(message: string, ...args: unknown[]): string;
}

function runtime(): PptxPaneRuntime | undefined { return (window as Window & { L?: PptxPaneRuntime }).L; }

type FailedThumbnailImages = Map<HTMLImageElement, string>;

const reactSvgAttributes: Record<string, string> = {
  'clip-path': 'clipPath', 'fill-opacity': 'fillOpacity', 'fill-rule': 'fillRule', 'font-family': 'fontFamily', 'font-size': 'fontSize', 'font-style': 'fontStyle', 'font-weight': 'fontWeight',
  'gradienttransform': 'gradientTransform', 'gradientunits': 'gradientUnits', 'lengthadjust': 'lengthAdjust', 'marker-end': 'markerEnd', 'marker-mid': 'markerMid', 'marker-start': 'markerStart', 'markerheight': 'markerHeight', 'markerunits': 'markerUnits', 'markerwidth': 'markerWidth',
  'paint-order': 'paintOrder', 'pointer-events': 'pointerEvents', 'preserveaspectratio': 'preserveAspectRatio', 'patterncontentunits': 'patternContentUnits', 'patterntransform': 'patternTransform', 'patternunits': 'patternUnits', 'refx': 'refX', 'refy': 'refY', 'shape-rendering': 'shapeRendering',
  'stop-color': 'stopColor', 'stop-opacity': 'stopOpacity', 'stroke-dasharray': 'strokeDasharray', 'stroke-dashoffset': 'strokeDashoffset', 'stroke-linecap': 'strokeLinecap', 'stroke-linejoin': 'strokeLinejoin', 'stroke-miterlimit': 'strokeMiterlimit', 'stroke-opacity': 'strokeOpacity', 'stroke-width': 'strokeWidth',
  'text-anchor': 'textAnchor', 'text-decoration': 'textDecoration', 'textlength': 'textLength', 'viewbox': 'viewBox', 'writing-mode': 'writingMode', 'xlink:href': 'xlinkHref', 'xml:space': 'xmlSpace',
};

function reactAttributeName(name: string): string {
  if (name === 'class') return 'className';
  if (name === 'for') return 'htmlFor';
  if (name === 'tabindex') return 'tabIndex';
  if (name === 'colspan') return 'colSpan';
  if (name === 'rowspan') return 'rowSpan';
  if (name === 'contenteditable') return 'contentEditable';
  return reactSvgAttributes[name.toLowerCase()] || name;
}

function reactStyle(element: HTMLElement | SVGElement): CSSProperties {
  const output: Record<string, string> = {};
  for (let index = 0; index < element.style.length; index++) {
    const property = element.style.item(index);
    const name = property.startsWith('--') ? property : property.replace(/^-ms-/, 'ms-').replace(/-([a-z])/g, (_match, letter: string) => letter.toUpperCase());
    output[name] = element.style.getPropertyValue(property);
  }
  return output as CSSProperties;
}

function thumbnailNodeToReact(node: Node, key: string, failures: FailedThumbnailImages, onImageError: (image: HTMLImageElement) => void, root = false, owner = 'pptx-slide-artwork'): ReactNode {
  if (node.nodeType === Node.TEXT_NODE) return node.textContent;
  if (!(node instanceof Element)) return null;
  if (node instanceof HTMLImageElement && failures.has(node)) return null;
  const props: Record<string, unknown> = { key };
  for (const attribute of Array.from(node.attributes)) {
    if (/^on/i.test(attribute.name)) continue;
    if (attribute.name === 'style') {
      props.style = reactStyle(node as HTMLElement | SVGElement);
      continue;
    }
    const name = reactAttributeName(attribute.name);
    if (name === 'draggable' && node instanceof HTMLImageElement) props[name] = node.draggable;
    else if (name === 'hidden' && node instanceof HTMLElement) props[name] = node.hidden;
    else if (name === 'tabIndex' && node instanceof HTMLElement) props[name] = node.tabIndex;
    else props[name] = attribute.value;
  }
  if (root) props['data-react-owned'] = owner;
  if (node instanceof HTMLImageElement && node.dataset.mediaId) props.onError = () => onImageError(node);
  const failedPicture = node instanceof HTMLElement && node.classList.contains('pic')
    ? [...failures].find(([image]) => node.contains(image))
    : undefined;
  if (failedPicture) props.className = `${node.getAttribute('class') || ''} missing`.trim();
  const children = Array.from(node.childNodes, (child, index) => thumbnailNodeToReact(child, `${key}.${index}`, failures, onImageError));
  if (failedPicture) children.push(createElement('div', { className: 'pic-missing', key: `${key}.missing` }, failedPicture[1]));
  return createElement(node.localName, props, ...children);
}

export function PptxThumbnailArtwork({ element }: { element: HTMLElement }) {
  const [state, setState] = useState<{ element: HTMLElement; failures: FailedThumbnailImages }>(() => ({ element, failures: new Map() }));
  const failures = state.element === element ? state.failures : new Map<HTMLImageElement, string>();
  const onImageError = (image: HTMLImageElement) => {
    const mediaId = image.dataset.mediaId || '';
    const media = runtime()?.media?.get(mediaId);
    const ext = media?.type ? runtime()?.mimeToExt?.(media.type) || '' : '';
    const L = runtime();
    const label = ext && ext !== 'png' ? L?.t?.('Picture (.{0})', ext) || `Picture (.${ext})` : L?.t?.('Picture') || 'Picture';
    setState((current) => {
      const next = current.element === element ? new Map(current.failures) : new Map<HTMLImageElement, string>();
      next.set(image, label);
      return { element, failures: next };
    });
  };
  return thumbnailNodeToReact(element, 'artwork', failures, onImageError, true);
}

/** Reconcile engine-generated preview markup as a React element tree. */
export function PptxPreviewArtwork({ markup }: { markup: string }) {
  const element = useMemo(() => {
    const parsed = new DOMParser().parseFromString(markup, 'text/html');
    const root = parsed.body.firstElementChild;
    if (parsed.body.children.length !== 1 || !root) return null;
    const validSvg = root.namespaceURI === 'http://www.w3.org/2000/svg' && root.localName === 'svg';
    const validThumbnail = root instanceof HTMLElement && root.localName === 'div' && root.classList.contains('thumb-box');
    return validSvg || validThumbnail ? root : null;
  }, [markup]);
  if (!element) return null;
  return thumbnailNodeToReact(element, 'svg', new Map(), () => {}, true, 'pptx-svg-preview');
}

/** React-owned slide thumbnails and selection/reorder gestures; the document engine supplies artwork and mutation commands. */
export function PptxSlidesLayer({ host }: { host: HTMLElement }) {
  const [mount, setMount] = useState<HTMLElement | null>(null);
  const [tabsMount, setTabsMount] = useState<HTMLElement | null>(null);
  const [taskMount, setTaskMount] = useState<HTMLElement | null>(null);
  const [sorterMount, setSorterMount] = useState<HTMLElement | null>(null);
  const [revision, setRevision] = useState(0);
  const [sorterFocusIndex, setSorterFocusIndex] = useState(0);
  const outlineMutated = useRef(false);
  const keyboardSelection = useRef<((index: number, event: { shiftKey: boolean; ctrlKey: boolean; metaKey: boolean }, inSorter: boolean) => void) | null>(null);
  const slideDrag = useRef<{ pointerId: number; startX: number; startY: number; dragging: boolean; target: number | null; line: HTMLElement | null; body: HTMLElement; list: HTMLElement; indices: number[] } | null>(null);
  const sorterDrag = useRef<{ pointerId: number; startX: number; startY: number; dragging: boolean; target: number | null; marker: HTMLElement | null; root: HTMLElement; indices: number[] } | null>(null);
  useEffect(() => {
    let slot: HTMLElement | null = null;
    let tabsSlot: HTMLElement | null = null;
    let taskSlot: HTMLElement | null = null;
    let leftBody: HTMLElement | null = null;
    let sorterSlot: HTMLElement | null = null;
    let sorterSurfaceCleanup: (() => void) | undefined;
    let offRender: (() => void) | undefined;
    let offMark: (() => void) | undefined;
    let offTask: (() => void) | undefined;
    let offSorterRender: (() => void) | undefined;
    let offSorterMark: (() => void) | undefined;
    let subscribed = false;
    const onLeftFocus = () => { const app = runtime()?.app; if (app) app.focusArea = 'slides'; };
    const onLeftBlur = () => { const app = runtime()?.app; if (app) app.focusArea = null; };
    const clearSlideDrag = (commit: boolean) => {
      const drag = slideDrag.current;
      if (!drag) return;
      slideDrag.current = null;
      drag.line?.remove();
      if (commit && drag.dragging && drag.target != null) runtime()?.app?.moveSlides?.(drag.indices, drag.target);
    };
    const clearSorterDrag = (commit: boolean) => {
      const drag = sorterDrag.current;
      if (!drag) return;
      sorterDrag.current = null;
      drag.marker?.remove();
      if (commit && drag.dragging && drag.target != null) runtime()?.app?.moveSlides?.(drag.indices, drag.target);
    };
    const updateSlideDrag = (event: PointerEvent) => {
      const drag = slideDrag.current;
      if (!drag || drag.pointerId !== event.pointerId) return;
      if (!drag.dragging && Math.hypot(event.clientX - drag.startX, event.clientY - drag.startY) < 6) return;
      drag.dragging = true;
      if (!drag.line) {
        drag.line = document.createElement('div');
        drag.line.className = 'drop-line';
        drag.body.appendChild(drag.line);
      }
      const bounds = drag.body.getBoundingClientRect();
      if (event.clientX < bounds.left || event.clientX > bounds.right) {
        drag.line.style.display = 'none';
        drag.target = null;
        return;
      }
      drag.line.style.display = '';
      const items = [...drag.list.querySelectorAll<HTMLElement>('.th')];
      drag.target = items.length;
      for (let index = 0; index < items.length; index++) {
        const item = items[index];
        if (!item) continue;
        const rect = item.getBoundingClientRect();
        if (event.clientY < rect.top + rect.height / 2) { drag.target = index; break; }
      }
      const reference = items[drag.target] || items[items.length - 1];
      if (reference) {
        const rect = reference.getBoundingClientRect();
        drag.line.style.top = `${(drag.target < items.length ? rect.top - 5 : rect.bottom + 4) - bounds.top + drag.body.scrollTop}px`;
      }
      if (event.clientY < bounds.top + 20) drag.body.scrollTop -= 8;
      else if (event.clientY > bounds.bottom - 20) drag.body.scrollTop += 8;
    };
    const updateSorterDrag = (event: PointerEvent) => {
      const drag = sorterDrag.current;
      if (!drag || drag.pointerId !== event.pointerId) return;
      if (!drag.dragging && Math.hypot(event.clientX - drag.startX, event.clientY - drag.startY) < 6) return;
      drag.dragging = true;
      if (!drag.marker) {
        drag.marker = document.createElement('div');
        Object.assign(drag.marker.style, { position: 'absolute', width: '3px', background: '#000', pointerEvents: 'none', zIndex: '5' });
        drag.root.appendChild(drag.marker);
      }
      const items = [...drag.root.querySelectorAll<HTMLElement>('.so-it')];
      drag.target = items.length;
      let reference: HTMLElement | undefined;
      let before = true;
      for (let index = 0; index < items.length; index++) {
        const item = items[index];
        if (!item) continue;
        const rect = item.getBoundingClientRect();
        if (event.clientY >= rect.top - 12 && event.clientY <= rect.bottom + 12 && event.clientX < rect.left + rect.width / 2) { drag.target = index; reference = item; before = true; break; }
        if (event.clientY >= rect.top - 12 && event.clientY <= rect.bottom + 12) { reference = item; before = false; drag.target = index + 1; }
      }
      if (!reference) return;
      const bounds = drag.root.getBoundingClientRect();
      const rect = reference.getBoundingClientRect();
      drag.marker.style.left = `${(before ? rect.left - 12 : rect.right + 10) - bounds.left + drag.root.scrollLeft}px`;
      drag.marker.style.top = `${rect.top - bounds.top + drag.root.scrollTop}px`;
      drag.marker.style.height = `${rect.height}px`;
    };
    const onPointerMove = (event: PointerEvent) => { updateSlideDrag(event); updateSorterDrag(event); };
    const onPointerUp = (event: PointerEvent) => {
      if (slideDrag.current?.pointerId === event.pointerId) { updateSlideDrag(event); clearSlideDrag(true); }
      if (sorterDrag.current?.pointerId === event.pointerId) { updateSorterDrag(event); clearSorterDrag(true); }
    };
    const onPointerCancel = (event: PointerEvent) => {
      if (slideDrag.current?.pointerId === event.pointerId) clearSlideDrag(false);
      if (sorterDrag.current?.pointerId === event.pointerId) clearSorterDrag(false);
    };
    const onKeyDown = (event: KeyboardEvent) => { if (event.key === 'Escape') { clearSlideDrag(false); clearSorterDrag(false); } };
    const onThumbnailKeyDownCapture = (event: KeyboardEvent) => {
      if (event.key !== 'Enter' && event.key !== ' ') return;
      const ElementConstructor = host.ownerDocument.defaultView?.Element;
      if (!ElementConstructor || !(event.target instanceof ElementConstructor)) return;
      const button = event.target.closest<HTMLButtonElement>('.sorter .so-frame, #react-slides-root .th-frame');
      if (!button || !host.contains(button)) return;
      const item = button.closest<HTMLElement>('.so-it[data-i], .th[data-i]');
      const index = Number(item?.dataset.i);
      if (!Number.isInteger(index) || index < 0) return;
      event.preventDefault();
      event.stopImmediatePropagation();
      keyboardSelection.current?.(index, event, button.classList.contains('so-frame'));
    };
    const onOtherPointerDown = (event: PointerEvent) => {
      if (slideDrag.current && event.pointerId !== slideDrag.current.pointerId) clearSlideDrag(false);
      if (sorterDrag.current && event.pointerId !== sorterDrag.current.pointerId) clearSorterDrag(false);
    };
    window.addEventListener('pointermove', onPointerMove);
    window.addEventListener('pointerup', onPointerUp);
    window.addEventListener('pointercancel', onPointerCancel);
    window.addEventListener('keydown', onKeyDown, true);
    window.addEventListener('pointerdown', onOtherPointerDown);
    host.addEventListener('keydown', onThumbnailKeyDownCapture, true);
    const attach = () => {
      const target = host.querySelector<HTMLElement>('#react-slides-root');
      const tabsTarget = host.querySelector<HTMLElement>('#react-slides-tabs-root');
      const taskTarget = host.querySelector<HTMLElement>('#react-task-head-root');
      const sorterTarget = host.querySelector<HTMLElement>('.sorter');
      const leftBodyTarget = host.querySelector<HTMLElement>('#leftpane > .lp-body');
      const engine = runtime();
      if (!engine) return;
      if (leftBodyTarget && leftBodyTarget !== leftBody) {
        leftBody?.removeEventListener('focus', onLeftFocus, true);
        leftBody?.removeEventListener('blur', onLeftBlur, true);
        leftBody = leftBodyTarget;
        leftBody.addEventListener('focus', onLeftFocus, true);
        leftBody.addEventListener('blur', onLeftBlur, true);
      }
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
        engine.panes?.left?.render?.();
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
        sorterSurfaceCleanup?.();
        sorterSlot?.classList.remove('vibeoffice-react-owned');
        sorterSlot = sorterTarget;
        sorterTarget.classList.add('vibeoffice-react-owned');
        const sorter = engine.panes?.sorter;
        if (sorter) sorter.reactMounted = true;
        sorterSurfaceCleanup = sorter?.attachReactSurface?.(sorterTarget);
        setSorterMount(sorterTarget);
        engine.bus?.emit('pptx-sorter-render');
      }
    };
    const observer = new MutationObserver(attach);
    observer.observe(host, { childList: true, subtree: true });
    attach();
    return () => {
      observer.disconnect(); offRender?.(); offMark?.(); offTask?.(); offSorterRender?.(); offSorterMark?.();
      clearSlideDrag(false);
      clearSorterDrag(false);
      window.removeEventListener('pointermove', onPointerMove);
      window.removeEventListener('pointerup', onPointerUp);
      window.removeEventListener('pointercancel', onPointerCancel);
      window.removeEventListener('keydown', onKeyDown, true);
      window.removeEventListener('pointerdown', onOtherPointerDown);
      host.removeEventListener('keydown', onThumbnailKeyDownCapture, true);
      sorterSurfaceCleanup?.();
      sorterSurfaceCleanup = undefined;
      leftBody?.removeEventListener('focus', onLeftFocus, true);
      leftBody?.removeEventListener('blur', onLeftBlur, true);
      leftBody = null;
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
  useEffect(() => {
    if (L && !sorterMount?.contains(sorterMount.ownerDocument.activeElement)) setSorterFocusIndex(L.ed?.idx || 0);
  }, [sorterMount, L?.ed?.idx, slides?.length]);
  if (!L || !pane) return null;
  const selectFromKeyboard = (index: number, event: { shiftKey: boolean; ctrlKey: boolean; metaKey: boolean }, inSorter: boolean) => {
    if (!L.app || !L.ed) return;
    L.app.focusArea = 'slides';
    if (event.shiftKey) {
      const start = Math.min(L.ed.idx, index), end = Math.max(L.ed.idx, index);
      L.app.slideSel = new Set(Array.from({ length: end - start + 1 }, (_unused, offset) => start + offset));
    } else if (event.ctrlKey || event.metaKey) {
      const selected = new Set(L.app.slideSel);
      if (selected.has(index) && selected.size > 1) selected.delete(index); else selected.add(index);
      L.app.slideSel = selected;
    } else {
      L.app.slideSel = new Set([index]);
      if (inSorter) L.ed.idx = index;
      else if (L.ed.idx !== index || L.ed.view !== 'normal') L.ed.goto?.(index);
    }
    L.bus?.emit(inSorter ? 'pptx-sorter-mark' : 'pptx-slides-mark');
  };
  keyboardSelection.current = selectFromKeyboard;
  const selectPaneTab = (tab: 'outline' | 'slides') => {
    if (tab === 'outline' && pane.tab !== 'outline') outlineMutated.current = false;
    if (tab === 'slides') outlineMutated.current = false;
    pane.setTab(tab);
  };
  const paneTabKeyDown = (event: ReactKeyboardEvent<HTMLButtonElement>, current: 'outline' | 'slides') => {
    const destination = event.key === 'ArrowRight' || event.key === 'ArrowDown' || event.key === 'End' ? 'slides'
      : event.key === 'ArrowLeft' || event.key === 'ArrowUp' || event.key === 'Home' ? 'outline' : null;
    if (!destination) return;
    event.preventDefault(); event.stopPropagation();
    selectPaneTab(destination);
    tabsMount?.querySelector<HTMLButtonElement>(`[data-pane-tab="${destination}"]`)?.focus();
  };
  const taskId = task ? task.stack[task.pos] : undefined;
  const taskTitle = taskId ? task?.registry[taskId]?.title : undefined;
  const taskHeader = taskMount && task && createPortal(<div className="tp-head">
    <button className="tp-nav" type="button" aria-label={L.t?.('Back') || 'Back'} disabled={task.pos <= 0} onClick={() => task.go(task.pos - 1)}><PptxPreviewArtwork markup={L.icons?.get('back') || ''} /></button>
    <button className="tp-nav" type="button" aria-label={L.t?.('Forward') || 'Forward'} disabled={task.pos >= task.stack.length - 1} onClick={() => task.go(task.pos + 1)}><PptxPreviewArtwork markup={L.icons?.get('forward') || ''} /></button>
    <button className="tp-nav" type="button" aria-label={L.t?.('Home') || 'Home'} onClick={() => task.show('getting-started')}><PptxPreviewArtwork markup={L.icons?.get('home') || ''} /></button>
    <button className="tp-title" type="button" aria-haspopup="menu" onClick={(event) => {
      const rect = event.currentTarget.getBoundingClientRect();
      L.ui?.openMenu(Object.values(task.registry).filter((item) => !item.hidden).map((item) => ({ label: item.title, run: () => task.show(item.id), checked: () => task.current === item.id })), { left: rect.left, bottom: rect.bottom });
    }}><span>{taskTitle ? L.t?.(taskTitle) || taskTitle : ''}</span><span className="dd-arrow" /></button>
    <button className="tp-nav" type="button" aria-label={L.t?.('Close task pane') || 'Close task pane'} onClick={() => L.app?.toggleTask(false)}><PptxPreviewArtwork markup={L.icons?.get('close', 12) || ''} /></button>
  </div>, taskMount);
  const tabs = tabsMount && createPortal(<div className="lp-tabs" role="tablist" aria-label="Presentation panes">
    <button className={`lp-tab${pane.tab === 'outline' ? ' on' : ''}`} type="button" role="tab" aria-selected={pane.tab === 'outline'} tabIndex={pane.tab === 'outline' ? 0 : -1} data-pane-tab="outline" onKeyDown={(event) => paneTabKeyDown(event, 'outline')} onClick={() => selectPaneTab('outline')}>{L.t?.('Outline|pane') || 'Outline'}</button>
    <button className={`lp-tab${pane.tab === 'slides' ? ' on' : ''}`} type="button" role="tab" aria-selected={pane.tab === 'slides'} tabIndex={pane.tab === 'slides' ? 0 : -1} data-pane-tab="slides" onKeyDown={(event) => paneTabKeyDown(event, 'slides')} onClick={() => selectPaneTab('slides')}>{L.t?.('Slides') || 'Slides'}</button>
    <button className="lp-close" type="button" aria-label={L.t?.('Close pane') || 'Close pane'} onClick={() => L.app?.toggleLeft(false)}><PptxPreviewArtwork markup={L.icons?.get('close', 12) || ''} /></button>
  </div>, tabsMount);
  const renderSlide = L.render?.thumb;
  const sorterView = sorterMount && sorter && presentation && slides && renderSlide && createPortal(<div className="sorter-grid" role="listbox" aria-label={L.t?.('Slide Sorter') || 'Slide Sorter'}>
    {slides.map((slide, index) => {
      const width = Math.max(90, Math.round(sorterMount.clientWidth ? Math.min(260, (sorterMount.clientWidth - 60) / 4) : 180));
      const artwork = renderSlide(presentation, slide, width, { index });
      return <div className={`so-it${L.app?.slideSel?.has(index) ? ' sel' : ''}${slide.hidden ? ' hiddenslide' : ''}`} data-i={index} key={slide.id || index} role="option" aria-selected={L.app?.slideSel?.has(index) || false} onPointerDown={(event: ReactPointerEvent<HTMLDivElement>) => {
        const handled = sorter.reactSelectThumb?.(event.nativeEvent, index);
        if (!handled || !sorterMount || !L.app) return;
        sorterDrag.current = { pointerId: event.pointerId, startX: event.clientX, startY: event.clientY, dragging: false, target: null, marker: null, root: sorterMount, indices: [...L.app.slideSel].sort((a, b) => a - b) };
      }} onDoubleClick={() => { if (L.ed) L.ed.idx = index; L.app?.setView('normal'); }} onContextMenu={(event: ReactMouseEvent<HTMLDivElement>) => { event.preventDefault(); if (!L.app?.slideSel.has(index)) { if (L.app) L.app.slideSel = new Set([index]); if (L.ed) L.ed.idx = index; sorter.mark(); } L.app?.contextThumb(event.nativeEvent); }}>
        <button className="so-frame" type="button" tabIndex={sorterFocusIndex === index ? 0 : -1} aria-label={L.t?.('Slide {0}', index + 1) || `Slide ${index + 1}`} onKeyDown={(event) => {
          if (['ArrowRight', 'ArrowDown', 'ArrowLeft', 'ArrowUp', 'Home', 'End'].includes(event.key)) {
            const columns = Math.max(1, getComputedStyle(event.currentTarget.parentElement?.parentElement || event.currentTarget).gridTemplateColumns.split(' ').length);
            const delta = event.key === 'ArrowRight' ? 1 : event.key === 'ArrowLeft' ? -1 : event.key === 'ArrowDown' ? columns : event.key === 'ArrowUp' ? -columns : 0;
            const destination = event.key === 'Home' ? 0 : event.key === 'End' ? slides.length - 1 : Math.max(0, Math.min(slides.length - 1, index + delta));
            event.preventDefault(); event.stopPropagation(); setSorterFocusIndex(destination);
            sorterMount?.querySelectorAll<HTMLButtonElement>('.so-frame')[destination]?.focus();
          } else if (event.key === 'Enter' || event.key === ' ') { event.preventDefault(); event.stopPropagation(); selectFromKeyboard(index, event, true); }
        }} onClick={(event) => { if (event.detail === 0) selectFromKeyboard(index, event, true); }}><PptxThumbnailArtwork element={artwork} /></button>
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
  const masterChoices = L.ed?.view === 'master' ? pane.reactMasterCatalog?.() : null;
  const masterView = mount && masterChoices && createPortal(<div className="thumbs" role="listbox" aria-label={L.t?.('Slide Masters') || 'Slide Masters'}>
    {masterChoices.map((choice, index) => <div className={`th${choice.selected ? ' cur' : ''}`} key={choice.kind} data-master-kind={choice.kind} role="option" aria-selected={choice.selected}>
      <div className="th-meta"><span className="th-num">{String(index + 1)}</span></div>
      <button className="th-frame" type="button" aria-label={choice.label} data-tip={choice.label} onClick={() => pane.reactSelectMaster?.(choice.kind)}><PptxThumbnailArtwork element={choice.preview} /></button>
    </div>)}
  </div>, mount);
  if (mount && L.ed?.view === 'master') return <>{taskHeader}{tabs}{masterView}{sorterView}</>;
  if (pane.tab === 'outline' && L.ed?.view !== 'master') return <>{taskHeader}{tabs}{outlineView}{sorterView}</>;
  if (!mount || !presentation || !slides || pane.tab !== 'slides' || L.ed?.view === 'master') return <>{taskHeader}{tabs}{sorterView}</>;
  const body = pane.body;
  const width = Math.max(60, (body?.clientWidth || 190) - 48);
  if (!renderSlide) return <>{taskHeader}{tabs}{sorterView}</>;
  return <>{taskHeader}{tabs}{sorterView}{createPortal(<div className="thumbs" role="listbox" aria-label="Slides" ref={(element) => { pane.list = element; }}>
    {slides.map((slide, index) => {
      const artwork = renderSlide(presentation, slide, width, { index });
      const selected = L.app?.slideSel?.has(index) || false;
      const current = L.ed?.idx === index;
      const hasAnimation = (slide.trans?.type && slide.trans.type !== 'none') || !!slide.anims?.length;
      return <div className={`th${current ? ' cur' : ''}${selected ? ' sel' : ''}${slide.hidden ? ' hiddenslide' : ''}`} data-i={index} key={slide.id || index} role="option" aria-selected={current} onPointerDown={(event: ReactPointerEvent<HTMLDivElement>) => {
        const handled = pane.reactSelectThumb?.(event.nativeEvent, index);
        if (!handled || !pane.body || !pane.list || !L.app) return;
        slideDrag.current = { pointerId: event.pointerId, startX: event.clientX, startY: event.clientY, dragging: false, target: null, line: null, body: pane.body, list: pane.list, indices: [...L.app.slideSel].sort((a, b) => a - b) };
      }} onContextMenu={(event: ReactMouseEvent<HTMLDivElement>) => pane.onThumbContext?.(event.nativeEvent, index)} onDoubleClick={() => pane.onThumbDoubleClick?.()}>
        <div className="th-meta"><span className="th-num">{String(index + (presentation.firstNum || 1))}</span>{hasAnimation && <button type="button" className="th-star" data-tip="Play animations" aria-label="Play animations" onClick={(event) => { event.stopPropagation(); const frame = event.currentTarget.closest('.th')?.querySelector('.th-frame') || null; L.show?.previewTransition(frame, presentation, slide, slides[index - 1]); }}><svg width="12" height="12" viewBox="0 0 12 12" aria-hidden="true"><path d="M6 .8l1.5 3.2 3.4.4-2.5 2.3.7 3.4L6 8.4 2.9 10.1l.7-3.4L1.1 4.4l3.4-.4z" fill="none" stroke="#333" /></svg></button>}</div>
        <button className="th-frame" type="button" aria-label={L.t?.('Slide {0}', index + 1) || `Slide ${index + 1}`} onKeyDown={(event) => { if (event.key === 'Enter' || event.key === ' ') { event.preventDefault(); event.stopPropagation(); selectFromKeyboard(index, event, false); } }} onClick={(event) => { if (event.detail === 0) selectFromKeyboard(index, event, false); }}><PptxThumbnailArtwork element={artwork} /></button>
      </div>;
    })}
  </div>, mount)}</>;
}

function formatTransitionDelay(milliseconds: number): string {
  const seconds = Math.round(milliseconds / 1000);
  return `${String(Math.floor(seconds / 60)).padStart(2, '0')}:${String(seconds % 60).padStart(2, '0')}`;
}
