import { useEffect, useState } from 'react';
import { createPortal } from 'react-dom';
import { PptxStatusBar } from './status-bar.js';

interface PptxEditorEvents {
  pointerdown(event: PointerEvent): void;
  pointermove(event: PointerEvent): void;
  dblclick(event: MouseEvent): void;
  contextmenu(event: MouseEvent): void;
  wheel(event: WheelEvent): void;
  emptyClick(): void;
  guidePointerDown(event: PointerEvent): PptxGuideGesture | null;
}
interface PptxGuideGesture { move(event: PointerEvent): void; finish(cancel: boolean): void }
interface PptxCanvas { reactEvents?: PptxEditorEvents; reactLayout?(): void }
interface PptxAppEvents { keydown(event: KeyboardEvent): void; copy(event: ClipboardEvent): void; cut(event: ClipboardEvent): void; paste(event: ClipboardEvent): void }
interface PptxSplitterEvents { left(event: PointerEvent): void; notes(event: PointerEvent): void }
interface PptxTextEditEvents { beforeinput(event: InputEvent): void; input(event: Event): void; paste(event: ClipboardEvent): void; focusin(event: FocusEvent): void }
interface PptxTextEditor { reactEvents?: PptxTextEditEvents; reactSelectionChange?(): void; setReactEventOwnership?(owned: boolean): void }
interface InteractionRuntime { ed?: PptxCanvas; te?: PptxTextEditor; app?: { reactAppEvents?: PptxAppEvents; reactSplitterEvents?: PptxSplitterEvents; reactTransparentPointerDown?(event: PointerEvent): boolean; reactTransparentCancel?(): void; reactBeforeUnload?(event: BeforeUnloadEvent): void; reactResize?(): void; reactDropFiles?(files: File[]): void }; bus?: { on(event: string, listener: () => void): () => void }; t?(message: string): string }

/** Owns registration and cleanup of the PPTX slide-canvas interaction adapter. */
export function PptxEditorInteractions({ host }: { host: HTMLElement }) {
  const [dropDepth, setDropDepth] = useState(0);
  useEffect(() => {
    const runtime = (window as Window & { L?: InteractionRuntime }).L;
    const textSelectionChange = () => runtime?.te?.reactSelectionChange?.();
    const beforeUnload = (event: BeforeUnloadEvent) => runtime?.app?.reactBeforeUnload?.(event);
    const resize = () => runtime?.app?.reactResize?.();
    const layoutResize = () => runtime?.ed?.reactLayout?.();
    const transparentPointerDown = (event: PointerEvent) => { runtime?.app?.reactTransparentPointerDown?.(event); };
    document.addEventListener('selectionchange', textSelectionChange);
    window.addEventListener('beforeunload', beforeUnload);
    window.addEventListener('resize', resize);
    const canvasResizeObserver = typeof ResizeObserver === 'undefined' ? null : new ResizeObserver(layoutResize);
    if (!canvasResizeObserver) window.addEventListener('resize', layoutResize);
    const bound = new Map<HTMLElement, PptxEditorEvents>();
    const splitterBindings = new Map<HTMLElement, (event: PointerEvent) => void>();
    let appEvents: PptxAppEvents | undefined;
    let textEditEvents: PptxTextEditEvents | undefined;
    let textEditor: PptxTextEditor | undefined;
    let dropTarget: HTMLElement | null = null;
    let guideTarget: HTMLElement | null = null;
    let guideEvents: PptxEditorEvents | undefined;
    let activeGuide: { pointerId: number; target: HTMLElement; gesture: PptxGuideGesture } | null = null;
    let depth = 0;
    const finishGuide = (cancel: boolean) => {
      const active = activeGuide;
      if (!active) return;
      activeGuide = null;
      active.target.removeEventListener('pointermove', guideMove);
      active.target.removeEventListener('pointerup', guideUp);
      active.target.removeEventListener('pointercancel', guideCancel);
      active.target.removeEventListener('lostpointercapture', guideLostCapture);
      active.gesture.finish(cancel);
      if (active.target.hasPointerCapture(active.pointerId)) active.target.releasePointerCapture(active.pointerId);
    };
    const guideMove = (event: Event) => { const pointer = event as PointerEvent; if (activeGuide?.pointerId === pointer.pointerId) activeGuide.gesture.move(pointer); };
    const guideUp = (event: Event) => { if (activeGuide?.pointerId === (event as PointerEvent).pointerId) finishGuide(false); };
    const guideCancel = (event: Event) => { if (activeGuide?.pointerId === (event as PointerEvent).pointerId) finishGuide(true); };
    const guideLostCapture = () => finishGuide(true);
    const guideDown = (event: Event) => {
      const pointer = event as PointerEvent;
      if (activeGuide || !guideTarget || !guideEvents) return;
      const gesture = guideEvents.guidePointerDown(pointer);
      if (!gesture) return;
      activeGuide = { pointerId: pointer.pointerId, target: guideTarget, gesture };
      guideTarget.addEventListener('pointermove', guideMove);
      guideTarget.addEventListener('pointerup', guideUp);
      guideTarget.addEventListener('pointercancel', guideCancel);
      guideTarget.addEventListener('lostpointercapture', guideLostCapture);
      try { guideTarget.setPointerCapture(pointer.pointerId); } catch { finishGuide(true); }
    };
    const hasFiles = (event: DragEvent) => !!event.dataTransfer && Array.from(event.dataTransfer.types || []).includes('Files');
    const dragEnter = (event: DragEvent) => { if (!hasFiles(event)) return; depth++; setDropDepth(depth); };
    const dragLeave = () => { depth = Math.max(0, depth - 1); setDropDepth(depth); };
    const dragOver = (event: DragEvent) => { if (hasFiles(event)) event.preventDefault(); };
    const drop = (event: DragEvent) => {
      depth = 0; setDropDepth(0);
      const files = Array.from(event.dataTransfer?.files || []);
      if (!files.length) return;
      event.preventDefault();
      runtime?.app?.reactDropFiles?.(files);
    };
    const attach = () => {
      const events = runtime?.ed?.reactEvents;
      if (!events) return;
      const scroller = host.querySelector<HTMLElement>('.ed-scroll');
      if (scroller && !bound.has(scroller)) {
        canvasResizeObserver?.observe(scroller);
        scroller.addEventListener('pointerdown', events.pointerdown);
        scroller.addEventListener('pointermove', events.pointermove);
        scroller.addEventListener('dblclick', events.dblclick);
        scroller.addEventListener('contextmenu', events.contextmenu);
        scroller.addEventListener('wheel', events.wheel, { passive: false });
        scroller.addEventListener('pointerdown', transparentPointerDown, true);
        scroller.dataset.reactInteractions = 'pptx';
        scroller.dataset.reactTransparentPointer = 'pptx';
        bound.set(scroller, events);
      }
      const empty = host.querySelector<HTMLElement>('.ed-empty');
      if (empty && !bound.has(empty)) {
        empty.addEventListener('click', events.emptyClick);
        empty.dataset.reactInteractions = 'pptx-empty';
        bound.set(empty, events);
      }
      const nextGuideTarget = host.querySelector<HTMLElement>('.ed-grid');
      if (nextGuideTarget && (nextGuideTarget !== guideTarget || events !== guideEvents)) {
        if (guideTarget) {
          finishGuide(true);
          guideTarget.removeEventListener('pointerdown', guideDown);
          delete guideTarget.dataset.reactGuideInteractions;
        }
        guideTarget = nextGuideTarget;
        guideTarget.addEventListener('pointerdown', guideDown);
        guideTarget.dataset.reactGuideInteractions = 'pptx';
        guideEvents = events;
      }
      const nextAppEvents = runtime?.app?.reactAppEvents;
      if (nextAppEvents && nextAppEvents !== appEvents) {
        if (appEvents) {
          host.removeEventListener('keydown', appEvents.keydown);
          host.removeEventListener('copy', appEvents.copy);
          host.removeEventListener('cut', appEvents.cut);
          host.removeEventListener('paste', appEvents.paste);
        }
        appEvents = nextAppEvents;
        host.addEventListener('keydown', appEvents.keydown);
        host.addEventListener('copy', appEvents.copy);
        host.addEventListener('cut', appEvents.cut);
        host.addEventListener('paste', appEvents.paste);
        host.dataset.reactAppEvents = 'pptx';
      }
      const splitterEvents = runtime?.app?.reactSplitterEvents;
      for (const [selector, handler] of [['#split-left', splitterEvents?.left], ['#split-notes', splitterEvents?.notes]] as const) {
        const target = host.querySelector<HTMLElement>(selector);
        if (!target || !handler || splitterBindings.get(target) === handler) continue;
        const previous = [...splitterBindings].find(([element]) => element.matches(selector));
        if (previous) {
          previous[0].removeEventListener('pointerdown', previous[1]);
          delete previous[0].dataset.reactSplitter;
          splitterBindings.delete(previous[0]);
        }
        target.addEventListener('pointerdown', handler);
        target.dataset.reactSplitter = selector === '#split-left' ? 'left' : 'notes';
        splitterBindings.set(target, handler);
      }
      const nextTextEditor = runtime?.te;
      const nextTextEvents = nextTextEditor?.reactEvents;
      if (nextTextEvents && nextTextEvents !== textEditEvents) {
        if (textEditEvents) {
          host.removeEventListener('beforeinput', textEditEvents.beforeinput);
          host.removeEventListener('input', textEditEvents.input);
          host.removeEventListener('paste', textEditEvents.paste);
          host.removeEventListener('focusin', textEditEvents.focusin, true);
          textEditor?.setReactEventOwnership?.(false);
        }
        textEditEvents = nextTextEvents;
        textEditor = nextTextEditor;
        host.addEventListener('beforeinput', textEditEvents.beforeinput);
        host.addEventListener('input', textEditEvents.input);
        host.addEventListener('paste', textEditEvents.paste);
        host.addEventListener('focusin', textEditEvents.focusin, true);
        textEditor?.setReactEventOwnership?.(true);
        host.dataset.reactTextEditEvents = 'pptx';
      }
      const nextDropTarget = host.querySelector<HTMLElement>('#app');
      if (nextDropTarget && nextDropTarget !== dropTarget) {
        if (dropTarget) {
          dropTarget.removeEventListener('dragenter', dragEnter);
          dropTarget.removeEventListener('dragleave', dragLeave);
          dropTarget.removeEventListener('dragover', dragOver);
          dropTarget.removeEventListener('drop', drop);
        }
        dropTarget = nextDropTarget;
        dropTarget.addEventListener('dragenter', dragEnter);
        dropTarget.addEventListener('dragleave', dragLeave);
        dropTarget.addEventListener('dragover', dragOver);
        dropTarget.addEventListener('drop', drop);
      }
    };
    const observer = new MutationObserver(attach);
    observer.observe(host, { childList: true, subtree: true });
    const splittersReady = () => attach();
    host.addEventListener('vibeoffice-pptx-splitters-ready', splittersReady);
    attach();
    return () => {
      observer.disconnect();
      host.removeEventListener('vibeoffice-pptx-splitters-ready', splittersReady);
      document.removeEventListener('selectionchange', textSelectionChange);
      window.removeEventListener('beforeunload', beforeUnload);
      window.removeEventListener('resize', resize);
      canvasResizeObserver?.disconnect();
      if (!canvasResizeObserver) window.removeEventListener('resize', layoutResize);
      for (const [element, events] of bound) {
        if (element.classList.contains('ed-scroll')) {
          element.removeEventListener('pointerdown', events.pointerdown);
          element.removeEventListener('pointermove', events.pointermove);
          element.removeEventListener('dblclick', events.dblclick);
          element.removeEventListener('contextmenu', events.contextmenu);
          element.removeEventListener('wheel', events.wheel);
          element.removeEventListener('pointerdown', transparentPointerDown, true);
          delete element.dataset.reactTransparentPointer;
        } else element.removeEventListener('click', events.emptyClick);
        delete element.dataset.reactInteractions;
      }
      bound.clear();
      for (const [element, handler] of splitterBindings) {
        element.removeEventListener('pointerdown', handler);
        delete element.dataset.reactSplitter;
      }
      splitterBindings.clear();
      if (guideTarget && guideEvents) {
        finishGuide(true);
        guideTarget.removeEventListener('pointerdown', guideDown);
        delete guideTarget.dataset.reactGuideInteractions;
      }
      if (appEvents) {
        host.removeEventListener('keydown', appEvents.keydown);
        host.removeEventListener('copy', appEvents.copy);
        host.removeEventListener('cut', appEvents.cut);
        host.removeEventListener('paste', appEvents.paste);
        delete host.dataset.reactAppEvents;
      }
      if (textEditEvents) {
        host.removeEventListener('beforeinput', textEditEvents.beforeinput);
        host.removeEventListener('input', textEditEvents.input);
        host.removeEventListener('paste', textEditEvents.paste);
        host.removeEventListener('focusin', textEditEvents.focusin, true);
        textEditor?.setReactEventOwnership?.(false);
        delete host.dataset.reactTextEditEvents;
      }
      if (dropTarget) {
        dropTarget.removeEventListener('dragenter', dragEnter);
        dropTarget.removeEventListener('dragleave', dragLeave);
        dropTarget.removeEventListener('dragover', dragOver);
        dropTarget.removeEventListener('drop', drop);
      }
      runtime?.app?.reactTransparentCancel?.();
      dropTarget = null;
      depth = 0;
    };
  }, [host]);
  const runtime = (window as Window & { L?: InteractionRuntime }).L;
  return <>{dropDepth ? createPortal(<div className="drop-hint" data-react-owned="pptx-drop-hint">{runtime?.t?.('Drop a .pptx to open it, or pictures to insert them') || 'Drop a .pptx to open it, or pictures to insert them'}</div>, host) : null}<PptxStatusBar host={host} /></>;
}
