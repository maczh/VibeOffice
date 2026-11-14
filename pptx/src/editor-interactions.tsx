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
  guidePointerDown(event: PointerEvent): void;
}
interface PptxAppEvents { keydown(event: KeyboardEvent): void; copy(event: ClipboardEvent): void; cut(event: ClipboardEvent): void; paste(event: ClipboardEvent): void }
interface PptxTextEditEvents { beforeinput(event: InputEvent): void; input(event: Event): void; paste(event: ClipboardEvent): void; focusin(event: FocusEvent): void }
interface PptxTextEditor { reactEvents?: PptxTextEditEvents; setReactEventOwnership?(owned: boolean): void }
interface InteractionRuntime { ed?: { reactEvents?: PptxEditorEvents }; te?: PptxTextEditor; app?: { reactAppEvents?: PptxAppEvents; reactDropFiles?(files: File[]): void }; t?(message: string): string }

/** Owns registration and cleanup of the PPTX slide-canvas interaction adapter. */
export function PptxEditorInteractions({ host }: { host: HTMLElement }) {
  const [dropDepth, setDropDepth] = useState(0);
  useEffect(() => {
    const runtime = (window as Window & { L?: InteractionRuntime }).L;
    const bound = new Map<HTMLElement, PptxEditorEvents>();
    let appEvents: PptxAppEvents | undefined;
    let textEditEvents: PptxTextEditEvents | undefined;
    let textEditor: PptxTextEditor | undefined;
    let dropTarget: HTMLElement | null = null;
    let guideTarget: HTMLElement | null = null;
    let guideEvents: PptxEditorEvents | undefined;
    let depth = 0;
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
        scroller.addEventListener('pointerdown', events.pointerdown);
        scroller.addEventListener('pointermove', events.pointermove);
        scroller.addEventListener('dblclick', events.dblclick);
        scroller.addEventListener('contextmenu', events.contextmenu);
        scroller.addEventListener('wheel', events.wheel, { passive: false });
        scroller.dataset.reactInteractions = 'pptx';
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
          if (guideEvents) guideTarget.removeEventListener('pointerdown', guideEvents.guidePointerDown as EventListener);
          delete guideTarget.dataset.reactGuideInteractions;
        }
        guideTarget = nextGuideTarget;
        guideTarget.addEventListener('pointerdown', events.guidePointerDown as EventListener);
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
    attach();
    return () => {
      observer.disconnect();
      for (const [element, events] of bound) {
        if (element.classList.contains('ed-scroll')) {
          element.removeEventListener('pointerdown', events.pointerdown);
          element.removeEventListener('pointermove', events.pointermove);
          element.removeEventListener('dblclick', events.dblclick);
          element.removeEventListener('contextmenu', events.contextmenu);
          element.removeEventListener('wheel', events.wheel);
        } else element.removeEventListener('click', events.emptyClick);
        delete element.dataset.reactInteractions;
      }
      bound.clear();
      if (guideTarget && guideEvents) {
        guideTarget.removeEventListener('pointerdown', guideEvents.guidePointerDown as EventListener);
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
      dropTarget = null;
      depth = 0;
    };
  }, [host]);
  const runtime = (window as Window & { L?: InteractionRuntime }).L;
  return <>{dropDepth ? createPortal(<div className="drop-hint" data-react-owned="pptx-drop-hint">{runtime?.t?.('Drop a .pptx to open it, or pictures to insert them') || 'Drop a .pptx to open it, or pictures to insert them'}</div>, host) : null}<PptxStatusBar host={host} /></>;
}
