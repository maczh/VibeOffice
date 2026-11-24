import { useEffect } from 'react';
import { useExcelxRuntime } from './editor-host';

interface ExcelxRuntime {
  grid?: { reactMountInteractions?(): () => void; reactInteractionMounted?: boolean; reactPointerLeave?(): void };
  app?: { reactDrawTrace?(): void; reactAppKeydown?(event: KeyboardEvent): void; reactBeforeUnload?(event: BeforeUnloadEvent): void };
  bus?: { on(event: string, listener: () => void): () => void };
}

/** Own the canvas grid's interaction listeners for the lifetime of the React editor mount. */
export function ExcelxGridInteractionLayer({ host }: { host: HTMLElement }) {
  const runtime = useExcelxRuntime<ExcelxRuntime>();
  const ownerDocument = host.ownerDocument;
  const ownerWindow = ownerDocument.defaultView;
  useEffect(() => {
    const offTrace = runtime?.bus?.on('grid-paint', () => runtime?.app?.reactDrawTrace?.());
    const pointerLeave = () => runtime?.grid?.reactPointerLeave?.();
    const keydown = (event: KeyboardEvent) => runtime?.app?.reactAppKeydown?.(event);
    const beforeUnload = (event: BeforeUnloadEvent) => runtime?.app?.reactBeforeUnload?.(event);
    ownerDocument.addEventListener('pointerleave', pointerLeave);
    host.addEventListener('keydown', keydown);
    ownerWindow?.addEventListener('beforeunload', beforeUnload);
    let target: HTMLElement | null = null;
    let dispose: (() => void) | null = null;
    const attach = () => {
      const next = host.querySelector<HTMLElement>('#gridwrap');
      if (next === target && (dispose || !next)) return;
      dispose?.();
      target = next;
      dispose = target ? runtime?.grid?.reactMountInteractions?.() || null : null;
    };
    const Observer = ownerWindow?.MutationObserver;
    const observer = Observer ? new Observer(attach) : null;
    observer?.observe(host, { childList: true, subtree: true });
    const off = runtime?.bus?.on('app-ready', attach);
    attach();
    return () => {
      observer?.disconnect();
      off?.();
      offTrace?.();
      ownerDocument.removeEventListener('pointerleave', pointerLeave);
      host.removeEventListener('keydown', keydown);
      ownerWindow?.removeEventListener('beforeunload', beforeUnload);
      dispose?.();
      dispose = null;
      target = null;
    };
  }, [host, ownerDocument, ownerWindow, runtime]);
  return null;
}
