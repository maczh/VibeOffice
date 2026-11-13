import { useEffect } from 'react';

interface ExcelxRuntime {
  grid?: { reactMountInteractions?(): () => void; reactInteractionMounted?: boolean };
  bus?: { on(event: string, listener: () => void): () => void };
}

function runtime(): ExcelxRuntime | undefined {
  return (window as Window & { L?: ExcelxRuntime }).L;
}

/** Own the canvas grid's interaction listeners for the lifetime of the React editor mount. */
export function ExcelxGridInteractionLayer({ host }: { host: HTMLElement }) {
  useEffect(() => {
    let target: HTMLElement | null = null;
    let dispose: (() => void) | null = null;
    const attach = () => {
      const next = host.querySelector<HTMLElement>('#gridwrap');
      if (next === target && (dispose || !next)) return;
      dispose?.();
      target = next;
      dispose = target ? runtime()?.grid?.reactMountInteractions?.() || null : null;
    };
    const observer = new MutationObserver(attach);
    observer.observe(host, { childList: true, subtree: true });
    const off = runtime()?.bus?.on('app-ready', attach);
    attach();
    return () => {
      observer.disconnect();
      off?.();
      dispose?.();
      dispose = null;
      target = null;
    };
  }, [host]);
  return null;
}
