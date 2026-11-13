import { useEffect, useState } from 'react';
import { createPortal } from 'react-dom';

interface ExcelxDropApi {
  openFiles?(files: File[]): void;
  insertPictureFile?(file: File): void;
  reactDropMounted?: boolean;
}

function runtime(): { app?: ExcelxDropApi; t?(message: string): string } | undefined {
  return (window as Window & { L?: { app?: ExcelxDropApi; t?(message: string): string } }).L;
}

/** React owns the file-drop hint and event lifetime; the engine retains file handling. */
export function ExcelxFileDropLayer({ host }: { host: HTMLElement }) {
  const [mount, setMount] = useState<HTMLElement | null>(null);
  const [active, setActive] = useState(false);
  useEffect(() => {
    let app: HTMLElement | null = null;
    let depth = 0;
    const fileDrag = (event: DragEvent) => !!event.dataTransfer && Array.from(event.dataTransfer.types || []).includes('Files');
    const enter = (event: DragEvent) => {
      if (!fileDrag(event)) return;
      depth++;
      setActive(true);
    };
    const leave = (event: DragEvent) => {
      if (!fileDrag(event)) return;
      depth = Math.max(0, depth - 1);
      if (!depth) setActive(false);
    };
    const over = (event: DragEvent) => { if (fileDrag(event)) event.preventDefault(); };
    const drop = (event: DragEvent) => {
      if (!fileDrag(event)) return;
      depth = 0;
      setActive(false);
      const files = Array.from(event.dataTransfer?.files || []);
      if (!files.length) return;
      event.preventDefault();
      const books = files.filter((file) => /\.(xlsx|xlsm|xltx|xltm|csv|txt|tsv|prn|xml|xls)$/i.test(file.name));
      if (books.length) runtime()?.app?.openFiles?.(books);
      for (const file of files.filter((item) => /^image\//.test(item.type))) runtime()?.app?.insertPictureFile?.(file);
    };
    const detach = () => {
      if (!app) return;
      app.removeEventListener('dragenter', enter);
      app.removeEventListener('dragleave', leave);
      app.removeEventListener('dragover', over);
      app.removeEventListener('drop', drop);
      if (runtime()?.app) runtime()!.app!.reactDropMounted = false;
      setMount((current) => current === app ? null : current);
      app = null;
      depth = 0;
      setActive(false);
    };
    const attach = () => {
      const next = host.querySelector<HTMLElement>('#app');
      if (next === app) return;
      detach();
      if (!next) return;
      app = next;
      setMount(next);
      if (runtime()?.app) runtime()!.app!.reactDropMounted = true;
      app.addEventListener('dragenter', enter);
      app.addEventListener('dragleave', leave);
      app.addEventListener('dragover', over);
      app.addEventListener('drop', drop);
    };
    const observer = new MutationObserver(attach);
    observer.observe(host, { childList: true, subtree: true });
    attach();
    return () => {
      observer.disconnect();
      detach();
    };
  }, [host]);
  return mount && active ? createPortal(<div className="drop-hint" data-react-owned="excelx-drop-hint">{runtime()?.t?.('Drop a workbook to open it, or pictures to insert them') || 'Drop a workbook to open it, or pictures to insert them'}</div>, mount) : null;
}
