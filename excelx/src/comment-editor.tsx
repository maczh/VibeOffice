import { useEffect, useRef, useState } from 'react';
import { useExcelxHost, useExcelxRuntime } from './editor-host.js';

type Catalog = { initialText: string; left: number; top: number; width: number; height: number };
type Runtime = { dlg?: { reactCommentCatalog?(): Catalog | null; reactCommentFinish?(save: boolean, value?: string, width?: number, height?: number): boolean }; bus?: { on(event: string, listener: () => void): () => void } };

/** React-owned in-grid comment editor preserving the legacy floating-note interaction. */
export function ExcelxCommentEditor() {
  const runtime = useExcelxRuntime<Runtime>();
  const host = useExcelxHost();
  const [catalog, setCatalog] = useState<Catalog | null>(null);
  const [value, setValue] = useState('');
  const area = useRef<HTMLTextAreaElement>(null);
  const valueRef = useRef(value);
  valueRef.current = value;
  const finishRef = useRef<(save: boolean) => void>(() => {});
  useEffect(() => {
    const ownerWindow = host.ownerDocument.defaultView;
    let focusFrame: number | null = null;
    const refresh = () => {
      const next = runtime?.dlg?.reactCommentCatalog?.() || null;
      setCatalog(next);
      if (next) { setValue(next.initialText); if (focusFrame !== null) ownerWindow?.cancelAnimationFrame(focusFrame); if (ownerWindow) focusFrame = ownerWindow.requestAnimationFrame(() => { focusFrame = null; if (area.current?.isConnected) area.current.focus(); }); else area.current?.focus(); }
    };
    const off = runtime?.bus?.on('excelx-comment-editor', refresh);
    refresh();
    return () => { off?.(); if (focusFrame !== null) ownerWindow?.cancelAnimationFrame(focusFrame); };
  }, [host, runtime]);
  useEffect(() => {
    if (!catalog) return;
    const outside = (event: PointerEvent) => { if (area.current && !area.current.parentElement?.contains(event.target as Node)) finishRef.current(true); };
    const ownerDocument = host.ownerDocument;
    ownerDocument.addEventListener('pointerdown', outside, true);
    return () => ownerDocument.removeEventListener('pointerdown', outside, true);
  }, [!!catalog, host]);
  if (!catalog) return null;
  const finish = (save: boolean) => runtime?.dlg?.reactCommentFinish?.(save, valueRef.current, area.current?.parentElement?.offsetWidth || catalog.width, area.current?.parentElement?.offsetHeight || catalog.height);
  finishRef.current = finish;
  return <div className="cmt-edit" style={{ position: 'fixed', left: catalog.left, top: catalog.top, width: catalog.width, height: catalog.height }}>
    <textarea ref={area} aria-label="Comment" spellCheck onChange={(event) => { const text = event.currentTarget.value; valueRef.current = text; setValue(text); }} value={value} onKeyDown={(event) => { event.stopPropagation(); if (event.key === 'Escape') { event.preventDefault(); finish(true); } }} />
  </div>;
}
