import { useEffect, useRef, useState } from 'react';

type Catalog = { initialText: string; left: number; top: number; width: number; height: number };
type Runtime = { dlg?: { reactCommentCatalog?(): Catalog | null; reactCommentFinish?(save: boolean, value?: string, width?: number, height?: number): boolean }; bus?: { on(event: string, listener: () => void): () => void } };
function runtime(): Runtime | undefined { return (window as Window & { L?: Runtime }).L; }

/** React-owned in-grid comment editor preserving the legacy floating-note interaction. */
export function ExcelxCommentEditor() {
  const [catalog, setCatalog] = useState<Catalog | null>(null);
  const [value, setValue] = useState('');
  const area = useRef<HTMLTextAreaElement>(null);
  const valueRef = useRef(value);
  valueRef.current = value;
  const finishRef = useRef<(save: boolean) => void>(() => {});
  useEffect(() => {
    const refresh = () => {
      const next = runtime()?.dlg?.reactCommentCatalog?.() || null;
      setCatalog(next);
      if (next) { setValue(next.initialText); requestAnimationFrame(() => area.current?.focus()); }
    };
    const off = runtime()?.bus?.on('excelx-comment-editor', refresh);
    refresh();
    return () => off?.();
  }, []);
  useEffect(() => {
    if (!catalog) return;
    const outside = (event: PointerEvent) => { if (area.current && !area.current.parentElement?.contains(event.target as Node)) finishRef.current(true); };
    document.addEventListener('pointerdown', outside, true);
    return () => document.removeEventListener('pointerdown', outside, true);
  }, [!!catalog]);
  if (!catalog) return null;
  const finish = (save: boolean) => runtime()?.dlg?.reactCommentFinish?.(save, valueRef.current, area.current?.parentElement?.offsetWidth || catalog.width, area.current?.parentElement?.offsetHeight || catalog.height);
  finishRef.current = finish;
  return <div className="cmt-edit" style={{ position: 'fixed', left: catalog.left, top: catalog.top, width: catalog.width, height: catalog.height }}>
    <textarea ref={area} aria-label="Comment" spellCheck onChange={(event) => { const text = event.currentTarget.value; valueRef.current = text; setValue(text); }} value={value} onKeyDown={(event) => { event.stopPropagation(); if (event.key === 'Escape') { event.preventDefault(); finish(true); } }} />
  </div>;
}
