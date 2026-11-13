import { OfficeDialog } from '@vibeoffice/core';
import { useEffect, useState } from 'react';

interface MoveCopyBook { index: number; name: string; sheets: string[] }
interface MoveCopyCatalog { books: MoveCopyBook[]; book: number; before: number }
interface MoveCopyRuntime {
  dlg?: { reactMoveCopyCatalog?(): MoveCopyCatalog | null; reactMoveCopyApply?(book: number, before: number, copy: boolean): boolean; reactMoveCopyClose?(): void };
  bus?: { on(event: string, listener: () => void): () => void };
  t?(message: string): string;
}

function runtime(): MoveCopyRuntime | undefined {
  return (window as Window & { L?: MoveCopyRuntime }).L;
}

/** React UI for moving or copying selected worksheet(s) to an open workbook. */
export function ExcelxMoveCopySheetDialog() {
  const [catalog, setCatalog] = useState<MoveCopyCatalog | null>(null);
  const [book, setBook] = useState(0);
  const [before, setBefore] = useState(0);
  const [copy, setCopy] = useState(false);
  useEffect(() => {
    const refresh = () => {
      const next = runtime()?.dlg?.reactMoveCopyCatalog?.() || null;
      setCatalog(next);
      if (next) {
        setBook(next.book);
        setBefore(next.before);
        setCopy(false);
      }
    };
    const off = runtime()?.bus?.on('excelx-move-copy-sheet-dialog', refresh);
    return () => {
      off?.();
      if (runtime()?.dlg?.reactMoveCopyCatalog?.()) runtime()?.dlg?.reactMoveCopyClose?.();
    };
  }, []);
  if (!catalog) return null;
  const translate = (message: string) => runtime()?.t?.(message) || message;
  const target = catalog.books.find((item) => item.index === book);
  const close = () => runtime()?.dlg?.reactMoveCopyClose?.();
  const apply = () => { if (target) runtime()?.dlg?.reactMoveCopyApply?.(book, before, copy); };
  return <OfficeDialog title={translate('Move or Copy')} className="excelx-move-copy-sheet-dialog" style={{ width: 'min(92vw, 380px)' }} onRequestClose={close}>
    <div className="excelx-move-copy-sheet-dialog__form">
      <div>{translate('Move selected sheets')}</div>
      <label>{translate('&To book:').replace('&', '')}<select aria-label={translate('&To book:').replace('&', '')} value={book} onChange={(event) => { const nextBook = Number(event.currentTarget.value); const nextTarget = catalog.books.find((item) => item.index === nextBook); setBook(nextBook); setBefore(nextTarget?.sheets.length || 0); }}>
        {catalog.books.map((item) => <option key={item.index} value={item.index} translate="no">{item.name}</option>)}
      </select></label>
      <label>{translate('&Before sheet:').replace('&', '')}<select aria-label={translate('&Before sheet:').replace('&', '')} value={before} onChange={(event) => setBefore(Number(event.currentTarget.value))}>
        {target?.sheets.map((name, index) => <option key={`${index}:${name}`} value={index} translate="no">{name}</option>)}
        <option value={target?.sheets.length || 0}>{translate('(move to end)')}</option>
      </select></label>
      <label><input type="checkbox" checked={copy} onChange={(event) => setCopy(event.currentTarget.checked)} />{translate('&Create a copy').replace('&', '')}</label>
    </div>
    <div className="excelx-move-copy-sheet-dialog__actions">
      <button type="button" className="primary" disabled={!target} onClick={apply}>{translate('OK')}</button>
      <button type="button" onClick={close}>{translate('Cancel')}</button>
    </div>
  </OfficeDialog>;
}
