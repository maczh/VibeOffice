import { OfficeDialog } from '@vibeoffice/core';
import { useEffect, useState } from 'react';
import { usePptxRuntime } from './editor-host.js';

interface InsertTableCatalog { rows: number; columns: number }
interface InsertTableRuntime {
  dlg?: { reactInsertTableCatalog?(): InsertTableCatalog | null; reactInsertTableAction?(action: string, rows?: number, columns?: number): boolean };
  bus?: { on(event: string, listener: () => void): () => void };
  t?(message: string): string;
}

/** React-owned Insert Table dimensions; insertion remains the caller's existing editor callback. */
export function PptxInsertTableDialog() {
  const runtime = usePptxRuntime<InsertTableRuntime>();
  const [catalog, setCatalog] = useState<InsertTableCatalog | null>(null);
  const [rows, setRows] = useState(2);
  const [columns, setColumns] = useState(3);
  useEffect(() => {
    const refresh = () => { const next = runtime?.dlg?.reactInsertTableCatalog?.() || null; setCatalog(next); if (next) { setRows(next.rows); setColumns(next.columns); } };
    const off = runtime?.bus?.on('pptx-insert-table-dialog', refresh);
    return () => off?.();
  }, []);
  if (!catalog) return null;
  const t = (message: string) => runtime?.t?.(message) || message;
  const close = () => runtime?.dlg?.reactInsertTableAction?.('close');
  const valid = Number.isInteger(rows) && rows >= 1 && rows <= 75 && Number.isInteger(columns) && columns >= 1 && columns <= 75;
  return <OfficeDialog title={t('Insert Table')} className="pptx-insert-table-dialog" style={{ width: 'min(92vw, 350px)' }} onRequestClose={close}>
    <label>{t('Number of &columns:').replace('&', '')}<input aria-label={t('Number of &columns:').replace('&', '')} type="number" min="1" max="75" step="1" value={columns} onChange={(event) => setColumns(Number(event.currentTarget.value))} /></label>
    <label>{t('Number of &rows:').replace('&', '')}<input aria-label={t('Number of &rows:').replace('&', '')} type="number" min="1" max="75" step="1" value={rows} onChange={(event) => setRows(Number(event.currentTarget.value))} /></label>
    <div className="vibeoffice-dialog__actions"><button className="primary" type="button" disabled={!valid} onClick={() => runtime?.dlg?.reactInsertTableAction?.('apply', rows, columns)}>{t('OK')}</button><button type="button" onClick={close}>{t('Cancel')}</button></div>
  </OfficeDialog>;
}
