import { OfficeDialog } from '@vibeoffice/core';
import { createPortal } from 'react-dom';
import { useEffect, useRef, useState } from 'react';
import './format-cells-dialog.css';

interface Catalog { title: string; content: HTMLElement; protection: { locked: boolean; hidden: boolean } }
interface Runtime {
  dlg?: { reactFormatCellsCatalog?(): Catalog | null; reactFormatCellsSetProtection?(locked: boolean, hidden: boolean): boolean; reactFormatCellsApply?(): boolean; reactFormatCellsClose?(): void };
  bus?: { on(event: string, listener: () => void): () => void };
  t?(message: string): string;
}

function runtime(): Runtime | undefined { return (window as Window & { L?: Runtime }).L; }

/** React-owned lifecycle and dialog shell for Excelx's full six-tab cell formatter. */
export function ExcelxFormatCellsDialog() {
  const [catalog, setCatalog] = useState<Catalog | null>(null);
  const [protectionHost, setProtectionHost] = useState<HTMLElement | null>(null);
  const [protection, setProtection] = useState({ locked: true, hidden: false });
  const contentHost = useRef<HTMLDivElement>(null);
  useEffect(() => {
    const refresh = () => {
      const next = runtime()?.dlg?.reactFormatCellsCatalog?.() || null;
      setCatalog(next);
      if (next) setProtection(next.protection);
    };
    const off = runtime()?.bus?.on('excelx-format-cells-dialog', refresh);
    return () => { off?.(); runtime()?.dlg?.reactFormatCellsClose?.(); };
  }, []);
  useEffect(() => {
    if (!catalog || !contentHost.current) return;
    contentHost.current.replaceChildren(catalog.content);
    return () => { catalog.content.remove(); };
  }, [catalog]);
  useEffect(() => {
    if (!catalog) { setProtectionHost(null); return; }
    setProtectionHost(catalog.content.querySelector<HTMLElement>('.fc-prot'));
  }, [catalog]);
  if (!catalog) return null;
  const translate = (message: string) => runtime()?.t?.(message) || message;
  const close = () => runtime()?.dlg?.reactFormatCellsClose?.();
  return <OfficeDialog title={translate(catalog.title)} className="excelx-format-cells-dialog" style={{ width: 'min(96vw, 520px)' }} onRequestClose={close}>
    <div ref={contentHost} className="excelx-format-cells-dialog__content" />
    {protectionHost && createPortal(<div className="excelx-format-cells-dialog__protection">
      <label><input type="checkbox" aria-label={translate('&Locked').replace('&', '')} checked={protection.locked} onChange={(event) => { const next = { ...protection, locked: event.currentTarget.checked }; setProtection(next); runtime()?.dlg?.reactFormatCellsSetProtection?.(next.locked, next.hidden); }} />{translate('&Locked').replace('&', '')}</label>
      <label><input type="checkbox" aria-label={translate('H&idden').replace('&', '')} checked={protection.hidden} onChange={(event) => { const next = { ...protection, hidden: event.currentTarget.checked }; setProtection(next); runtime()?.dlg?.reactFormatCellsSetProtection?.(next.locked, next.hidden); }} />{translate('H&idden').replace('&', '')}</label>
      <p className="hint">{translate('Locking cells or hiding formulas has no effect unless the worksheet is protected. To protect the worksheet, choose Protection from the Tools menu, and then choose Protect Sheet. A password is optional.')}</p>
    </div>, protectionHost)}
    <div className="excelx-format-cells-dialog__actions">
      <span />
      <button type="button" className="primary" onClick={() => runtime()?.dlg?.reactFormatCellsApply?.()}>{translate('OK')}</button>
      <button type="button" onClick={close}>{translate('Cancel')}</button>
    </div>
  </OfficeDialog>;
}
