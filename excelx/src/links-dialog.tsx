import { OfficeDialog } from '@vibeoffice/core';
import { useEffect, useState } from 'react';

type Item = { source: string; type: string; update: string; status: string };
type Catalog = { items: Item[] };
type Runtime = { dlg?: { reactLinksCatalog?(): Catalog | null; reactLinksAction?(action: string): boolean }; bus?: { on(event: string, listener: () => void): () => void }; t?(message: string): string };
function runtime(): Runtime | undefined { return (window as Window & { L?: Runtime }).L; }

/** React-owned External Links list and Break Link action; cached formula values remain engine-authored. */
export function ExcelxLinksDialog() {
  const [catalog, setCatalog] = useState<Catalog | null>(null);
  useEffect(() => {
    const refresh = () => setCatalog(runtime()?.dlg?.reactLinksCatalog?.() || null);
    const off = runtime()?.bus?.on('excelx-links-dialog', refresh);
    return () => off?.();
  }, []);
  if (!catalog) return null;
  const t = (message: string) => runtime()?.t?.(message) || message;
  const close = () => runtime()?.dlg?.reactLinksAction?.('close');
  return <OfficeDialog title={t('Edit Links')} className="excelx-links-dialog" style={{ width: 'min(94vw, 600px)' }} onRequestClose={close}>
    <div className="excelx-links-dialog__body">
      <table><thead><tr>{['Source', 'Type', 'Update', 'Status'].map((heading) => <th key={heading}>{t(heading)}</th>)}</tr></thead>
        <tbody>{catalog.items.map((item, index) => <tr key={`${item.source}-${index}`}><td translate="no">{item.source}</td><td>{t(item.type)}</td><td>{t(item.update)}</td><td>{t(item.status)}</td></tr>)}</tbody>
      </table>
      <p>{t('Formulas that refer to other workbooks use the values Excel saved with this file. Break Link replaces those formulas with their values.')}</p>
    </div>
    <div className="excelx-links-dialog__actions"><button type="button" disabled={!catalog.items.length} onClick={() => runtime()?.dlg?.reactLinksAction?.('break')}>{t('&Break Link').replace('&', '')}</button><span /><button type="button" className="primary" onClick={close}>{t('Close')}</button></div>
  </OfficeDialog>;
}
