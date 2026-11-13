import { useEffect, useState } from 'react';

type Item = { index: number; book: string; sheet: string; name: string; cell: string; value: string; formula: string };
type Catalog = { items: Item[]; selected: number };
type Runtime = { dlg?: { reactWatchCatalog?(): Catalog | null; reactWatchAction?(action: string, index?: number): boolean }; bus?: { on(event: string, listener: () => void): () => void }; t?(message: string): string };
function runtime(): Runtime | undefined { return (window as Window & { L?: Runtime }).L; }

/** React modeless Watch Window; watched cells and their live values remain workbook-engine state. */
export function ExcelxWatchWindow() {
  const [catalog, setCatalog] = useState<Catalog | null>(null);
  useEffect(() => {
    const refresh = () => setCatalog(runtime()?.dlg?.reactWatchCatalog?.() || null);
    const bus = runtime()?.bus;
    const offs = ['excelx-watch-window', 'changed'].map((event) => bus?.on(event, refresh)).filter((off): off is () => void => !!off);
    return () => offs.forEach((off) => off());
  }, []);
  if (!catalog) return null;
  const t = (message: string) => runtime()?.t?.(message) || message;
  const action = runtime()?.dlg?.reactWatchAction;
  return <section className="dlg excelx-watch-window" role="dialog" aria-label={t('Watch Window')} style={{ position: 'fixed', zIndex: 1200, top: '8vh', right: '3vw', width: 'min(94vw, 760px)' }}>
    <div className="dlg-ttl">{t('Watch Window')}</div>
    <div className="dlg-body">
      <div className="excelx-watch-window__tools"><button type="button" onClick={() => action?.('add')}>{t('&Add Watch...').replace('&', '')}</button><button type="button" disabled={catalog.selected < 0} onClick={() => action?.('delete', catalog.selected)}>{t('&Delete Watch').replace('&', '')}</button></div>
      <div className="excelx-watch-window__table"><table className="fr-table"><thead><tr>{['Book|workbook', 'Sheet', 'Name', 'Cell', 'Value', 'Formula'].map((heading) => <th key={heading}>{t(heading)}</th>)}</tr></thead><tbody>
        {catalog.items.map((item) => <tr key={`${item.book}-${item.sheet}-${item.cell}`} className={item.index === catalog.selected ? 'on' : ''} aria-selected={item.index === catalog.selected} onClick={() => action?.('select', item.index)} onDoubleClick={() => action?.('activate', item.index)}><td translate="no">{item.book}</td><td translate="no">{item.sheet}</td><td translate="no">{item.name}</td><td translate="no">{item.cell}</td><td translate="no">{item.value}</td><td translate="no">{item.formula}</td></tr>)}
      </tbody></table></div>
    </div>
    <div className="dlg-foot"><button type="button" className="primary" onClick={() => action?.('close')}>{t('Close')}</button></div>
  </section>;
}
