import { OfficeDialog } from '@vibeoffice/core';
import { useEffect, useState } from 'react';

type PropertyKey = 'title' | 'subject' | 'author' | 'company' | 'category' | 'keywords' | 'comments';
interface PropertiesCatalog { title: string; values: Record<PropertyKey, string>; stats: Array<[string, string | number]> }
interface PropertiesRuntime {
  dlg?: { reactPropertiesCatalog?(): PropertiesCatalog | null; reactPropertiesUpdate?(key: PropertyKey, value: string): boolean; reactPropertiesClose?(): void };
  bus?: { on(event: string, listener: () => void): () => void };
  t?(message: string, ...args: unknown[]): string;
}
function runtime(): PropertiesRuntime | undefined { return (window as Window & { L?: PropertiesRuntime }).L; }
const fields: Array<[PropertyKey, string]> = [['title', '&Title:'], ['subject', '&Subject:'], ['author', '&Author:'], ['company', 'C&ompany:'], ['category', 'Ca&tegory:'], ['keywords', '&Keywords:']];

/** React-owned presentation metadata and statistics; live property updates stay in the presentation model. */
export function PptxPropertiesDialog() {
  const [catalog, setCatalog] = useState<PropertiesCatalog | null>(null);
  const [tab, setTab] = useState(0);
  useEffect(() => {
    const refresh = () => { const next = runtime()?.dlg?.reactPropertiesCatalog?.() || null; setCatalog(next ? { ...next, values: { ...next.values }, stats: [...next.stats] } : null); };
    const off = runtime()?.bus?.on('pptx-properties-dialog', refresh);
    return () => off?.();
  }, []);
  if (!catalog) return null;
  const t = (message: string) => runtime()?.t?.(message) || message;
  const update = (key: PropertyKey, value: string) => { runtime()?.dlg?.reactPropertiesUpdate?.(key, value); setCatalog({ ...catalog, values: { ...catalog.values, [key]: value } }); };
  const close = () => runtime()?.dlg?.reactPropertiesClose?.();
  return <OfficeDialog title={t(catalog.title)} className="pptx-properties-dialog" style={{ width: 'min(94vw, 520px)' }} onRequestClose={close}>
    <div role="tablist" aria-label={t('Properties')} className="tabs-head"><button type="button" role="tab" aria-selected={tab === 0} onClick={() => setTab(0)}>{t('Summary')}</button><button type="button" role="tab" aria-selected={tab === 1} onClick={() => setTab(1)}>{t('Statistics')}</button></div>
    {tab === 0 ? <section role="tabpanel" className="col">{fields.map(([key, label]) => <label key={key}>{t(label.replace('&', ''))}<input aria-label={t(label.replace('&', ''))} type="text" value={catalog.values[key]} onChange={(event) => update(key, event.currentTarget.value)} /></label>)}
      <label>{t('&Comments:').replace('&', '')}<textarea aria-label={t('&Comments:').replace('&', '')} rows={3} value={catalog.values.comments} onChange={(event) => update('comments', event.currentTarget.value)} /></label>
    </section> : <section role="tabpanel"><table className="kbd-table"><tbody>{catalog.stats.map(([label, value]) => <tr key={label}><td>{t(String(label))}</td><td translate="no">{value}</td></tr>)}</tbody></table></section>}
    <div className="vibeoffice-dialog__actions"><button type="button" onClick={close}>{t('OK')}</button></div>
  </OfficeDialog>;
}
