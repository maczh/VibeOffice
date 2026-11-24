import { OfficeDialog } from '@vibeoffice/core';
import { useEffect, useId, useState } from 'react';
import { usePptxRuntime } from './editor-host.js';

type PropertyKey = 'title' | 'subject' | 'author' | 'company' | 'category' | 'keywords' | 'comments';
interface PropertiesCatalog { title: string; values: Record<PropertyKey, string>; stats: Array<[string, string | number]> }
interface PropertiesRuntime {
  dlg?: { reactPropertiesCatalog?(): PropertiesCatalog | null; reactPropertiesUpdate?(key: PropertyKey, value: string): boolean; reactPropertiesClose?(): void };
  bus?: { on(event: string, listener: () => void): () => void };
  t?(message: string, ...args: unknown[]): string;
}
const fields: Array<[PropertyKey, string]> = [['title', '&Title:'], ['subject', '&Subject:'], ['author', '&Author:'], ['company', 'C&ompany:'], ['category', 'Ca&tegory:'], ['keywords', '&Keywords:']];

/** React-owned presentation metadata and statistics; live property updates stay in the presentation model. */
export function PptxPropertiesDialog() {
  const runtime = usePptxRuntime<PropertiesRuntime>();
  const [catalog, setCatalog] = useState<PropertiesCatalog | null>(null);
  const [tab, setTab] = useState(0);
  const id = useId().replaceAll(':', '');
  const tabId = (index: 0 | 1) => `${id}-properties-tab-${index}`;
  const panelId = (index: 0 | 1) => `${id}-properties-panel-${index}`;
  useEffect(() => {
    const refresh = () => { const next = runtime?.dlg?.reactPropertiesCatalog?.() || null; setCatalog(next ? { ...next, values: { ...next.values }, stats: [...next.stats] } : null); };
    const off = runtime?.bus?.on('pptx-properties-dialog', refresh);
    return () => off?.();
  }, []);
  if (!catalog) return null;
  const t = (message: string) => runtime?.t?.(message) || message;
  const update = (key: PropertyKey, value: string) => { runtime?.dlg?.reactPropertiesUpdate?.(key, value); setCatalog({ ...catalog, values: { ...catalog.values, [key]: value } }); };
  const close = () => runtime?.dlg?.reactPropertiesClose?.();
  return <OfficeDialog title={t(catalog.title)} className="pptx-properties-dialog" style={{ width: 'min(94vw, 520px)' }} onRequestClose={close}>
    <div role="tablist" aria-label={t('Properties')} className="tabs-head">{([0, 1] as const).map((index) => <button key={index} id={tabId(index)} type="button" role="tab" aria-controls={panelId(index)} aria-selected={tab === index} tabIndex={tab === index ? 0 : -1} onClick={() => setTab(index)} onKeyDown={(event) => {
      const next = event.key === 'ArrowRight' ? (index + 1) % 2
        : event.key === 'ArrowLeft' ? (index - 1 + 2) % 2
          : event.key === 'Home' ? 0
            : event.key === 'End' ? 1 : -1;
      if (next < 0) return;
      event.preventDefault();
      setTab(next);
      event.currentTarget.parentElement?.querySelectorAll<HTMLButtonElement>('[role="tab"]')[next]?.focus();
    }}>{t(index === 0 ? 'Summary' : 'Statistics')}</button>)}</div>
    <section id={panelId(0)} aria-labelledby={tabId(0)} role="tabpanel" className="col" hidden={tab !== 0}>{fields.map(([key, label]) => <label key={key}>{t(label.replace('&', ''))}<input aria-label={t(label.replace('&', ''))} type="text" value={catalog.values[key]} onChange={(event) => update(key, event.currentTarget.value)} /></label>)}
      <label>{t('&Comments:').replace('&', '')}<textarea aria-label={t('&Comments:').replace('&', '')} rows={3} value={catalog.values.comments} onChange={(event) => update('comments', event.currentTarget.value)} /></label>
    </section>
    <section id={panelId(1)} aria-labelledby={tabId(1)} role="tabpanel" hidden={tab !== 1}><table className="kbd-table"><tbody>{catalog.stats.map(([label, value]) => <tr key={label}><td>{t(String(label))}</td><td translate="no">{value}</td></tr>)}</tbody></table></section>
    <div className="vibeoffice-dialog__actions"><button type="button" onClick={close}>{t('OK')}</button></div>
  </OfficeDialog>;
}
