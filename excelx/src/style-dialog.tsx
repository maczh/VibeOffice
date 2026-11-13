import { OfficeDialog } from '@vibeoffice/core';
import { useEffect, useState } from 'react';

type Description = { key: string; label: string; value: string };
type Catalog = { name: string; names: string[]; includes: Record<string, boolean>; descriptions: Description[] };
type Runtime = { dlg?: { reactStyleCatalog?(): Catalog | null; reactStyleUpdate?(name: string): void; reactStyleAction?(action: string, name: string, includes: Record<string, boolean>): boolean; reactStyleClose?(): void }; bus?: { on(event: string, listener: () => void): () => void }; t?(message: string): string };
function runtime(): Runtime | undefined { return (window as Window & { L?: Runtime }).L; }

export function ExcelxStyleDialog() {
  const [catalog, setCatalog] = useState<Catalog | null>(null);
  const [name, setName] = useState('Normal');
  const [includes, setIncludes] = useState<Record<string, boolean>>({});
  useEffect(() => {
    const refresh = () => { const next = runtime()?.dlg?.reactStyleCatalog?.() || null; setCatalog(next); if (next) { setName(next.name); setIncludes(next.includes); } };
    const off = runtime()?.bus?.on('excelx-style-dialog', refresh);
    return () => { off?.(); runtime()?.dlg?.reactStyleClose?.(); };
  }, []);
  if (!catalog) return null;
  const t = (message: string) => runtime()?.t?.(message) || message;
  const close = () => runtime()?.dlg?.reactStyleClose?.();
  const action = (kind: string) => runtime()?.dlg?.reactStyleAction?.(kind, name, includes);
  return <OfficeDialog title={t('Style')} className="excelx-style-dialog" style={{ width: 'min(94vw, 460px)' }} onRequestClose={close}>
    <div className="excelx-style-dialog__body"><label>{t('&Style name:')}<input list="excelx-style-names" value={name} onChange={(event) => { setName(event.currentTarget.value); runtime()?.dlg?.reactStyleUpdate?.(event.currentTarget.value); }} /></label><datalist id="excelx-style-names">{catalog.names.map((item) => <option key={item} value={item} />)}</datalist>
      <fieldset><legend>{t('Style includes')}</legend>{catalog.descriptions.map(({ key, label, value }) => <label className="excelx-style-dialog__include" key={key}><input type="checkbox" checked={!!includes[key]} onChange={(event) => setIncludes((old) => ({ ...old, [key]: event.currentTarget.checked }))} /><span>{t(label)}</span><small>{t(value)}</small></label>)}</fieldset>
    </div><div className="excelx-style-dialog__actions"><button type="button" onClick={() => action('add')}>{t('&Add').replace('&', '')}</button><button type="button" onClick={() => action('delete')}>{t('&Delete').replace('&', '')}</button><button type="button" onClick={() => action('modify')}>{t('&Modify...').replace('&', '')}</button><span /><button type="button" className="primary" onClick={() => action('apply')}>{t('OK')}</button><button type="button" onClick={close}>{t('Cancel')}</button></div>
  </OfficeDialog>;
}
