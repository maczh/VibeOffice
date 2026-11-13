import { OfficeDialog } from '@vibeoffice/core';
import { useEffect, useState } from 'react';

interface DataFormCatalog { title: string; fields: string[]; values: string[]; readOnly: boolean[]; count: string; criteriaMode: boolean }
type DataFormAction = 'new' | 'delete' | 'restore' | 'prev' | 'next' | 'criteria' | 'enter' | 'close';
interface DataFormRuntime {
  dlg?: { reactDataFormCatalog?(): DataFormCatalog | null; reactDataFormAction?(action: DataFormAction, values?: string[]): boolean; reactDataFormClose?(): void };
  bus?: { on(event: string, listener: () => void): () => void };
  t?(message: string): string;
}
function runtime(): DataFormRuntime | undefined { return (window as Window & { L?: DataFormRuntime }).L; }

export function ExcelxDataFormDialog() {
  const [catalog, setCatalog] = useState<DataFormCatalog | null>(null);
  const [values, setValues] = useState<string[]>([]);
  useEffect(() => {
    const refresh = () => { const next = runtime()?.dlg?.reactDataFormCatalog?.() || null; setCatalog(next); if (next) setValues(next.values); };
    const off = runtime()?.bus?.on('excelx-data-form-dialog', refresh);
    return () => { off?.(); runtime()?.dlg?.reactDataFormClose?.(); };
  }, []);
  if (!catalog) return null;
  const translate = (message: string) => runtime()?.t?.(message) || message;
  const action = (name: DataFormAction) => runtime()?.dlg?.reactDataFormAction?.(name, values);
  const close = () => action('close');
  const update = (index: number, value: string) => setValues((current) => current.map((item, i) => i === index ? value : item));
  const buttons: { action: DataFormAction; label: string }[] = [
    { action: 'new', label: 'Ne&w' }, { action: 'delete', label: '&Delete' }, { action: 'restore', label: '&Restore' },
    { action: 'prev', label: 'Find &Prev' }, { action: 'next', label: 'Find &Next' }, { action: 'criteria', label: catalog.criteriaMode ? '&Form' : '&Criteria' },
  ];
  return <OfficeDialog title={catalog.title} className="excelx-data-form-dialog" style={{ width: 'min(94vw, 620px)' }} onRequestClose={close}>
    <div className="excelx-data-form-dialog__body">
      <div className="excelx-data-form-dialog__fields">{catalog.fields.map((field, index) => <label key={`${field}-${index}`}>{translate(field)}<input type="text" value={values[index] || ''} readOnly={catalog.readOnly[index] && !catalog.criteriaMode} onChange={(event) => update(index, event.currentTarget.value)} onKeyDown={(event) => { if (event.key === 'Enter') { event.preventDefault(); action('enter'); } }} /></label>)}</div>
      <aside className="excelx-data-form-dialog__commands"><output>{translate(catalog.count)}</output>{buttons.map(({ action: command, label }) => <button type="button" key={command} onClick={() => action(command)}>{translate(label).replace('&', '')}</button>)}</aside>
    </div>
    <div className="excelx-data-form-dialog__actions"><button type="button" className="primary" onClick={close}>{translate('Cl&ose').replace('&', '')}</button></div>
  </OfficeDialog>;
}
