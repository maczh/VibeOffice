import { OfficeDialog } from '@vibeoffice/core';
import { useEffect, useState } from 'react';

interface WorkbookProperty { name: string; type?: string; value: string }
interface WorkbookPropertiesCatalog {
  title: string;
  summary: Record<string, string>;
  stats: { label: string; value: string }[];
  sheets: { name: string; hidden: boolean }[];
  names: string[];
  custom: WorkbookProperty[];
}
interface PropertiesRuntime {
  dlg?: { reactPropertiesCatalog?(): WorkbookPropertiesCatalog | null; reactPropertiesApply?(summary: Record<string, string>, custom: WorkbookProperty[]): boolean; reactPropertiesClose?(): void };
  bus?: { on(event: string, listener: () => void): () => void };
  t?(message: string, ...args: unknown[]): string;
}

function runtime(): PropertiesRuntime | undefined {
  return (window as Window & { L?: PropertiesRuntime }).L;
}

const summaryFields: [string, string][] = [
  ['title', '&Title:'], ['subject', '&Subject:'], ['creator', '&Author:'], ['manager', '&Manager:'],
  ['company', 'C&ompany:'], ['category', 'Ca&tegory:'], ['keywords', '&Keywords:'], ['description', 'C&omments:'],
];

export function ExcelxPropertiesDialog() {
  const [catalog, setCatalog] = useState<WorkbookPropertiesCatalog | null>(null);
  const [summary, setSummary] = useState<Record<string, string>>({});
  const [custom, setCustom] = useState<WorkbookProperty[]>([]);
  const [name, setName] = useState('');
  const [value, setValue] = useState('');
  const [tab, setTab] = useState('Summary');
  useEffect(() => {
    const refresh = () => {
      const next = runtime()?.dlg?.reactPropertiesCatalog?.() || null;
      setCatalog(next);
      if (next) { setSummary(next.summary); setCustom(next.custom); setTab('Summary'); setName(''); setValue(''); }
    };
    const off = runtime()?.bus?.on('excelx-properties-dialog', refresh);
    return () => {
      off?.();
      if (runtime()?.dlg?.reactPropertiesCatalog?.()) runtime()?.dlg?.reactPropertiesClose?.();
    };
  }, []);
  if (!catalog) return null;
  const translate = (message: string, ...args: unknown[]) => runtime()?.t?.(message, ...args) || message;
  const close = () => runtime()?.dlg?.reactPropertiesClose?.();
  const apply = () => runtime()?.dlg?.reactPropertiesApply?.(summary, custom);
  const tabs = ['Summary', 'Statistics', 'Contents', 'Custom'];
  const addProperty = () => {
    const propertyName = name.trim();
    if (!propertyName) return;
    setCustom((previous) => [...previous, { name: propertyName, type: 'lpwstr', value }]);
    setName('');
    setValue('');
  };
  return <OfficeDialog title={catalog.title} className="excelx-properties-dialog" style={{ width: 'min(94vw, 460px)' }} onRequestClose={close}>
    <div className="excelx-properties-dialog__tabs" role="tablist" aria-label={catalog.title}>
      {tabs.map((item) => <button key={item} type="button" role="tab" aria-selected={tab === item} onClick={() => setTab(item)}>{translate(item)}</button>)}
    </div>
    {tab === 'Summary' && <div className="excelx-properties-dialog__summary">
      {summaryFields.map(([key, label]) => { const text = translate(label).replace('&', ''); return <label key={key}>{text}<input aria-label={text} value={summary[key] || ''} onChange={(event) => { const fieldValue = event.currentTarget.value; setSummary((previous) => ({ ...previous, [key]: fieldValue })); }} /></label>; })}
    </div>}
    {tab === 'Statistics' && <table className="kv"><tbody>{catalog.stats.map(({ label, value: statistic }) => <tr key={label}><td>{translate(label)}</td><td translate="no">{statistic}</td></tr>)}</tbody></table>}
    {tab === 'Contents' && <div className="excelx-properties-dialog__contents">
      <strong>{translate('Worksheets')}</strong>
      {catalog.sheets.map(({ name: sheetName, hidden }) => <div key={sheetName} translate="no">{hidden ? translate('{0} (hidden)', sheetName) : sheetName}</div>)}
      <strong>{translate('Named Ranges')}</strong>
      {catalog.names.map((definedName) => <div key={definedName} translate="no">{definedName}</div>)}
    </div>}
    {tab === 'Custom' && <div className="excelx-properties-dialog__custom">
      <div className="excelx-properties-dialog__custom-add">
        <input aria-label={translate('Name')} placeholder={translate('Name')} value={name} onChange={(event) => { const fieldValue = event.currentTarget.value; setName(fieldValue); }} />
        <input aria-label={translate('Value')} placeholder={translate('Value')} value={value} onChange={(event) => { const fieldValue = event.currentTarget.value; setValue(fieldValue); }} />
        <button type="button" onClick={addProperty}>{translate('&Add').replace('&', '')}</button>
      </div>
      <table className="kv"><tbody>{custom.map((property, index) => <tr key={`${property.name}:${index}`}><td translate="no">{property.name}</td><td translate="no">{property.value}</td><td><button type="button" className="btn small" onClick={() => setCustom((previous) => previous.filter((_, current) => current !== index))}>{translate('Delete')}</button></td></tr>)}</tbody></table>
    </div>}
    <div className="excelx-properties-dialog__actions">
      <button type="button" className="primary" onClick={apply}>{translate('OK')}</button>
      <button type="button" onClick={close}>{translate('Cancel')}</button>
    </div>
  </OfficeDialog>;
}
