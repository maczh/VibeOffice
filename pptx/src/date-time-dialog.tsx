import { OfficeDialog } from '@vibeoffice/core';
import { useEffect, useState } from 'react';

interface DateTimeFormat { id: string; label: string }
interface DateTimeCatalog { formats: DateTimeFormat[]; selected: string }
interface DateTimeRuntime {
  dlg?: { reactDateTimeCatalog?(): DateTimeCatalog | null; reactDateTimeApply?(format: string, automatic: boolean): boolean; reactDateTimeClose?(): void };
  bus?: { on(event: string, listener: () => void): () => void };
  t?(message: string): string;
}

function runtime(): DateTimeRuntime | undefined { return (window as Window & { L?: DateTimeRuntime }).L; }

/** React-owned Date and Time form; insertion and Word field creation stay in the presentation engine. */
export function PptxDateTimeDialog() {
  const [catalog, setCatalog] = useState<DateTimeCatalog | null>(null);
  const [selected, setSelected] = useState('');
  const [automatic, setAutomatic] = useState(false);
  useEffect(() => {
    const refresh = () => {
      const next = runtime()?.dlg?.reactDateTimeCatalog?.() || null;
      setCatalog(next);
      if (next) setSelected(next.selected);
    };
    const off = runtime()?.bus?.on('pptx-date-time-dialog', refresh);
    return () => off?.();
  }, []);
  if (!catalog) return null;
  const t = (message: string) => runtime()?.t?.(message) || message;
  const close = () => runtime()?.dlg?.reactDateTimeClose?.();
  const apply = () => { if (runtime()?.dlg?.reactDateTimeApply?.(selected, automatic)) return; };
  return <OfficeDialog title={t('Date and Time')} className="pptx-date-time-dialog" style={{ width: 'min(92vw, 340px)' }} onRequestClose={close}>
    <label htmlFor="pptx-date-time-format">{t('&Available formats:').replace('&', '')}</label>
    <select id="pptx-date-time-format" aria-label={t('&Available formats:').replace('&', '')} size={12} style={{ width: '100%' }} value={selected} onChange={(event) => setSelected(event.currentTarget.value)}>
      {catalog.formats.map((format) => <option key={format.id} value={format.id}>{format.label}</option>)}
    </select>
    <label><input type="checkbox" checked={automatic} onChange={(event) => setAutomatic(event.currentTarget.checked)} />{t('&Update automatically').replace('&', '')}</label>
    <div className="vibeoffice-dialog__actions"><button className="primary" type="button" onClick={apply}>{t('OK')}</button><button type="button" onClick={close}>{t('Cancel')}</button></div>
  </OfficeDialog>;
}
