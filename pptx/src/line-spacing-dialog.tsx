import { OfficeDialog } from '@vibeoffice/core';
import { useEffect, useState } from 'react';
import { usePptxRuntime } from './editor-host.js';

type SpacingUnit = 'l' | 'p';
type SpacingKey = 'line' | 'before' | 'after';
interface SpacingValue { value: number; unit: SpacingUnit }
interface LineSpacingCatalog { fontSize: number; line: SpacingValue; before: SpacingValue; after: SpacingValue }
interface LineSpacingRuntime {
  dlg?: { reactLineSpacingCatalog?(): LineSpacingCatalog | null; reactLineSpacingAction?(action: string, values?: LineSpacingCatalog): boolean };
  bus?: { on(event: string, listener: () => void): () => void };
  t?(message: string): string;
}
const keys: SpacingKey[] = ['line', 'before', 'after'];
const labels: Record<SpacingKey, string> = { line: 'Line spacing', before: 'Before paragraph', after: 'After paragraph' };

/** React-owned line-spacing controls; paragraph-format mutation remains in the presentation formatter. */
export function PptxLineSpacingDialog() {
  const runtime = usePptxRuntime<LineSpacingRuntime>();
  const [catalog, setCatalog] = useState<LineSpacingCatalog | null>(null);
  useEffect(() => {
    const refresh = () => setCatalog(runtime?.dlg?.reactLineSpacingCatalog?.() || null);
    const off = runtime?.bus?.on('pptx-line-spacing-dialog', refresh);
    return () => off?.();
  }, []);
  if (!catalog) return null;
  const t = (message: string) => runtime?.t?.(message) || message;
  const close = () => runtime?.dlg?.reactLineSpacingAction?.('close');
  const update = (key: SpacingKey, next: SpacingValue) => setCatalog({ ...catalog, [key]: next });
  const setUnit = (key: SpacingKey, unit: SpacingUnit) => {
    const current = catalog[key];
    if (current.unit === unit) return;
    const factor = (catalog.fontSize || 18) * 1.2;
    update(key, { unit, value: unit === 'p' ? Math.round(current.value * factor * 10) / 10 : Math.round(current.value / factor * 100) / 100 });
  };
  return <OfficeDialog title={t('Line Spacing')} className="pptx-line-spacing-dialog" style={{ width: 'min(92vw, 380px)' }} onRequestClose={close}>
    {keys.map((key) => <fieldset key={key}><legend>{t(labels[key])}</legend>
      <input aria-label={`${t(labels[key])} value`} type="number" min="0" max="1584" step={catalog[key].unit === 'p' ? '1' : '0.1'} value={catalog[key].value} onChange={(event) => update(key, { ...catalog[key], value: Number(event.currentTarget.value) })} />
      <select aria-label={`${t(labels[key])} unit`} value={catalog[key].unit} onChange={(event) => setUnit(key, event.currentTarget.value as SpacingUnit)}><option value="l">{t('Lines')}</option><option value="p">{t('Points')}</option></select>
    </fieldset>)}
    <div className="vibeoffice-dialog__actions"><button className="primary" type="button" onClick={() => runtime?.dlg?.reactLineSpacingAction?.('apply', catalog)}>{t('OK')}</button><button type="button" onClick={close}>{t('Cancel')}</button></div>
  </OfficeDialog>;
}
