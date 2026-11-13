import { OfficeDialog } from '@vibeoffice/core';
import { useEffect, useState } from 'react';

type Condition = { kind: string; op: string; v1: string; v2: string; dxf: Record<string, unknown>; preview?: string; keep?: Record<string, unknown>; stop?: boolean };
type Catalog = { conditions: Condition[]; operators: [string, string][] };
type Runtime = { dlg?: { reactConditionalCatalog?(): Catalog | null; reactConditionalFormat?(index: number, state: Catalog): boolean; reactConditionalApply?(state: Catalog): boolean; reactConditionalClose?(): void }; bus?: { on(event: string, listener: () => void): () => void }; t?(message: string): string };
function runtime(): Runtime | undefined { return (window as Window & { L?: Runtime }).L; }
const blank = (): Condition => ({ kind: 'value', op: 'between', v1: '', v2: '', dxf: {}, preview: '' });

export function ExcelxConditionalFormatDialog() {
  const [catalog, setCatalog] = useState<Catalog | null>(null);
  const [conditions, setConditions] = useState<Condition[]>([]);
  useEffect(() => {
    const refresh = () => { const next = runtime()?.dlg?.reactConditionalCatalog?.() || null; setCatalog(next); if (next) setConditions(next.conditions); };
    const off = runtime()?.bus?.on('excelx-conditional-format-dialog', refresh);
    return () => { off?.(); runtime()?.dlg?.reactConditionalClose?.(); };
  }, []);
  if (!catalog) return null;
  const t = (message: string) => runtime()?.t?.(message) || message;
  const close = () => runtime()?.dlg?.reactConditionalClose?.();
  const update = (index: number, patch: Partial<Condition>) => setConditions((old) => old.map((condition, i) => i === index ? { ...condition, ...patch } : condition));
  const remove = (index: number) => setConditions((old) => { const next = old.filter((_, i) => i !== index); return next.length ? next : [blank()]; });
  return <OfficeDialog title={t('Conditional Formatting')} className="excelx-conditional-format-dialog" style={{ width: 'min(94vw, 620px)' }} onRequestClose={close}>
    <div className="excelx-conditional-format-dialog__conditions">{conditions.map((condition, index) => {
      const two = condition.kind === 'value' && ['between', 'notBetween'].includes(condition.op);
      return <fieldset key={index}><legend>{t('Condition {0}').replace('{0}', String(index + 1))}</legend>
        <label>{t('Condition type')}<select value={condition.kind} onChange={(event) => update(index, { kind: event.currentTarget.value })}><option value="value">{t('Cell Value Is')}</option><option value="formula">{t('Formula Is')}</option></select></label>
        {condition.kind === 'value' && <label>{t('Operator')}<select value={condition.op} onChange={(event) => update(index, { op: event.currentTarget.value })}>{catalog.operators.map(([key, label]) => <option key={key} value={key}>{t(label)}</option>)}</select></label>}
        <label>{condition.kind === 'formula' ? t('Formula:') : t('Value:')}<input type="text" value={condition.v1} onChange={(event) => update(index, { v1: event.currentTarget.value })} /></label>
        {two && <label>{t('and')}<input type="text" value={condition.v2} onChange={(event) => update(index, { v2: event.currentTarget.value })} /></label>}
        <div className="excelx-conditional-format-dialog__preview"><span>{t('Preview of format to use when condition is true:')}</span><div className="cf-prev" style={Object.fromEntries((condition.preview || '').split(';').filter(Boolean).map((part) => { const [key = '', ...value] = part.split(':'); return [key.replace(/-([a-z])/g, (_, letter: string) => letter.toUpperCase()), value.join(':')]; }))}>{Object.keys(condition.dxf || {}).length ? 'AaBbCcYyZz' : t('No Format Set')}</div><button type="button" onClick={() => runtime()?.dlg?.reactConditionalFormat?.(index, { ...catalog, conditions })}>{t('&Format...').replace('&', '')}</button><button type="button" disabled={conditions.length === 1} onClick={() => remove(index)}>{t('&Delete').replace('&', '')}</button></div>
      </fieldset>;
    })}</div>
    <div className="excelx-conditional-format-dialog__actions"><button type="button" disabled={conditions.length >= 3} onClick={() => setConditions((old) => [...old, blank()])}>{t('&Add >>').replace('&', '')}</button><span /><button type="button" className="primary" onClick={() => runtime()?.dlg?.reactConditionalApply?.({ ...catalog, conditions })}>{t('OK')}</button><button type="button" onClick={close}>{t('Cancel')}</button></div>
  </OfficeDialog>;
}
