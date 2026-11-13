import { OfficeDialog } from '@vibeoffice/core';
import { useEffect, useState } from 'react';

type ScenarioCell = { r: number; c: number; label?: string; v: string | number };
type Scenario = { name: string; comment: string; cells: ScenarioCell[] };
type Edit = { index: number; name: string; comment: string; reference: string };
type Catalog = { stage: 'manager' | 'edit' | 'values'; selected: number; scenarios: Scenario[]; edit?: Edit; refs?: ScenarioCell[] };
type Runtime = { dlg?: { reactScenarioCatalog?(): Catalog | null; reactScenarioAction?(action: string, ...args: unknown[]): unknown; reactScenarioClose?(): void }; bus?: { on(event: string, listener: () => void): () => void }; t?(message: string): string };
function runtime(): Runtime | undefined { return (window as Window & { L?: Runtime }).L; }

export function ExcelxScenariosDialog() {
  const [catalog, setCatalog] = useState<Catalog | null>(null);
  const [stage, setStage] = useState<Catalog['stage']>('manager');
  const [edit, setEdit] = useState<Edit>({ index: -1, name: '', comment: '', reference: '' });
  const [values, setValues] = useState<string[]>([]);
  useEffect(() => {
    const refresh = () => { const next = runtime()?.dlg?.reactScenarioCatalog?.() || null; setCatalog(next); if (next) { setStage(next.stage); if (next.edit) setEdit(next.edit); } };
    const off = runtime()?.bus?.on('excelx-scenarios-dialog', refresh);
    return () => { off?.(); runtime()?.dlg?.reactScenarioClose?.(); };
  }, []);
  if (!catalog) return null;
  const t = (message: string) => runtime()?.t?.(message) || message;
  const run = (action: string, ...args: unknown[]) => runtime()?.dlg?.reactScenarioAction?.(action, ...args);
  const close = () => runtime()?.dlg?.reactScenarioClose?.();
  const scenario = catalog.scenarios[catalog.selected];
  const begin = (index: number) => { setEdit({ index, name: index >= 0 ? catalog.scenarios[index]?.name || '' : '', comment: index >= 0 ? catalog.scenarios[index]?.comment || '' : '', reference: index >= 0 ? catalog.scenarios[index]?.cells.map((cell) => cell.label).join(',') || '' : '' }); run('begin', index); };
  const prepare = () => { const result = run('prepare', edit.reference, edit.name, edit.comment); if (Array.isArray(result)) setValues(result.map(String)); };
  const save = () => { run('save', values); };
  const title = stage === 'manager' ? t('Scenario Manager') : stage === 'edit' ? t(edit.index >= 0 ? 'Edit Scenario' : 'Add Scenario') : t('Scenario Values');
  return <OfficeDialog title={title} className="excelx-scenarios-dialog" style={{ width: 'min(94vw, 500px)' }} onRequestClose={close}>
    {stage === 'manager' && <div className="excelx-scenarios-dialog__manager"><label>{t('Sc&enarios:')}<select size={Math.min(7, Math.max(2, catalog.scenarios.length))} value={scenario ? catalog.selected : ''} onChange={(event) => run('select', Number(event.currentTarget.value))}>{catalog.scenarios.map((item, index) => <option key={`${item.name}-${index}`} value={index}>{item.name}</option>)}</select></label>{scenario ? <div className="excelx-scenarios-dialog__info">{t('Changing cells: {0}').replace('{0}', scenario.cells.map((cell) => cell.label || '').join(','))}{'\n'}{t('Comment: {0}').replace('{0}', scenario.comment || '')}</div> : <p>{t('No Scenarios defined. Choose Add to add scenarios.')}</p>}<div className="excelx-scenarios-dialog__actions"><button type="button" disabled={!scenario} onClick={() => run('show')}>{t('&Show').replace('&', '')}</button><button type="button" onClick={() => begin(-1)}>{t('&Add...').replace('&', '')}</button><button type="button" disabled={!scenario} onClick={() => run('delete')}>{t('&Delete').replace('&', '')}</button><button type="button" disabled={!scenario} onClick={() => begin(catalog.selected)}>{t('&Edit...').replace('&', '')}</button><button type="button" disabled={!scenario} onClick={() => run('summary')}>{t('S&ummary...').replace('&', '')}</button><span /><button type="button" className="primary" onClick={close}>{t('Close')}</button></div></div>}
    {stage === 'edit' && <div className="excelx-scenarios-dialog__edit"><label>{t('Scenario &name:')}<input value={edit.name} onChange={(event) => { const name = event.currentTarget.value; setEdit((old) => ({ ...old, name })); }} /></label><label>{t('Changing &cells:')}<input value={edit.reference} onChange={(event) => { const reference = event.currentTarget.value; setEdit((old) => ({ ...old, reference })); }} /></label><label>{t('Co&mment:')}<textarea rows={3} value={edit.comment} onChange={(event) => { const comment = event.currentTarget.value; setEdit((old) => ({ ...old, comment })); }} /></label><div className="excelx-scenarios-dialog__actions"><button type="button" onClick={() => run('back')}>{t('Cancel')}</button><span /><button type="button" className="primary" disabled={!edit.name.trim()} onClick={prepare}>{t('OK')}</button></div></div>}
    {stage === 'values' && <div className="excelx-scenarios-dialog__values"><p>{t('Enter values for each of the changing cells.')}</p>{(catalog.refs || []).map((cell, index) => <label key={`${cell.r}:${cell.c}`}>{cell.label}<input value={values[index] || ''} onChange={(event) => { const value = event.currentTarget.value; setValues((old) => old.map((current, i) => i === index ? value : current)); }} /></label>)}<div className="excelx-scenarios-dialog__actions"><button type="button" onClick={() => run('back')}>{t('Cancel')}</button><span /><button type="button" className="primary" onClick={save}>{t('OK')}</button></div></div>}
  </OfficeDialog>;
}
