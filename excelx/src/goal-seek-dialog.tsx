import { OfficeDialog } from '@vibeoffice/core';
import { useEffect, useState } from 'react';

interface GoalSeekForm { stage: 'form'; target: string; goal: string; changing: string }
interface GoalSeekStatus { stage: 'status'; ok: boolean; cell: string; goal: string; current: string }
type GoalSeekState = GoalSeekForm | GoalSeekStatus;
interface GoalSeekRuntime {
  dlg?: { reactGoalSeekCatalog?(): GoalSeekState | null; reactGoalSeekApply?(state: GoalSeekForm): boolean; reactGoalSeekStatusAction?(undo: boolean): void; reactGoalSeekClose?(): void };
  bus?: { on(event: string, listener: () => void): () => void };
  t?(message: string, ...values: unknown[]): string;
}
function runtime(): GoalSeekRuntime | undefined { return (window as Window & { L?: GoalSeekRuntime }).L; }

export function ExcelxGoalSeekDialog() {
  const [catalog, setCatalog] = useState<GoalSeekState | null>(null);
  const [state, setState] = useState<GoalSeekState>({ stage: 'form', target: '', goal: '', changing: '' });
  useEffect(() => {
    const refresh = () => { const next = runtime()?.dlg?.reactGoalSeekCatalog?.() || null; setCatalog(next); if (next) setState(next); };
    const off = runtime()?.bus?.on('excelx-goal-seek-dialog', refresh);
    return () => { off?.(); if (runtime()?.dlg?.reactGoalSeekCatalog?.()) runtime()?.dlg?.reactGoalSeekClose?.(); };
  }, []);
  if (!catalog || !state) return null;
  const translate = (message: string, ...values: unknown[]) => runtime()?.t?.(message, ...values) || message;
  const close = () => runtime()?.dlg?.reactGoalSeekClose?.();
  if (state.stage === 'status') {
    const title = state.ok ? translate('Goal Seeking with Cell {0} found a solution.', state.cell) : translate('Goal Seeking with Cell {0} may not have found a solution.', state.cell);
    return <OfficeDialog title={translate('Goal Seek Status')} className="excelx-goal-seek-status-dialog" style={{ width: 'min(92vw, 380px)' }} onRequestClose={close}>
      <div className="excelx-goal-seek-status-dialog__body"><p>{title}</p><table><tbody><tr><th>{translate('Target value:')}</th><td>{state.goal}</td></tr><tr><th>{translate('Current value:')}</th><td>{state.current}</td></tr></tbody></table></div>
      <div className="excelx-goal-seek-status-dialog__actions"><button type="button" className="primary" onClick={() => runtime()?.dlg?.reactGoalSeekStatusAction?.(false)}>{translate('OK')}</button><button type="button" onClick={() => runtime()?.dlg?.reactGoalSeekStatusAction?.(true)}>{translate('Cancel')}</button></div>
    </OfficeDialog>;
  }
  const update = (patch: Partial<GoalSeekForm>) => setState((current) => current.stage === 'form' ? { ...current, ...patch } : current);
  const apply = () => { if (state.stage === 'form') runtime()?.dlg?.reactGoalSeekApply?.(state); };
  return <OfficeDialog title={translate('Goal Seek')} className="excelx-goal-seek-dialog" style={{ width: 'min(92vw, 380px)' }} onRequestClose={close}>
    <div className="excelx-goal-seek-dialog__body">
      <label>{translate('S&et cell:')}<input type="text" value={state.target} onChange={(event) => update({ target: event.currentTarget.value })} /></label>
      <label>{translate('To &value:')}<input type="text" value={state.goal} onChange={(event) => update({ goal: event.currentTarget.value })} /></label>
      <label>{translate('By &changing cell:')}<input type="text" value={state.changing} onChange={(event) => update({ changing: event.currentTarget.value })} /></label>
    </div>
    <div className="excelx-goal-seek-dialog__actions"><button type="button" className="primary" onClick={apply}>{translate('OK')}</button><button type="button" onClick={close}>{translate('Cancel')}</button></div>
  </OfficeDialog>;
}
