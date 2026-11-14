import { OfficeDialog } from '@vibeoffice/core';
import { useEffect, useRef, useState } from 'react';

interface SetupShowState { kiosk: boolean; loop: boolean; noAnimation: boolean; penColor: string; range: boolean; from: number; to: number; useTimings: boolean; slideCount: number }
interface SetupShowRuntime {
  dlg?: { reactSetupShowCatalog?(): SetupShowState | null; reactSetupShowOpenPenColor?(anchor: HTMLElement): boolean; reactSetupShowAction?(action: string, values?: SetupShowState): boolean; reactSetupShowResolveColor?(color: string): string };
  bus?: { on(event: string, listener: () => void): () => void };
  t?(message: string): string;
}
function runtime(): SetupShowRuntime | undefined { return (window as Window & { L?: SetupShowRuntime }).L; }

/** React-owned slide-show setup; settings/history and the Office color menu stay engine-backed. */
export function PptxSetupShowDialog() {
  const [state, setState] = useState<SetupShowState | null>(null);
  const colorButton = useRef<HTMLButtonElement>(null);
  useEffect(() => {
    const refresh = () => { const next = runtime()?.dlg?.reactSetupShowCatalog?.() || null; setState(next ? { ...next } : null); };
    const off = runtime()?.bus?.on('pptx-setup-show-dialog', refresh);
    return () => off?.();
  }, []);
  if (!state) return null;
  const t = (message: string) => runtime()?.t?.(message) || message;
  const update = (changes: Partial<SetupShowState>) => setState({ ...state, ...changes });
  const close = () => runtime()?.dlg?.reactSetupShowAction?.('close');
  const color = state.penColor.startsWith('#') ? state.penColor : runtime()?.dlg?.reactSetupShowResolveColor?.(state.penColor) || state.penColor;
  return <OfficeDialog title={t('Set Up Show')} className="pptx-setup-show-dialog" style={{ width: 'min(94vw, 640px)' }} onRequestClose={close}>
    <div className="row" style={{ alignItems: 'start', gap: 24 }}>
      <div className="col">
        <fieldset><legend>{t('Show type')}</legend>
          <label><input type="radio" name="pptx-show-type" checked={!state.kiosk} onChange={() => update({ kiosk: false })} />{t('&Presented by a speaker (full screen)').replace('&', '')}</label>
          <label><input type="radio" name="pptx-show-type" checked={state.kiosk} onChange={() => update({ kiosk: true })} />{t('Browsed at a &kiosk (full screen)').replace('&', '')}</label>
        </fieldset>
        <fieldset><legend>{t('Show options')}</legend>
          <label><input type="checkbox" checked={state.loop} onChange={(event) => update({ loop: event.currentTarget.checked })} />{t("&Loop continuously until 'Esc'").replace('&', '')}</label>
          <label><input type="checkbox" checked={state.noAnimation} onChange={(event) => update({ noAnimation: event.currentTarget.checked })} />{t('Show without &animation').replace('&', '')}</label>
          <label>{t('P&en color:')} <button ref={colorButton} type="button" className="btn small" aria-label={t('Pen color')} onClick={() => { if (colorButton.current) runtime()?.dlg?.reactSetupShowOpenPenColor?.(colorButton.current); }}><span style={{ display: 'inline-block', width: 60, height: 14, background: color, border: '1px solid #555', verticalAlign: 'middle' }} /></button></label>
        </fieldset>
      </div>
      <div className="col">
        <fieldset><legend>{t('Show slides')}</legend>
          <label><input type="radio" name="pptx-show-range" checked={!state.range} onChange={() => update({ range: false })} />{t('&All').replace('&', '')}</label>
          <label><input type="radio" name="pptx-show-range" checked={state.range} onChange={() => update({ range: true })} />{t('&From:').replace('&', '')}
            <input type="number" aria-label={t('From slide')} min="1" max={Math.max(1, state.slideCount)} value={state.from} disabled={!state.range} onChange={(event) => update({ from: Number(event.currentTarget.value) })} />
            {t('To:')} <input type="number" aria-label={t('To slide')} min="1" max={Math.max(1, state.slideCount)} value={state.to} disabled={!state.range} onChange={(event) => update({ to: Number(event.currentTarget.value) })} />
          </label>
        </fieldset>
        <fieldset><legend>{t('Advance slides')}</legend>
          <label><input type="radio" name="pptx-show-advance" checked={!state.useTimings} onChange={() => update({ useTimings: false })} />{t('&Manually').replace('&', '')}</label>
          <label><input type="radio" name="pptx-show-advance" checked={state.useTimings} onChange={() => update({ useTimings: true })} />{t('&Using timings, if present').replace('&', '')}</label>
        </fieldset>
      </div>
    </div>
    <div className="vibeoffice-dialog__actions"><button className="primary" type="button" onClick={() => runtime()?.dlg?.reactSetupShowAction?.('apply', state)}>{t('OK')}</button><button type="button" onClick={close}>{t('Cancel')}</button></div>
  </OfficeDialog>;
}
