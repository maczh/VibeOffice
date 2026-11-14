import { OfficeDialog } from '@vibeoffice/core';
import { useEffect, useState } from 'react';

interface GridCatalog { snap: boolean; size: number; show: boolean; guides: boolean }
interface GridRuntime {
  dlg?: { reactGridCatalog?(): GridCatalog | null; reactGridApply?(value: GridCatalog): void; reactGridClose?(): void };
  bus?: { on(event: string, listener: () => void): () => void };
  t?(message: string): string;
}

function runtime(): GridRuntime | undefined { return (window as Window & { L?: GridRuntime }).L; }
const SPACINGS: Array<[number, string]> = [[4.5, '1/16"'], [6, '0.083"'], [7.2, '0.1"'], [9, '0.125"'], [12, '0.167"'], [14.4, '0.2"'], [18, '0.25"'], [24, '0.333"'], [36, '0.5"'], [72, '1"']];

/** React-owned grid and guide preferences; drawing and snapping stay in the engine. */
export function PptxGridDialog() {
  const [catalog, setCatalog] = useState<GridCatalog | null>(null);
  const [draft, setDraft] = useState<GridCatalog>({ snap: false, size: 9, show: false, guides: false });
  useEffect(() => {
    const refresh = () => {
      const next = runtime()?.dlg?.reactGridCatalog?.() || null;
      setCatalog(next);
      if (next) setDraft(next);
    };
    const off = runtime()?.bus?.on('pptx-grid-dialog', refresh);
    return () => off?.();
  }, []);
  if (!catalog) return null;
  const t = (message: string) => runtime()?.t?.(message) || message;
  const close = () => runtime()?.dlg?.reactGridClose?.();
  const update = (key: keyof GridCatalog, value: boolean | number) => setDraft((current) => ({ ...current, [key]: value }));
  return <OfficeDialog title={t('Grid and Guides')} className="pptx-grid-dialog" style={{ width: 'min(92vw, 360px)' }} onRequestClose={close}>
    <fieldset><legend>{t('Snap to')}</legend><label><input type="checkbox" checked={draft.snap} onChange={(event) => update('snap', event.currentTarget.checked)} />{t('&Snap objects to grid').replace('&', '')}</label></fieldset>
    <fieldset><legend>{t('Grid settings')}</legend><label>{t('Spac&ing:').replace('&', '')}<select aria-label={t('Spacing')} value={draft.size} onChange={(event) => update('size', Number(event.currentTarget.value))}>{SPACINGS.map(([value, label]) => <option key={value} value={value}>{label}</option>)}</select></label><label><input type="checkbox" checked={draft.show} onChange={(event) => update('show', event.currentTarget.checked)} />{t('&Display grid on screen').replace('&', '')}</label></fieldset>
    <fieldset><legend>{t('Guide settings')}</legend><label><input type="checkbox" checked={draft.guides} onChange={(event) => update('guides', event.currentTarget.checked)} />{t('Display drawing &guides on screen').replace('&', '')}</label></fieldset>
    <div className="vibeoffice-dialog__actions"><button className="primary" type="button" onClick={() => runtime()?.dlg?.reactGridApply?.(draft)}>{t('OK')}</button><button type="button" onClick={close}>{t('Cancel')}</button></div>
  </OfficeDialog>;
}
