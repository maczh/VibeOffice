import { OfficeDialog } from '@vibeoffice/core';
import { useEffect, useState } from 'react';

interface ReplaceFontsCatalog { used: string[]; fonts: string[] }
interface ReplaceFontsRuntime {
  dlg?: { reactReplaceFontsCatalog?(): ReplaceFontsCatalog | null; reactReplaceFontsAction?(action: string, from?: string, to?: string): boolean };
  bus?: { on(event: string, listener: () => void): () => void };
  t?(message: string): string;
}
function runtime(): ReplaceFontsRuntime | undefined { return (window as Window & { L?: ReplaceFontsRuntime }).L; }

/** React-owned Replace Font selectors; document-wide replacement remains in the presentation engine. */
export function PptxReplaceFontsDialog() {
  const [catalog, setCatalog] = useState<ReplaceFontsCatalog | null>(null);
  const [from, setFrom] = useState('');
  const [to, setTo] = useState('Arial');
  useEffect(() => {
    const refresh = () => {
      const next = runtime()?.dlg?.reactReplaceFontsCatalog?.() || null;
      setCatalog(next);
      if (next) { setFrom(next.used[0] || ''); setTo(next.fonts.includes('Arial') ? 'Arial' : next.fonts[0] || ''); }
    };
    const off = runtime()?.bus?.on('pptx-replace-fonts-dialog', refresh);
    return () => off?.();
  }, []);
  if (!catalog) return null;
  const t = (message: string) => runtime()?.t?.(message) || message;
  const act = (action: string) => runtime()?.dlg?.reactReplaceFontsAction?.(action, from, to);
  return <OfficeDialog title={t('Replace Font')} className="pptx-replace-fonts-dialog" style={{ width: 'min(92vw, 390px)' }} onRequestClose={() => act('close')}>
    <label>{t('&Replace:').replace('&', '')}<select aria-label={t('&Replace:').replace('&', '')} size={7} style={{ width: '100%' }} value={from} onChange={(event) => setFrom(event.currentTarget.value)}>{catalog.used.map((font) => <option key={font} value={font}>{font}</option>)}</select></label>
    <label>{t('&With:').replace('&', '')}<select aria-label={t('&With:').replace('&', '')} value={to} onChange={(event) => setTo(event.currentTarget.value)}>{catalog.fonts.map((font) => <option key={font} value={font}>{font}</option>)}</select></label>
    <div className="vibeoffice-dialog__actions"><button className="primary" type="button" disabled={!from || !to} onClick={() => act('apply')}>{t('Replace')}</button><button type="button" onClick={() => act('close')}>{t('Close')}</button></div>
  </OfficeDialog>;
}
