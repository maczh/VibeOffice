import { OfficeDialog } from '@vibeoffice/core';
import { useEffect, useRef, useState } from 'react';
import './table-format-dialog.css';

type BorderKind = 'all' | 'out' | 'in' | 'none' | 't' | 'b' | 'l' | 'r';
interface TableFormatCatalog {
  color: string;
  colorHex: string;
  dash: string;
  dashes: Array<{ value: string; label: string }>;
  width: number;
  widths: number[];
  borderChoices: Array<{ kind: BorderKind; icon: string; label: string }>;
}
interface TableFormatRuntime {
  app?: {
    reactTableFormatCatalog?(): TableFormatCatalog | null;
    reactTableFormatClose?(): void;
    reactTableFormatSet?(key: string, value: string | number): boolean;
    reactTableFormatBorder?(kind: BorderKind): boolean;
    reactTableFormatColorMenu?(kind: 'border' | 'fill', anchor: HTMLElement): boolean;
  };
  icons?: { get(name: string): string };
  bus?: { on(event: string, listener: () => void): () => void };
  t?(message: string, ...args: unknown[]): string;
}
function runtime(): TableFormatRuntime | undefined { return (window as Window & { L?: TableFormatRuntime }).L; }

/** React owns the table-format surface; border, fill and alignment edits remain engine transactions. */
export function PptxTableFormatDialog() {
  const [catalog, setCatalog] = useState<TableFormatCatalog | null>(null);
  const [tab, setTab] = useState<'borders' | 'fill' | 'text'>('borders');
  const wasOpen = useRef(false);
  useEffect(() => {
    const refresh = () => {
      const next = runtime()?.app?.reactTableFormatCatalog?.() || null;
      if (next && !wasOpen.current) setTab('borders');
      wasOpen.current = !!next;
      setCatalog(next ? { ...next, dashes: [...next.dashes], widths: [...next.widths], borderChoices: [...next.borderChoices] } : null);
    };
    const off = runtime()?.bus?.on('pptx-table-format-dialog', refresh);
    return () => off?.();
  }, []);
  if (!catalog) return null;
  const t = (message: string, ...args: unknown[]) => runtime()?.t?.(message, ...args) || message;
  const chooseColor = (kind: 'border' | 'fill', anchor: HTMLButtonElement) => runtime()?.app?.reactTableFormatColorMenu?.(kind, anchor);
  return <OfficeDialog title={t('Format Table')} className="pptx-table-format-dialog" style={{ width: 'min(94vw, 420px)' }} onRequestClose={() => runtime()?.app?.reactTableFormatClose?.()}>
    <div className="pptx-table-format__tabs" role="tablist" aria-label={t('Format Table')}>
      {([['borders', 'Borders'], ['fill', 'Fill'], ['text', 'Text Box']] as const).map(([id, label]) => <button key={id} type="button" role="tab" aria-selected={tab === id} onClick={() => setTab(id)}>{t(label)}</button>)}
    </div>
    {tab === 'borders' && <section className="pptx-table-format__body" role="tabpanel" aria-label={t('Borders')}>
      <label>{t('&Style:').replace('&', '')}<select aria-label={t('Border style')} value={catalog.dash} onChange={(event) => runtime()?.app?.reactTableFormatSet?.('dash', event.currentTarget.value)}>{catalog.dashes.map((dash) => <option key={dash.value} value={dash.value}>{dash.label}</option>)}</select></label>
      <label>{t('&Color:').replace('&', '')}<button type="button" className="btn small" aria-label={t('Border color')} onClick={(event) => chooseColor('border', event.currentTarget)}><span aria-hidden="true" style={{ display: 'inline-block', width: 64, height: 14, border: '1px solid #555', background: catalog.colorHex }} /> <span translate="no">{catalog.color}</span></button></label>
      <label>{t('&Width:').replace('&', '')}<select aria-label={t('Border width')} value={catalog.width} onChange={(event) => runtime()?.app?.reactTableFormatSet?.('width', Number(event.currentTarget.value))}>{catalog.widths.map((width) => <option key={width} value={width}>{t('{0} pt', width)}</option>)}</select></label>
      <span>{t('Click a button to apply borders:')}</span>
      <div className="pptx-table-format__borders" role="group" aria-label={t('Apply borders')}>
        {catalog.borderChoices.map((choice) => <button key={choice.kind} type="button" className="btn small" aria-label={choice.label} title={choice.label} dangerouslySetInnerHTML={{ __html: runtime()?.icons?.get(choice.icon) || '' }} onClick={() => runtime()?.app?.reactTableFormatBorder?.(choice.kind)} />)}
      </div>
    </section>}
    {tab === 'fill' && <section className="pptx-table-format__body" role="tabpanel" aria-label={t('Fill')}>
      <label>{t('Fill &color:').replace('&', '')}<button type="button" className="btn small" aria-label={t('Fill color')} onClick={(event) => chooseColor('fill', event.currentTarget)}><span aria-hidden="true" style={{ display: 'inline-block', width: 70, height: 14, border: '1px solid #555', background: 'repeating-linear-gradient(45deg,#fff 0 4px,#eee 4px 8px)' }} /> <span>{t('Choose fill')}</span></button></label>
    </section>}
    {tab === 'text' && <section className="pptx-table-format__body" role="tabpanel" aria-label={t('Text Box')}>
      <label>{t('Text &alignment:').replace('&', '')}<select aria-label={t('Text alignment')} defaultValue="t" onChange={(event) => runtime()?.app?.reactTableFormatSet?.('anchor', event.currentTarget.value)}><option value="t">{t('Top')}</option><option value="ctr">{t('Middle')}</option><option value="b">{t('Bottom')}</option></select></label>
    </section>}
    <div className="vibeoffice-dialog__actions"><button type="button" className="primary" onClick={() => runtime()?.app?.reactTableFormatClose?.()}>{t('OK')}</button></div>
  </OfficeDialog>;
}
