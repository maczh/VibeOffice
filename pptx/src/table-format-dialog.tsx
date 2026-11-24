import { OfficeDialog } from '@vibeoffice/core';
import { useEffect, useId, useRef, useState } from 'react';
import { usePptxRuntime } from './editor-host.js';
import { PptxPreviewArtwork } from './slides-pane.js';
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

/** React owns the table-format surface; border, fill and alignment edits remain engine transactions. */
export function PptxTableFormatDialog() {
  const runtime = usePptxRuntime<TableFormatRuntime>();
  const [catalog, setCatalog] = useState<TableFormatCatalog | null>(null);
  const [tab, setTab] = useState<'borders' | 'fill' | 'text'>('borders');
  const id = useId().replaceAll(':', '');
  const tabId = (key: 'borders' | 'fill' | 'text') => `${id}-table-format-tab-${key}`;
  const panelId = (key: 'borders' | 'fill' | 'text') => `${id}-table-format-panel-${key}`;
  const wasOpen = useRef(false);
  useEffect(() => {
    const refresh = () => {
      const next = runtime?.app?.reactTableFormatCatalog?.() || null;
      if (next && !wasOpen.current) setTab('borders');
      wasOpen.current = !!next;
      setCatalog(next ? { ...next, dashes: [...next.dashes], widths: [...next.widths], borderChoices: [...next.borderChoices] } : null);
    };
    const off = runtime?.bus?.on('pptx-table-format-dialog', refresh);
    return () => off?.();
  }, []);
  if (!catalog) return null;
  const t = (message: string, ...args: unknown[]) => runtime?.t?.(message, ...args) || message;
  const chooseColor = (kind: 'border' | 'fill', anchor: HTMLButtonElement) => runtime?.app?.reactTableFormatColorMenu?.(kind, anchor);
  return <OfficeDialog title={t('Format Table')} className="pptx-table-format-dialog" style={{ width: 'min(94vw, 420px)' }} onRequestClose={() => runtime?.app?.reactTableFormatClose?.()}>
    <div className="pptx-table-format__tabs" role="tablist" aria-label={t('Format Table')}>
      {([['borders', 'Borders'], ['fill', 'Fill'], ['text', 'Text Box']] as const).map(([key, label], index) => <button key={key} id={tabId(key)} type="button" role="tab" aria-controls={panelId(key)} aria-selected={tab === key} tabIndex={tab === key ? 0 : -1} onClick={() => setTab(key)} onKeyDown={(event) => {
        const next = event.key === 'ArrowRight' ? (index + 1) % 3
          : event.key === 'ArrowLeft' ? (index - 1 + 3) % 3
            : event.key === 'Home' ? 0
              : event.key === 'End' ? 2 : -1;
        if (next < 0) return;
        event.preventDefault();
        const nextTab = (['borders', 'fill', 'text'] as const)[next];
        if (nextTab) setTab(nextTab);
        event.currentTarget.parentElement?.querySelectorAll<HTMLButtonElement>('[role="tab"]')[next]?.focus();
      }}>{t(label)}</button>)}
    </div>
    <section id={panelId('borders')} aria-labelledby={tabId('borders')} className="pptx-table-format__body" role="tabpanel" hidden={tab !== 'borders'}>
      <label>{t('&Style:').replace('&', '')}<select aria-label={t('Border style')} value={catalog.dash} onChange={(event) => runtime?.app?.reactTableFormatSet?.('dash', event.currentTarget.value)}>{catalog.dashes.map((dash) => <option key={dash.value} value={dash.value}>{dash.label}</option>)}</select></label>
      <label>{t('&Color:').replace('&', '')}<button type="button" className="btn small" aria-label={t('Border color')} onClick={(event) => chooseColor('border', event.currentTarget)}><span aria-hidden="true" style={{ display: 'inline-block', width: 64, height: 14, border: '1px solid #555', background: catalog.colorHex }} /> <span translate="no">{catalog.color}</span></button></label>
      <label>{t('&Width:').replace('&', '')}<select aria-label={t('Border width')} value={catalog.width} onChange={(event) => runtime?.app?.reactTableFormatSet?.('width', Number(event.currentTarget.value))}>{catalog.widths.map((width) => <option key={width} value={width}>{t('{0} pt', width)}</option>)}</select></label>
      <span>{t('Click a button to apply borders:')}</span>
      <div className="pptx-table-format__borders" role="group" aria-label={t('Apply borders')}>
        {catalog.borderChoices.map((choice) => <button key={choice.kind} type="button" className="btn small" aria-label={choice.label} title={choice.label} onClick={() => runtime?.app?.reactTableFormatBorder?.(choice.kind)}><PptxPreviewArtwork markup={runtime?.icons?.get(choice.icon) || ''} /></button>)}
      </div>
    </section>
    <section id={panelId('fill')} aria-labelledby={tabId('fill')} className="pptx-table-format__body" role="tabpanel" hidden={tab !== 'fill'}>
      <label>{t('Fill &color:').replace('&', '')}<button type="button" className="btn small" aria-label={t('Fill color')} onClick={(event) => chooseColor('fill', event.currentTarget)}><span aria-hidden="true" style={{ display: 'inline-block', width: 70, height: 14, border: '1px solid #555', background: 'repeating-linear-gradient(45deg,#fff 0 4px,#eee 4px 8px)' }} /> <span>{t('Choose fill')}</span></button></label>
    </section>
    <section id={panelId('text')} aria-labelledby={tabId('text')} className="pptx-table-format__body" role="tabpanel" hidden={tab !== 'text'}>
      <label>{t('Text &alignment:').replace('&', '')}<select aria-label={t('Text alignment')} defaultValue="t" onChange={(event) => runtime?.app?.reactTableFormatSet?.('anchor', event.currentTarget.value)}><option value="t">{t('Top')}</option><option value="ctr">{t('Middle')}</option><option value="b">{t('Bottom')}</option></select></label>
    </section>
    <div className="vibeoffice-dialog__actions"><button type="button" className="primary" onClick={() => runtime?.app?.reactTableFormatClose?.()}>{t('OK')}</button></div>
  </OfficeDialog>;
}
