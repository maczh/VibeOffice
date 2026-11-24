import { OfficeDialog } from '@vibeoffice/core';
import { useEffect, useId, useState } from 'react';
import { usePptxRuntime } from './editor-host.js';

interface EffectOptionsCatalog {
  title: string;
  directions: Array<[string, string]>;
  dir: string;
  amount: number;
  amountEnabled: boolean;
  unit: string;
  start: string;
  delay: number;
  speeds: Array<{ name: string; value: number; label: string }>;
  dur: number;
  repeat: number;
  by: string;
  hasText: boolean;
}
interface EffectOptionsRuntime {
  dlg?: { reactEffectOptionsCatalog?(): EffectOptionsCatalog | null; reactEffectOptionsAction?(action: string, values?: Record<string, string | number>): boolean };
  bus?: { on(event: string, listener: () => void): () => void };
  t?(message: string): string;
}
const startOptions: Array<[string, string]> = [['click', 'On Click'], ['with', 'With Previous'], ['after', 'After Previous']];
const repeatOptions: Array<[number, string]> = [[1, '(none)'], [2, '2'], [3, '3'], [4, '4'], [5, '5'], [10, '10']];

/** React owns the three-tab animation options form; the presentation callback commits it. */
export function PptxEffectOptionsDialog() {
  const runtime = usePptxRuntime<EffectOptionsRuntime>();
  const [catalog, setCatalog] = useState<EffectOptionsCatalog | null>(null);
  const [tab, setTab] = useState(0);
  const [dir, setDir] = useState('');
  const [amount, setAmount] = useState(150);
  const [start, setStart] = useState('click');
  const [delay, setDelay] = useState(0);
  const [dur, setDur] = useState(500);
  const [repeat, setRepeat] = useState(1);
  const [by, setBy] = useState('all');
  const id = useId().replaceAll(':', '');
  const panelId = `${id}-animation-options-panel`;
  const tabId = (index: number) => `${id}-animation-options-tab-${index}`;
  useEffect(() => {
    const refresh = () => {
      const next = runtime?.dlg?.reactEffectOptionsCatalog?.() || null;
      setCatalog(next ? { ...next, directions: [...next.directions], speeds: [...next.speeds] } : null);
      if (next) { setTab(0); setDir(next.dir); setAmount(next.amount || 150); setStart(next.start); setDelay(next.delay); setDur(next.dur); setRepeat(next.repeat); setBy(next.by); }
    };
    const off = runtime?.bus?.on('pptx-effect-options-dialog', refresh);
    return () => off?.();
  }, []);
  if (!catalog) return null;
  const t = (message: string) => runtime?.t?.(message) || message;
  const close = () => runtime?.dlg?.reactEffectOptionsAction?.('close');
  const apply = () => runtime?.dlg?.reactEffectOptionsAction?.('apply', { dir, amount, start, delay, dur, repeat, by });
  const tabs = ['Effect', 'Timing', 'Text Animation'];
  return <OfficeDialog title={catalog.title} className="pptx-effect-options-dialog" style={{ width: 'min(96vw, 420px)' }} onRequestClose={close}>
    <div role="tablist" aria-label={t('Animation options')} className="row">
      {tabs.map((label, index) => <button key={label} id={tabId(index)} type="button" role="tab" aria-controls={panelId} aria-selected={tab === index} tabIndex={tab === index ? 0 : -1} onClick={() => setTab(index)} onKeyDown={(event) => {
        const next = event.key === 'ArrowRight' ? (index + 1) % tabs.length
          : event.key === 'ArrowLeft' ? (index - 1 + tabs.length) % tabs.length
            : event.key === 'Home' ? 0
              : event.key === 'End' ? tabs.length - 1 : -1;
        if (next < 0) return;
        event.preventDefault();
        setTab(next);
        event.currentTarget.parentElement?.querySelectorAll<HTMLButtonElement>('[role="tab"]')[next]?.focus();
      }}>{t(label)}</button>)}
    </div>
    <section id={panelId} aria-labelledby={tabId(tab)} role="tabpanel" className="col" style={{ minHeight: 132, paddingTop: 12 }}>
      {tab === 0 && <fieldset><legend>{t('Settings')}</legend>{catalog.directions.length ? <label>{t('&Direction:').replace('&', '')}<select aria-label={t('Direction')} value={dir} onChange={(event) => setDir(event.currentTarget.value)}>{catalog.directions.map(([value, label]) => <option key={value} value={value}>{label}</option>)}</select></label> : <div className="tp-note">{t('This effect has no direction settings.')}</div>}{catalog.amountEnabled && <label>{t(catalog.unit === '°' ? '&Amount:' : '&Size:').replace('&', '')}<input aria-label={t(catalog.unit === '°' ? 'Amount' : 'Size')} type="number" min={10} max={1440} step={10} value={amount} onChange={(event) => setAmount(Number(event.currentTarget.value))} /><span>{catalog.unit}</span></label>}</fieldset>}
      {tab === 1 && <fieldset><legend>{t('Timing')}</legend><label>{t('&Start:').replace('&', '')}<select aria-label={t('Start')} value={start} onChange={(event) => setStart(event.currentTarget.value)}>{startOptions.map(([value, label]) => <option key={value} value={value}>{t(label)}</option>)}</select></label><label>{t('&Delay:').replace('&', '')}<input aria-label={t('Delay')} type="number" min={0} max={60} step={0.5} value={delay / 1000} onChange={(event) => setDelay(Number(event.currentTarget.value) * 1000)} />{t('seconds')}</label><label>{t('Sp&eed:').replace('&', '')}<select aria-label={t('Speed')} value={dur} onChange={(event) => setDur(Number(event.currentTarget.value))}>{catalog.speeds.map((speed) => <option key={speed.value} value={speed.value}>{speed.label}</option>)}</select></label><label>{t('&Repeat:').replace('&', '')}<select aria-label={t('Repeat')} value={repeat} onChange={(event) => setRepeat(Number(event.currentTarget.value))}>{repeatOptions.map(([value, label]) => <option key={value} value={value}>{t(label)}</option>)}</select></label></fieldset>}
      {tab === 2 && <fieldset><legend>{t('Text Animation')}</legend><label>{t('&Group text:').replace('&', '')}<select aria-label={t('Group text')} value={by} disabled={!catalog.hasText} onChange={(event) => setBy(event.currentTarget.value)}><option value="all">{t('As One Object')}</option><option value="para">{t('By 1st Level Paragraphs')}</option></select></label></fieldset>}
    </section>
    <div className="vibeoffice-dialog__actions"><button className="primary" type="button" onClick={apply}>{t('OK')}</button><button type="button" onClick={close}>{t('Cancel')}</button></div>
  </OfficeDialog>;
}
