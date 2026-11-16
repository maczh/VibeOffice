import { OfficeDialog } from '@vibeoffice/core';
import { useEffect, useId, useRef, useState } from 'react';

type BulletKind = 'char' | 'num';
interface BulletChoice { value: string | null; preview: string[]; index: number }
interface BulletsCatalog { kind: BulletKind; pick: string | null; customChar: string; size: number; color: string | null; start: number; tab: number; bullets: BulletChoice[]; numbers: BulletChoice[] }
interface BulletsApply { kind: BulletKind; pick: string | null; customChar: string; size: number; color: string | null; start: number }
interface BulletsAction extends BulletsApply { value?: string | null }
interface BulletsRuntime {
  dlg?: { reactBulletsCatalog?(): BulletsCatalog | null; reactBulletsOpenColor?(anchor: HTMLElement): boolean; reactBulletsCustomize?(): boolean; reactBulletsAction?(action: string, values?: BulletsAction): boolean; reactFontResolveColor?(color: string): string };
  bus?: { on(event: string, listener: () => void): () => void };
  t?(message: string): string;
}
function runtime(): BulletsRuntime | undefined { return (window as Window & { L?: BulletsRuntime }).L; }

/** React-owned bullet and numbering choices; paragraph list formatting stays in the text engine. */
export function PptxBulletsDialog() {
  const [catalog, setCatalog] = useState<BulletsCatalog | null>(null);
  const id = useId().replaceAll(':', '');
  const tabId = (index: number) => `${id}-bullets-tab-${index}`;
  const panelId = (index: number) => `${id}-bullets-panel-${index}`;
  const bulletColor = useRef<HTMLButtonElement>(null);
  const numberColor = useRef<HTMLButtonElement>(null);
  useEffect(() => {
    const refresh = () => { const next = runtime()?.dlg?.reactBulletsCatalog?.() || null; setCatalog(next ? { ...next, bullets: [...next.bullets], numbers: [...next.numbers] } : null); };
    const off = runtime()?.bus?.on('pptx-bullets-dialog', refresh);
    return () => off?.();
  }, []);
  if (!catalog) return null;
  const t = (message: string) => runtime()?.t?.(message) || message;
  const update = (changes: Partial<BulletsCatalog>) => setCatalog({ ...catalog, ...changes });
  const select = (kind: BulletKind, value: string | null) => {
    runtime()?.dlg?.reactBulletsAction?.('select', { kind, value, pick: value, customChar: catalog.customChar, size: catalog.size, color: catalog.color, start: catalog.start });
    update({ kind, pick: value, customChar: kind === 'char' && value ? value : catalog.customChar });
  };
  const color = catalog.color ? (catalog.color.startsWith('#') ? catalog.color : runtime()?.dlg?.reactFontResolveColor?.(catalog.color) || catalog.color) : '#000000';
  const close = () => runtime()?.dlg?.reactBulletsAction?.('close');
  const apply = () => runtime()?.dlg?.reactBulletsAction?.('apply', { kind: catalog.kind, pick: catalog.pick, customChar: catalog.customChar, size: catalog.size, color: catalog.color, start: catalog.start });
  const renderChoices = (kind: BulletKind, choices: BulletChoice[]) => <div className={kind === 'num' ? 'num-grid' : 'bul-grid'}>
    {choices.map((choice) => <button key={`${kind}-${choice.index}`} type="button" className={catalog.kind === kind && catalog.pick === choice.value ? 'on' : ''} aria-label={choice.value || t('None')} onClick={() => select(kind, choice.value)} onDoubleClick={apply}>
      {choice.preview.length === 1 ? <span style={{ fontSize: 13, alignSelf: 'center' }}>{t('None')}</span> : choice.preview.map((preview, index) => <div key={`${preview}:${index}`}><span style={{ minWidth: 16 }}>{preview}</span><i /></div>)}
    </button>)}
  </div>;
  const colorButton = (kind: BulletKind) => <button ref={kind === 'char' ? bulletColor : numberColor} type="button" className="btn small" aria-label={t('Color')} onClick={() => { const anchor = kind === 'char' ? bulletColor.current : numberColor.current; if (anchor) runtime()?.dlg?.reactBulletsOpenColor?.(anchor); }}><span style={{ display: 'inline-block', width: 60, height: 14, background: color, border: '1px solid #555', verticalAlign: 'middle' }} /></button>;
  return <OfficeDialog title={t('Bullets and Numbering')} className="pptx-bullets-dialog" style={{ width: 'min(94vw, 560px)' }} onRequestClose={close}>
    <div role="tablist" aria-label={t('List type')} className="tabs-head">{(['Bulleted', 'Numbered'] as const).map((label, index) => <button key={label} type="button" id={tabId(index)} role="tab" aria-controls={panelId(index)} aria-selected={catalog.tab === index} tabIndex={catalog.tab === index ? 0 : -1} onClick={() => update({ tab: index })} onKeyDown={(event) => {
      const next = event.key === 'ArrowRight' ? (index + 1) % 2
        : event.key === 'ArrowLeft' ? (index - 1 + 2) % 2
          : event.key === 'Home' ? 0
            : event.key === 'End' ? 1 : -1;
      if (next < 0) return;
      event.preventDefault();
      update({ tab: next });
      event.currentTarget.parentElement?.querySelectorAll<HTMLButtonElement>('[role="tab"]')[next]?.focus();
    }}>{t(label === 'Bulleted' ? 'Bulle&ted' : '&Numbered').replace('&', '')}</button>)}</div>
    <section id={panelId(0)} aria-labelledby={tabId(0)} role="tabpanel" hidden={catalog.tab !== 0}>{renderChoices('char', catalog.bullets)}<div className="row" style={{ alignItems: 'center', gap: 8, marginTop: 8 }}><label>{t('Size:')} <input type="number" min="25" max="400" step="5" value={catalog.size} onChange={(event) => update({ size: Number(event.currentTarget.value) })} /></label><span>{t('% of text')}</span><label>{t('Color:')} {colorButton('char')}</label><button type="button" className="btn small" onClick={() => runtime()?.dlg?.reactBulletsCustomize?.()}>{t('C&ustomize...').replace('&', '')}</button></div></section>
    <section id={panelId(1)} aria-labelledby={tabId(1)} role="tabpanel" hidden={catalog.tab !== 1}>{renderChoices('num', catalog.numbers)}<div className="row" style={{ alignItems: 'center', gap: 8, marginTop: 8 }}><label>{t('Size:')} <input type="number" min="25" max="400" step="5" value={catalog.size} onChange={(event) => update({ size: Number(event.currentTarget.value) })} /></label><span>{t('% of text')}</span><label>{t('Color:')} {colorButton('num')}</label><label>{t('Start at:')} <input type="number" min="1" max="9999" step="1" value={catalog.start} onChange={(event) => update({ start: Number(event.currentTarget.value) })} /></label></div></section>
    <div className="vibeoffice-dialog__actions"><button className="primary" type="button" onClick={apply}>{t('OK')}</button><button type="button" onClick={close}>{t('Cancel')}</button></div>
  </OfficeDialog>;
}
