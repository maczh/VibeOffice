import { OfficeDialog } from '@vibeoffice/core';
import { useEffect, useId, useState } from 'react';
import { usePptxRuntime } from './editor-host.js';

interface Fill { t: string; [key: string]: unknown }
interface FillColor { value: string; label: string; color: string }
interface FillAsset { name: string; src: string }
interface FillPattern extends FillAsset { key: string; label: string }
interface FillCatalog { initial: Fill | null; presets: Array<{ name: string; stops: Array<[number, string]> }>; textures: FillAsset[]; patterns: FillPattern[]; colors: FillColor[] }
interface GradientConfig { mode: 'one' | 'two' | 'preset'; style: string; c1: string; c2: string; dark: number; preset: number; tr1: number; tr2: number; fg: string; bg: string }
interface FillVariant { fill: Fill; css: string }
interface FillRuntime {
  dlg?: {
    reactFillEffectsCatalog?(): FillCatalog | null;
    reactFillEffectsBuild?(config: GradientConfig): { variants: FillVariant[]; patterns: FillPattern[] } | null;
    reactFillEffectsTexture?(name: string): Promise<Fill | null>;
    reactFillEffectsImage?(file: File, tile: boolean): Fill | null;
    reactFillEffectsPreview?(fill: Fill): string;
    reactFillEffectsFinish?(fill: Fill | null): boolean;
  };
  bus?: { on(event: string, listener: () => void): () => void };
  t?(message: string): string;
}
const initialConfig = (fill?: Fill | null): GradientConfig => {
  const stops = Array.isArray(fill?.stops) ? fill.stops as Array<{ c?: string; a?: number }> : [];
  return {
    mode: 'two',
    style: fill?.path === 'circle' ? 'center' : fill?.path && fill.path !== 'lin' ? 'corner' : fill?.ang === 0 ? 'vert' : fill?.ang === 45 ? 'ddn' : fill?.ang === 315 ? 'dup' : 'horz',
    c1: stops[0]?.c || '#3366CC', c2: stops.at(-1)?.c || '#99B3E5', dark: .5, preset: 0,
    tr1: 1 - (stops[0]?.a ?? 1), tr2: 1 - (stops.at(-1)?.a ?? 1), fg: '#000000', bg: '#FFFFFF',
  };
};

/** React owns the Fill Effects interaction; rendering and Office fill/media models remain engine-backed. */
export function PptxFillEffectsDialog() {
  const runtime = usePptxRuntime<FillRuntime>();
  const [catalog, setCatalog] = useState<FillCatalog | null>(null);
  const [config, setConfig] = useState<GradientConfig>(() => initialConfig());
  const [variant, setVariant] = useState(0);
  const [tab, setTab] = useState(0);
  const [fill, setFill] = useState<Fill | null>(null);
  const [assetName, setAssetName] = useState('');
  const id = useId().replaceAll(':', '');
  const panelId = `${id}-fill-effects-panel`;
  const tabId = (index: number) => `${id}-fill-effects-tab-${index}`;
  useEffect(() => {
    const refresh = () => {
      const next = runtime?.dlg?.reactFillEffectsCatalog?.() || null;
      setCatalog(next ? { ...next } : null);
      if (next) {
        const start = initialConfig(next.initial);
        setConfig(start);
        setVariant(0);
        setTab(next.initial?.t === 'patt' ? 2 : next.initial?.t === 'img' ? (next.initial.tile ? 1 : 3) : 0);
        setFill(next.initial || null);
        setAssetName('');
      }
    };
    const off = runtime?.bus?.on('pptx-fill-effects-dialog', refresh);
    return () => off?.();
  }, []);
  const gradient = runtime?.dlg?.reactFillEffectsBuild?.(config);
  if (!catalog) return null;
  const t = (message: string) => runtime?.t?.(message) || message;
  const activeGradient = gradient?.variants[Math.max(0, Math.min(3, variant))];
  const update = (key: keyof GradientConfig, value: string | number) => {
    const next = { ...config, [key]: value } as GradientConfig;
    const nextVariant = key === 'style' ? 0 : variant;
    setConfig(next);
    if (key === 'style') setVariant(0);
    const model = runtime?.dlg?.reactFillEffectsBuild?.(next);
    if (tab === 0 && model?.variants[nextVariant]) setFill(model.variants[nextVariant].fill);
    else if (tab === 2 && fill?.t === 'patt' && (key === 'fg' || key === 'bg')) setFill({ ...fill, fg: next.fg, bg: next.bg });
  };
  const colors = catalog.colors;
  const colorControl = (label: string, value: string, key: 'c1' | 'c2' | 'fg' | 'bg') => <label className="col" style={{ minWidth: 150 }}>{t(label)}<span className="row" style={{ alignItems: 'center' }}><select aria-label={t(label)} value={colors.some((color) => color.value === value) ? value : '__custom'} onChange={(event) => { if (event.currentTarget.value !== '__custom') update(key, event.currentTarget.value); }} style={{ minWidth: 110, maxWidth: 180 }}><option value="__custom">{t('Custom color')}</option>{colors.map((color, index) => <option key={`${color.value}-${index}`} value={color.value}>{color.label}</option>)}</select><input aria-label={t(`${label} custom color`)} type="color" value={/^#[0-9a-f]{6}$/i.test(value) ? value : '#3366CC'} onChange={(event) => update(key, event.currentTarget.value.toUpperCase())} /></span></label>;
  const image = (event: React.ChangeEvent<HTMLInputElement>, tile: boolean) => {
    const file = event.currentTarget.files?.[0];
    event.currentTarget.value = '';
    if (!file) return;
    const next = runtime?.dlg?.reactFillEffectsImage?.(file, tile);
    if (next) { setFill(next); setAssetName(file.name); }
  };
  const preview = fill ? runtime?.dlg?.reactFillEffectsPreview?.(fill) || '' : activeGradient?.css || '';
  const tabs = ['Gradient', 'Texture', 'Pattern', 'Picture'];
  return <OfficeDialog title={t('Fill Effects')} className="pptx-fill-effects-dialog" style={{ width: 'min(96vw, 760px)' }} onRequestClose={() => runtime?.dlg?.reactFillEffectsFinish?.(null)}>
    <div role="tablist" aria-label={t('Fill effect types')} className="row" style={{ gap: 4, borderBottom: '1px solid #888' }}>{tabs.map((name, index) => <button key={name} id={tabId(index)} type="button" role="tab" aria-controls={panelId} aria-selected={tab === index} tabIndex={tab === index ? 0 : -1} onClick={() => setTab(index)} onKeyDown={(event) => {
      const next = event.key === 'ArrowRight' ? (index + 1) % tabs.length
        : event.key === 'ArrowLeft' ? (index - 1 + tabs.length) % tabs.length
          : event.key === 'Home' ? 0
            : event.key === 'End' ? tabs.length - 1 : -1;
      if (next < 0) return;
      event.preventDefault();
      setTab(next);
      event.currentTarget.parentElement?.querySelectorAll<HTMLButtonElement>('[role="tab"]')[next]?.focus();
    }}>{t(name)}</button>)}</div>
    <div className="row" style={{ alignItems: 'stretch', gap: 16, minHeight: 310, paddingTop: 10 }}>
      <div id={panelId} aria-labelledby={tabId(tab)} role="tabpanel" className="col" style={{ flex: '1 1 auto', minWidth: 0 }}>
        {tab === 0 && <div className="col" style={{ gap: 8 }}>
          <div className="row" style={{ alignItems: 'start', flexWrap: 'wrap' }}><fieldset className="col"><legend>{t('Colors')}</legend>{(['one', 'two', 'preset'] as const).map((mode) => <label key={mode}><input type="radio" name="pptx-fill-mode" checked={config.mode === mode} onChange={() => update('mode', mode)} />{t(mode === 'one' ? 'One color' : mode === 'two' ? 'Two colors' : 'Preset')}</label>)}</fieldset>
            {config.mode === 'one' && <div className="col">{colorControl('Color 1:', config.c1, 'c1')}<label>{t('Dark')} — {t('Light')}<input aria-label={t('Dark to light')} type="range" min="0" max="100" value={config.dark * 100} onChange={(event) => update('dark', Number(event.currentTarget.value) / 100)} /></label></div>}
            {config.mode === 'two' && <div className="row" style={{ flexWrap: 'wrap' }}>{colorControl('Color 1:', config.c1, 'c1')}{colorControl('Color 2:', config.c2, 'c2')}</div>}
            {config.mode === 'preset' && <label>{t('Preset colors')}<select aria-label={t('Preset colors')} value={config.preset} onChange={(event) => update('preset', Number(event.currentTarget.value))}>{catalog.presets.map((preset, index) => <option key={preset.name} value={index}>{t(preset.name)}</option>)}</select></label>}
          </div>
          <fieldset className="row" style={{ flexWrap: 'wrap' }}><legend>{t('Transparency')}</legend><label>{t('From')}: {Math.round(config.tr1 * 100)}%<input type="range" min="0" max="100" step="5" aria-label={t('From')} value={config.tr1 * 100} onChange={(event) => update('tr1', Number(event.currentTarget.value) / 100)} /></label><label>{t('To')}: {Math.round(config.tr2 * 100)}%<input type="range" min="0" max="100" step="5" aria-label={t('To')} value={config.tr2 * 100} onChange={(event) => update('tr2', Number(event.currentTarget.value) / 100)} /></label></fieldset>
          <div className="row" style={{ alignItems: 'start', flexWrap: 'wrap' }}><label>{t('Shading styles')}<select aria-label={t('Shading styles')} value={config.style} onChange={(event) => update('style', event.currentTarget.value)}>{([['horz', 'Horizontal'], ['vert', 'Vertical'], ['dup', 'Diagonal up'], ['ddn', 'Diagonal down'], ['corner', 'From corner'], ['center', 'From center']] as const satisfies ReadonlyArray<readonly [string, string]>).map(([value, label]) => <option key={value} value={value}>{t(label)}</option>)}</select></label><div className="col"><label>{t('Variants')}</label><div className="row" style={{ gap: 5 }}>{gradient?.variants.map((item, index) => <button key={index} type="button" aria-label={`${t('Variant')} ${index + 1}`} aria-pressed={variant === index} onClick={() => { setVariant(index); setFill(item.fill); }} style={{ width: 46, height: 46, border: '1px solid #888', padding: 0, background: item.css, outline: variant === index ? '2px solid #0b3fa8' : undefined }} />)}</div></div></div>
        </div>}
        {tab === 1 && <div className="col"><label>{t('Texture')}</label><div className="row" style={{ flexWrap: 'wrap', gap: 5 }}>{catalog.textures.map((texture) => <button key={texture.name} type="button" aria-label={texture.name} title={t(texture.name)} onClick={async () => { const next = await runtime?.dlg?.reactFillEffectsTexture?.(texture.name); if (next) { setFill(next); setAssetName(texture.name); } }} style={{ width: 48, height: 48, border: '1px solid #888', backgroundImage: `url(${texture.src})`, backgroundSize: '44px 44px' }} />)}</div><div>{assetName}</div><label>{t('Other Texture...')}<input type="file" accept="image/*" onChange={(event) => image(event, true)} /></label></div>}
        {tab === 2 && <div className="col"><label>{t('Pattern')}</label><div className="row" style={{ flexWrap: 'wrap', gap: 4, maxHeight: 220, overflow: 'auto' }}>{(gradient?.patterns || catalog.patterns).map((pattern) => <button key={pattern.key} type="button" aria-label={pattern.label} title={t(pattern.label)} aria-pressed={(fill?.prst as string) === pattern.key} onClick={() => setFill({ t: 'patt', prst: pattern.key, fg: config.fg, bg: config.bg })} style={{ width: 38, height: 34, border: '1px solid #888', backgroundImage: pattern.src }} />)}</div><div className="row" style={{ flexWrap: 'wrap' }}>{colorControl('Foreground:', config.fg, 'fg')}{colorControl('Background:', config.bg, 'bg')}</div></div>}
        {tab === 3 && <div className="col"><label>{t('Select Picture...')}<input type="file" accept="image/*" onChange={(event) => image(event, false)} /></label><div>{assetName}</div></div>}
      </div>
      <aside className="col" style={{ flex: '0 0 130px' }}><label>{t('Sample')}</label><div aria-label={t('Fill preview')} className="fill-effects-preview" style={{ width: 120, height: 120, border: '1px solid #777', background: preview }} /></aside>
    </div>
    <div className="vibeoffice-dialog__actions"><button type="button" className="primary" onClick={() => runtime?.dlg?.reactFillEffectsFinish?.(fill)}>{t('OK')}</button><button type="button" onClick={() => runtime?.dlg?.reactFillEffectsFinish?.(null)}>{t('Cancel')}</button></div>
  </OfficeDialog>;
}
