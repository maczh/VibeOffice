import { OfficeDialog } from '@vibeoffice/core';
import { useEffect, useId, useLayoutEffect, useRef, useState, type CSSProperties, type KeyboardEvent } from 'react';
import { usePptxRuntime } from './editor-host.js';

interface Fill { t: string; c?: string; a?: number; [key: string]: unknown }
interface ColorOption { value: string; label: string; color: string }
interface FormatCatalog {
  title: string; isPic: boolean; isLine: boolean; isWA: boolean; hasText: boolean; showArrows: boolean; canRotate: boolean;
  fill: Fill; line: Fill; colors: ColorOption[]; width: number; height: number; x: number; y: number; rotation: number; lockAspect: boolean;
  crop: { l: number; r: number; t: number; b: number }; cropFullWidthPt: number; cropFullHeightPt: number;
  imageMode: string; brightness: number; contrast: number; margins: number[]; anchor: string; wrap: boolean; autofit: boolean; verticalText: boolean; alt: string;
  lineWidth: number; dash: string; compound: string; head: string; tail: string; headSize: string; tailSize: string; alpha: number;
  original?: { width: number; height: number };
}
interface Runtime {
  dlg?: {
    reactFormatShapeCatalog?(): FormatCatalog | null;
    reactFormatShapeAction?(action: string, changes?: Record<string, unknown>): boolean;
    reactFormatShapePreview?(fill: Fill): string;
    reactFormatShapeEffects?(fill: Fill): Promise<Fill | null>;
  };
  bus?: { on(event: string, listener: () => void): () => void };
  t?(message: string, ...args: unknown[]): string;
}
function ColorFill({ label, fill, colors, line = false, onChange }: { label: string; fill: Fill; colors: ColorOption[]; line?: boolean; onChange(value: Fill): void }) {
  const runtime = usePptxRuntime<Runtime>();
  const [open, setOpen] = useState(false);
  const [activeIndex, setActiveIndex] = useState(0);
  const trigger = useRef<HTMLButtonElement>(null);
  const itemRefs = useRef<Array<HTMLButtonElement | null>>([]);
  const id = `${useId().replaceAll(':', '')}-colors`;
  const t = (message: string, ...args: unknown[]) => runtime?.t?.(message, ...args) || message;
  const swatch = fill.t === 'none' ? 'repeating-linear-gradient(45deg,#fff 0 4px,#eee 4px 8px)' : line ? fill.c || '#111111' : runtime?.dlg?.reactFormatShapePreview?.(fill) || fill.c || '#FFFFFF';
  const colorStart = line ? 1 : 0;
  const lastItem = colorStart + colors.length - 1 + (line ? 0 : 1);
  useLayoutEffect(() => {
    if (!open) return;
    setActiveIndex(0);
    itemRefs.current[0]?.focus({ preventScroll: true });
  }, [open]);
  const close = () => { setOpen(false); trigger.current?.focus(); };
  const onKeyDown = (event: KeyboardEvent<HTMLDivElement>) => {
    if (event.key === 'Escape') { event.preventDefault(); event.stopPropagation(); close(); return; }
    const target = event.target;
    if (!(target instanceof HTMLButtonElement)) return;
    const index = itemRefs.current.indexOf(target);
    if (index < 0) return;
    let next = index;
    if (event.key === 'ArrowLeft') next = Math.max(0, index - 1);
    else if (event.key === 'ArrowRight') next = Math.min(lastItem, index + 1);
    else if (event.key === 'ArrowUp') next = Math.max(0, index - 8);
    else if (event.key === 'ArrowDown') next = Math.min(lastItem, index + 8);
    else if (event.key === 'Home') next = 0;
    else if (event.key === 'End') next = lastItem;
    else return;
    event.preventDefault(); event.stopPropagation();
    setActiveIndex(next);
    itemRefs.current[next]?.focus();
  };
  const choose = (value: Fill) => { onChange(value); close(); };
  return <div className="col" style={{ position: 'relative' }}><label>{t(label)}</label><button ref={trigger} type="button" className="btn small" aria-label={t(label)} aria-haspopup="menu" aria-expanded={open} aria-controls={open ? id : undefined} onClick={() => setOpen((value) => !value)}><span aria-hidden="true" style={{ display: 'inline-block', width: 70, height: 14, border: '1px solid var(--field-border)', background: swatch }} /><span className="dd-arrow" /></button>{open && <div id={id} className="vibeoffice-format-shape-colors" data-menu-keyboard-owner="react" role="menu" aria-label={t(`${label} colors`)} onKeyDown={onKeyDown} style={{ position: 'absolute', zIndex: 2, top: '100%', left: 0, display: 'grid', gridTemplateColumns: 'repeat(8,24px)', gap: 3, padding: 6, border: '1px solid var(--menu-border)', background: 'var(--surface)', width: 'max-content' }}>{line && <button ref={(element) => { itemRefs.current[0] = element; }} type="button" role="menuitem" className="cp-wide" tabIndex={activeIndex === 0 ? 0 : -1} style={{ gridColumn: '1 / -1' }} onFocus={() => setActiveIndex(0)} onClick={() => choose({ t: 'none' })}>{t('No Line')}</button>}{colors.map((color, index) => { const itemIndex = colorStart + index; return <button ref={(element) => { itemRefs.current[itemIndex] = element; }} key={`${color.value}-${index}`} type="button" role="menuitem" aria-label={color.label} title={color.label} tabIndex={activeIndex === itemIndex ? 0 : -1} style={{ width: 24, height: 22, padding: 0, background: color.color }} onFocus={() => setActiveIndex(itemIndex)} onClick={() => choose({ t: 'solid', c: color.value, a: 1 })} />; })}<label className="cp-wide" style={{ gridColumn: '1 / -1' }}>{t('More Fill Colors...')}<input type="color" aria-label={t(`${label} custom color`)} value={/^#[0-9a-f]{6}$/i.test(fill.c || '') ? fill.c : '#3366CC'} onChange={(event) => choose({ t: 'solid', c: event.currentTarget.value.toUpperCase(), a: 1 })} /></label>{!line && <button ref={(element) => { itemRefs.current[lastItem] = element; }} type="button" role="menuitem" className="cp-wide" tabIndex={activeIndex === lastItem ? 0 : -1} style={{ gridColumn: '1 / -1' }} onFocus={() => setActiveIndex(lastItem)} onClick={async () => { const selected = await runtime?.dlg?.reactFormatShapeEffects?.(fill); if (selected) onChange(selected); close(); }}>{t('Fill Effects...')}</button>}</div>}</div>;
}

/** React-owned Format Shape/Picture controls; the legacy model transaction remains authoritative. */
export function PptxFormatShapeDialog({ initialTab }: { initialTab?: string }) {
  const runtime = usePptxRuntime<Runtime>();
  const [catalog, setCatalog] = useState<FormatCatalog | null>(null);
  const [changes, setChanges] = useState<Record<string, unknown>>({});
  const [tab, setTab] = useState('');
  const id = useId().replaceAll(':', '');
  const panelId = `${id}-format-shape-panel`;
  useEffect(() => {
    const refresh = () => {
      const next = runtime?.dlg?.reactFormatShapeCatalog?.() || null;
      setCatalog(next ? { ...next, crop: { ...next.crop }, margins: [...next.margins], colors: [...next.colors] } : null);
      setChanges({});
      if (next) setTab(initialTab && ['Colors and Lines', 'Size', 'Position', 'Picture', 'Text Box', 'Web'].includes(initialTab) ? initialTab : 'Colors and Lines');
    };
    const off = runtime?.bus?.on('pptx-format-shape-dialog', refresh);
    return () => off?.();
  }, [initialTab]);
  if (!catalog) return null;
  const t = (message: string, ...args: unknown[]) => runtime?.t?.(message, ...args) || message;
  const update = (key: string, value: unknown) => setChanges((current) => ({ ...current, [key]: value }));
  const value = <T,>(key: string, original: T): T => Object.hasOwn(changes, key) ? changes[key] as T : original;
  const number = (label: string, key: string, original: number, min: number, max: number, step = .01, unit = '') => <label className="col">{t(label)}<span className="row" style={{ alignItems: 'center' }}><input aria-label={t(label)} type="number" min={min} max={max} step={step} value={value(key, original)} onChange={(event) => update(key, Number(event.currentTarget.value))} />{unit}</span></label>;
  const line = value<Fill>('line', catalog.line);
  const fill = value<Fill>('fill', catalog.fill);
  const tabs = ['Colors and Lines', 'Size', 'Position', ...(catalog.isPic ? ['Picture'] : []), ...(catalog.hasText ? ['Text Box'] : []), 'Web'];
  const dashOptions = [['solid','Solid'],['sysDot','Round Dot'],['sysDash','Square Dot'],['dash','Dash'],['dashDot','Dash Dot'],['lgDash','Long Dash'],['lgDashDot','Long Dash Dot'],['lgDashDotDot','Long Dash Dot Dot']] as const;
  const compoundOptions = [['sng','Single'],['dbl','Double'],['thickThin','Thick Thin'],['thinThick','Thin Thick'],['tri','Triple']] as const;
  const arrowItems = [['head','Begin style:',catalog.head],['headSz','Begin size:',catalog.headSize],['tail','End style:',catalog.tail],['tailSz','End size:',catalog.tailSize]] as const;
  const arrowTypes = [['none','None'],['triangle','Arrow'],['arrow','Open Arrow'],['stealth','Stealth Arrow'],['diamond','Diamond Arrow'],['oval','Oval Arrow']] as const;
  const arrowSizes = [['sm','Small'],['med','Medium'],['lg','Large']] as const;
  const imageModes = [['','Automatic'],['gray','Grayscale'],['bw','Black & White'],['wash','Washout']] as const;
  const anchors = [['t','Top'],['ctr','Middle'],['b','Bottom'],['t|c','Top Centered'],['ctr|c','Middle Centered'],['b|c','Bottom Centered']] as const;
  const sizeNumber = (label: string, key: 'w' | 'h', original: number) => <label className="col">{t(label)}<span className="row" style={{ alignItems: 'center' }}><input aria-label={t(label)} type="number" min="0" max="56" step=".01" value={value(key, original)} onChange={(event) => {
    const nextValue = Number(event.currentTarget.value), locked = value('lockAspect', catalog.lockAspect), currentWidth = Number(value('w', catalog.width)), currentHeight = Number(value('h', catalog.height));
    setChanges((current) => ({ ...current, [key]: nextValue, ...(locked && key === 'h' && currentHeight ? { w: currentWidth * nextValue / currentHeight } : {}), ...(locked && key === 'w' && currentWidth ? { h: currentHeight * nextValue / currentWidth } : {}) }));
  }} />in</span></label>;
  const style: CSSProperties = { minWidth: 0 };
  const commit = () => runtime?.dlg?.reactFormatShapeAction?.('apply', changes);
  return <OfficeDialog title={t(catalog.title)} className="pptx-format-shape-dialog" style={{ width: 'min(96vw, 560px)' }} onRequestClose={() => runtime?.dlg?.reactFormatShapeAction?.('close')}>
    <div role="tablist" aria-label={t('Format options')} className="row" style={{ flexWrap: 'wrap', gap: 3 }}>{tabs.map((name, index) => <button key={name} id={`${id}-format-shape-tab-${index}`} type="button" role="tab" aria-controls={panelId} aria-selected={tab === name} tabIndex={tab === name ? 0 : -1} onClick={() => setTab(name)} onKeyDown={(event) => {
      const current = tabs.indexOf(name);
      const next = event.key === 'ArrowRight' ? (current + 1) % tabs.length
        : event.key === 'ArrowLeft' ? (current - 1 + tabs.length) % tabs.length
          : event.key === 'Home' ? 0
            : event.key === 'End' ? tabs.length - 1 : -1;
      if (next < 0) return;
      event.preventDefault();
      const nextTab = tabs[next];
      if (nextTab) setTab(nextTab);
      event.currentTarget.parentElement?.querySelectorAll<HTMLButtonElement>('[role="tab"]')[next]?.focus();
    }}>{t(name)}</button>)}</div>
    <div id={panelId} aria-labelledby={`${id}-format-shape-tab-${Math.max(0, tabs.indexOf(tab))}`} role="tabpanel" className="col" style={{ ...style, minHeight: 290, padding: 8 }}>
      {tab === 'Colors and Lines' && <div className="col" style={{ gap: 12 }}>
        {(!catalog.isLine || catalog.isPic) && <fieldset className="row" style={{ flexWrap: 'wrap', alignItems: 'end' }}><legend>{t('Fill')}</legend><ColorFill label="Color:" fill={fill} colors={catalog.colors} onChange={(next) => update('fill', next)} />{!catalog.isPic && <label>{t('Transparency')}: {Math.round((1 - Number(value('alpha', catalog.alpha)))*100)}%<input aria-label={t('Transparency')} type="range" min="0" max="100" step="5" value={Math.round((1 - Number(value('alpha', catalog.alpha)))*100)} onChange={(event) => update('alpha', 1 - Number(event.currentTarget.value) / 100)} /></label>}</fieldset>}
        <fieldset className="row" style={{ flexWrap: 'wrap', alignItems: 'end' }}><legend>{t('Line')}</legend><ColorFill label="Color:" fill={line} colors={catalog.colors} line onChange={(next) => update('lineColor', next)} /><label>{t('Dashed')}<select aria-label={t('Dashed')} value={value('dash', catalog.dash)} onChange={(event) => update('dash', event.currentTarget.value)}>{dashOptions.map(([key,name])=><option key={key} value={key}>{t(name)}</option>)}</select></label><label>{t('Style')}<select aria-label={t('Style')} value={value('cmpd', catalog.compound)} onChange={(event) => update('cmpd', event.currentTarget.value)}>{compoundOptions.map(([key,name])=><option key={key} value={key}>{t(name)}</option>)}</select></label>{number('Weight:', 'lineW', catalog.lineWidth, 0, 1584, .25, 'pt')}</fieldset>
        {catalog.showArrows && <fieldset className="row" style={{ flexWrap: 'wrap' }}><legend>{t('Arrows')}</legend>{arrowItems.map(([key,label,original])=><label key={key}>{t(label)}<select aria-label={t(label)} value={value(key, original)} onChange={(event) => update(key, event.currentTarget.value)}>{(key.endsWith('Sz')?arrowSizes:arrowTypes).map(([id,name])=><option key={id} value={id}>{t(name)}</option>)}</select></label>)}</fieldset>}
      </div>}
      {tab === 'Size' && <div className="col" style={{ gap: 12 }}><fieldset className="row" style={{ flexWrap: 'wrap' }}><legend>{t('Size and rotate')}</legend>{sizeNumber('Height:', 'h', catalog.height)}{sizeNumber('Width:', 'w', catalog.width)}<label className="col">{t('Rotation:')}<input aria-label={t('Rotation:')} type="number" min="-360" max="360" step="1" disabled={!catalog.canRotate} value={value('rot',catalog.rotation)} onChange={(event)=>update('rot',((Number(event.currentTarget.value)%360)+360)%360)} />°</label></fieldset><label><input type="checkbox" checked={value('lockAspect',catalog.lockAspect)} onChange={(event)=>update('lockAspect',event.currentTarget.checked)} />{t('Lock aspect ratio')}</label>{catalog.isPic&&<div className="col"><label>{t('Original size')}: {t('Height')} {catalog.original?.height??'—'}" {t('Width')} {catalog.original?.width??'—'}"</label><button type="button" disabled={!catalog.original} onClick={()=>setChanges((current)=>({...current,w:catalog.original!.width,h:catalog.original!.height,crop:{l:0,r:0,t:0,b:0},img:{}}))}>{t('Reset')}</button></div>}</div>}
      {tab === 'Position' && <div className="col"><fieldset className="col"><legend>{t('Position on slide')}</legend>{number('Horizontal:', 'x', catalog.x, -56, 56, .01, 'in')}{number('Vertical:', 'y', catalog.y, -56, 56, .01, 'in')}<div className="tp-note">{t('From: Top Left Corner')}</div></fieldset></div>}
      {tab === 'Picture' && catalog.isPic && <div className="col" style={{ gap: 10 }}><fieldset className="row" style={{ flexWrap: 'wrap' }}><legend>{t('Crop from')}</legend>{(['l','r','t','b'] as const).map((edge)=>number(`${{l:'Left:',r:'Right:',t:'Top:',b:'Bottom:'}[edge]}`,`crop${edge}`,catalog.crop[edge],-20,20,.01,'in'))}</fieldset><fieldset className="col"><legend>{t('Image control')}</legend><label>{t('Color')}<select aria-label={t('Color')} value={value('imgMode',catalog.imageMode)} onChange={(event)=>update('imgMode',event.currentTarget.value)}>{imageModes.map(([key,name])=><option key={key} value={key}>{t(name)}</option>)}</select></label><label>{t('Brightness')}: {Math.round(Number(value('bright',catalog.brightness/100))*100)}<input aria-label={t('Brightness')} type="range" min="-100" max="100" value={Math.round(Number(value('bright',catalog.brightness/100))*100)} onChange={(event)=>update('bright',Number(event.currentTarget.value)/100)} /></label><label>{t('Contrast')}: {Math.round(Number(value('contrast',catalog.contrast/100))*100)}<input aria-label={t('Contrast')} type="range" min="-100" max="100" value={Math.round(Number(value('contrast',catalog.contrast/100))*100)} onChange={(event)=>update('contrast',Number(event.currentTarget.value)/100)} /></label></fieldset></div>}
      {tab === 'Text Box' && <div className="col" style={{ gap: 10 }}><label>{t('Text anchor point')}<select aria-label={t('Text anchor point')} value={value('anchor',catalog.anchor)} onChange={(event)=>update('anchor',event.currentTarget.value)}>{anchors.map(([key,name])=><option key={key} value={key}>{t(name)}</option>)}</select></label><fieldset className="row" style={{flexWrap:'wrap'}}><legend>{t('Internal margin')}</legend>{(['Left:','Top:','Right:','Bottom:'] as const).map((label,index)=><label key={label}>{t(label)}<input aria-label={t(label)} type="number" min="0" max="20" step=".01" value={((value('ins',catalog.margins.map((margin)=>margin*72)) as number[])[index]||0)/72} onChange={(event)=>{const margins=[...(value('ins',catalog.margins.map((margin)=>margin*72)) as number[])];margins[index]=Number(event.currentTarget.value)*72;update('ins',margins)}} /></label>)}</fieldset><label><input type="checkbox" checked={value('wrap',catalog.wrap)} onChange={(event)=>update('wrap',event.currentTarget.checked)} />{t('Word wrap text in AutoShape')}</label><label><input type="checkbox" checked={value('autofit',catalog.autofit)} onChange={(event)=>update('autofit',event.currentTarget.checked)} />{t('Resize AutoShape to fit text')}</label><label><input type="checkbox" checked={value('vert',catalog.verticalText)} onChange={(event)=>update('vert',event.currentTarget.checked)} />{t('Rotate text within AutoShape by 90°')}</label></div>}
      {tab === 'Web' && <div className="col"><label>{t('Alternative text')}<textarea aria-label={t('Alternative text')} rows={5} value={value('alt',catalog.alt)} onChange={(event)=>update('alt',event.currentTarget.value)} /></label><div className="tp-note">{t('Screen readers and web pages use alternative text to describe the object.')}</div></div>}
    </div>
    <div className="vibeoffice-dialog__actions"><button type="button" className="primary" onClick={commit}>{t('OK')}</button><button type="button" onClick={()=>runtime?.dlg?.reactFormatShapeAction?.('close')}>{t('Cancel')}</button></div>
  </OfficeDialog>;
}
