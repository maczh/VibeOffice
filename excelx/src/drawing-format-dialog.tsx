import { OfficeDialog } from '@vibeoffice/core';
import { useEffect, useId, useState } from 'react';
import './drawing-format-dialog.css';
import { useExcelxHost, useExcelxRuntime } from './editor-host';

type Catalog = { title: string; kind: string; fill: string | null; line: string | null; lineW: number; width: number; height: number; editAs: string; alt: string; locked: boolean; print: boolean; aspect: boolean; hasLine: boolean; hasPosition: boolean; hasChart: boolean };
type Runtime = { drawing?: { reactFormatCatalog?(): Catalog | null; reactFormatAction?(action: string, patch?: Record<string, unknown>): boolean }; t?(message: string): string };
/** React-owned Format Object dialog; drawing geometry and workbook mutations stay in drawing.js. */
export function ExcelxDrawingFormatDialog() {
  const host = useExcelxHost();
  const runtime = useExcelxRuntime<Runtime>();
  const [catalog, setCatalog] = useState<Catalog | null>(null);
  const [tab, setTab] = useState('');
  const id = useId().replaceAll(':', '');
  const t = (text: string) => runtime?.t?.(text) || text;
  const action = (name: string, patch?: Record<string, unknown>) => runtime?.drawing?.reactFormatAction?.(name, patch);
  const listen = (eventName: string) => {
    const handler = (event: Event) => { const detail = (event as CustomEvent<Catalog | null>).detail; setCatalog(detail); if (detail) setTab(detail.hasChart ? 'Chart' : detail.hasLine ? 'Colors and Lines' : 'Size'); };
    host.ownerDocument.addEventListener(eventName, handler); return () => host.ownerDocument.removeEventListener(eventName, handler);
  };
  useEffect(() => { const dispose = listen('excelx-drawing-format'); return () => { dispose(); runtime?.drawing?.reactFormatAction?.('cancel'); }; }, [host, runtime]);
  if (!catalog) return null;
  const change = (patch: Record<string, unknown>) => action('update', patch);
  const close = () => { action('cancel'); setCatalog(null); };
  const save = () => { action('apply'); setCatalog(null); };
  const tabs = [catalog.hasChart ? 'Chart' : '', catalog.hasLine ? 'Colors and Lines' : '', 'Size', catalog.hasPosition ? 'Properties' : '', 'Web'].filter(Boolean);
  const tabIndex = tabs.indexOf(tab);
  const tabId = (index: number) => `${id}-drawing-format-tab-${index}`;
  const panelId = `${id}-drawing-format-panel`;
  return <OfficeDialog title={t(catalog.title)} className="excelx-drawing-format" style={{ width: 'min(94vw, 500px)' }} onRequestClose={close}>
    <div className="excelx-drawing-format__tabs" role="tablist" aria-label={t(catalog.title)}>{tabs.map((name, index) => <button type="button" role="tab" id={tabId(index)} aria-controls={panelId} aria-selected={tab===name} tabIndex={tab===name?0:-1} key={name} onClick={()=>setTab(name)} onKeyDown={(event)=>{
      const next=event.key==='ArrowRight'?(index+1)%tabs.length:event.key==='ArrowLeft'?(index-1+tabs.length)%tabs.length:event.key==='Home'?0:event.key==='End'?tabs.length-1:-1;
      if(next<0)return;
      event.preventDefault();
      const nextTab=tabs[next];if(nextTab)setTab(nextTab);
      event.currentTarget.parentElement?.querySelectorAll<HTMLButtonElement>('[role="tab"]')[next]?.focus();
    }}>{t(name)}</button>)}</div>
    <div id={panelId} aria-labelledby={tabId(Math.max(0,tabIndex))} role="tabpanel" className="excelx-drawing-format__body">
      {tab==='Chart'&&<div className="excelx-drawing-format__chart"><button type="button" onClick={()=>{action('chartOptions');setCatalog(null)}}>{t('Chart &Options...').replace('&','')}</button><button type="button" onClick={(event)=>{action('chartType',{x:event.clientX,y:event.clientY});setCatalog(null)}}>{t('Chart &Type...').replace('&','')}</button></div>}
      {tab==='Colors and Lines'&&<><fieldset><legend>{t('Fill')}</legend><label>{t('&Color:')}<button type="button" className="swatch" aria-label={t('Choose fill color')} style={{background:catalog.fill||'transparent'}} onClick={(event)=>action('color',{key:'fill',anchor:event.currentTarget})}><span className="dd-arrow"/></button></label></fieldset>{catalog.kind==='shape'&&<fieldset><legend>{t('Line')}</legend><label>{t('C&olor:')}<button type="button" className="swatch" aria-label={t('Choose line color')} style={{background:catalog.line||'transparent'}} onClick={(event)=>action('color',{key:'line',anchor:event.currentTarget})}><span className="dd-arrow"/></button></label><label>{t('&Weight:')}<input type="number" min="0.25" max="20" step="0.25" value={catalog.lineW} onChange={(event)=>change({lineW:Number(event.currentTarget.value)})}/>{t(' pt')}</label></fieldset>}</>}
      {tab==='Size'&&<><fieldset><legend>{t('Size and rotate')}</legend><label>{t('H&eight:')}<input type="number" min="0.01" max="50" step="0.1" value={(catalog.height/72).toFixed(2)} onChange={(event)=>change({height:Number(event.currentTarget.value)*72})}/>{'"'}</label><label>{t('Wi&dth:')}<input type="number" min="0.01" max="50" step="0.1" value={(catalog.width/72).toFixed(2)} onChange={(event)=>change({width:Number(event.currentTarget.value)*72})}/>{'"'}</label><label><input type="checkbox" checked={catalog.aspect} onChange={(event)=>change({aspect:event.currentTarget.checked})}/>{t('Lock &aspect ratio')}</label></fieldset></>}
      {tab==='Properties'&&<><fieldset><legend>{t('Object positioning')}</legend>{[['twoCell','&Move and size with cells'],['oneCell',"Move &but don't size with cells"],['absolute',"&Don't move or size with cells"]].map(([value,label])=><label key={value}><input type="radio" name="drawing-edit-as" value={value} checked={catalog.editAs===value} onChange={()=>change({editAs:value})}/>{t(label || '')}</label>)}</fieldset><label><input type="checkbox" checked={catalog.print} onChange={(event)=>change({print:event.currentTarget.checked})}/>{t('&Print object')}</label><label><input type="checkbox" checked={catalog.locked} onChange={(event)=>change({locked:event.currentTarget.checked})}/>{t('&Locked')}</label></>}
      {tab==='Web'&&<label>{t('Alternative text:')}<textarea rows={4} value={catalog.alt} onChange={(event)=>change({alt:event.currentTarget.value})}/></label>}
    </div><div className="excelx-drawing-format__actions"><button type="button" className="primary" onClick={save}>{t('OK')}</button><button type="button" onClick={close}>{t('Cancel')}</button></div>
  </OfficeDialog>;
}
