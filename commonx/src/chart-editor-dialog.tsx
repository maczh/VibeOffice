import { OfficeDialog } from './dialog.js';
import { useEffect, useState } from 'react';
import './chart-editor-dialog.css';

interface ChartSeries { index: number; name: string; vals: number[]; color: string }
interface ChartCatalog {
  kind: string;
  kinds: Array<{ value: string; label: string }>;
  title: string;
  legend: string;
  legends: Array<{ value: string; label: string }>;
  gridY: boolean;
  labels: boolean;
  cats: string[];
  series: ChartSeries[];
  preview: string;
}
interface ChartRuntime {
  chart?: {
    reactEditorCatalog?(): ChartCatalog | null;
    reactEditorAction?(action: string, first?: unknown, second?: unknown): boolean;
    reactEditorColorMenu?(index: number, anchor: HTMLElement): boolean;
  };
  bus?: { on(event: string, listener: () => void): () => void };
  t?(message: string, ...args: unknown[]): string;
}
function runtime(): ChartRuntime | undefined { return (window as Window & { L?: ChartRuntime }).L; }

/** Shared React Chart dialog; chart rendering and file-model updates stay in the existing chart service. */
export function ChartEditorDialog() {
  const [catalog, setCatalog] = useState<ChartCatalog | null>(null);
  useEffect(() => {
    const refresh = () => {
      const next = runtime()?.chart?.reactEditorCatalog?.() || null;
      setCatalog(next ? { ...next, kinds: [...next.kinds], legends: [...next.legends], cats: [...next.cats], series: next.series.map((series) => ({ ...series, vals: [...series.vals] })) } : null);
    };
    const off = runtime()?.bus?.on('common-chart-editor', refresh);
    return () => off?.();
  }, []);
  if (!catalog) return null;
  const t = (message: string, ...args: unknown[]) => runtime()?.t?.(message, ...args) || message;
  const update = (action: string, first?: unknown, second?: unknown) => runtime()?.chart?.reactEditorAction?.(action, first, second);
  const close = () => update('close');
  return <OfficeDialog title={t('Chart')} className="vibeoffice-chart-editor-dialog" style={{ width: 'min(96vw, 760px)' }} onRequestClose={close}>
    <div className="vibeoffice-chart-editor">
      <div className="vibeoffice-chart-editor__settings">
        <label>{t('Chart type:')}<select aria-label={t('Chart type')} value={catalog.kind} onChange={(event) => update('kind', event.currentTarget.value)}>{catalog.kinds.map((kind) => <option key={kind.value} value={kind.value}>{kind.label}</option>)}</select></label>
        <label>{t('Chart title:')}<input aria-label={t('Chart title')} type="text" value={catalog.title} onChange={(event) => update('title', event.currentTarget.value)} /></label>
        <label>{t('Legend:')}<select aria-label={t('Legend')} value={catalog.legend} onChange={(event) => update('legend', event.currentTarget.value)}>{catalog.legends.map((legend) => <option key={legend.value} value={legend.value}>{legend.label}</option>)}</select></label>
        <div className="vibeoffice-chart-editor__checks"><label><input type="checkbox" checked={catalog.gridY} onChange={(event) => update('grid', event.currentTarget.checked)} />{t('Major gridlines')}</label><label><input type="checkbox" checked={catalog.labels} onChange={(event) => update('labels', event.currentTarget.checked)} />{t('Show values')}</label></div>
      </div>
      <div className="vibeoffice-chart-editor__preview"><span>{t('Preview')}</span><div aria-label={t('Chart preview')} dangerouslySetInnerHTML={{ __html: catalog.preview }} /></div>
      <fieldset className="vibeoffice-chart-editor__datasheet"><legend>{t('Datasheet')}</legend>
        <div className="vibeoffice-chart-editor__table-scroll"><table><thead><tr><th className="corner" />{catalog.cats.map((category, index) => <th key={`category-${index}`}><input aria-label={t('Category {0}', index + 1)} value={category} onChange={(event) => update('category', index, event.currentTarget.value)} /></th>)}</tr></thead>
          <tbody>{catalog.series.map((series) => <tr key={`series-${series.index}`}><th><button type="button" className="vibeoffice-chart-editor__swatch" aria-label={t('Series color: {0}', series.name)} title={t('Choose series color')} style={{ background: series.color }} onClick={(event) => runtime()?.chart?.reactEditorColorMenu?.(series.index, event.currentTarget)} /><input aria-label={t('Series name {0}', series.index + 1)} value={series.name} onChange={(event) => update('series-name', series.index, event.currentTarget.value)} /></th>{series.vals.map((value, category) => <td key={`${series.index}-${category}`}><input aria-label={t('{0}, {1}', series.name, catalog.cats[category])} inputMode="decimal" value={value} onChange={(event) => { const parsed = Number.parseFloat(event.currentTarget.value); update('value', series.index, { category, value: Number.isFinite(parsed) ? parsed : 0 }); }} /></td>)}</tr>)}</tbody>
        </table></div>
        <div className="vibeoffice-chart-editor__tools"><button type="button" onClick={() => update('add-category')}>{t('Add Column')}</button><button type="button" disabled={catalog.cats.length <= 1} onClick={() => update('delete-category')}>{t('Delete Column')}</button><button type="button" onClick={() => update('add-series')}>{t('Add Row')}</button><button type="button" disabled={catalog.series.length <= 1} onClick={() => update('delete-series')}>{t('Delete Row')}</button></div>
      </fieldset>
    </div>
    <div className="vibeoffice-dialog__actions"><button type="button" className="primary" onClick={() => update('commit')}>{t('OK')}</button><button type="button" onClick={close}>{t('Cancel')}</button></div>
  </OfficeDialog>;
}
