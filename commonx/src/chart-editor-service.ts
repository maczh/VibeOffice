type ChartRecord = Record<string, any>;
type ChartApi = {
  edit(shape: ChartRecord, onDone: (chart: ChartRecord) => void): void;
  sample(): ChartRecord;
  render(chart: ChartRecord, width: number, height: number, design?: unknown): SVGSVGElement;
  kinds(): Array<[string, string]>;
  syncV2(chart: ChartRecord, previousKind?: string): void;
  setSeriesColor(chart: ChartRecord, series: ChartRecord, color: string): void;
  PALETTE: string[];
  [key: string]: any;
};
type ChartRuntime = Window & typeof globalThis & {
  L?: {
    chart?: ChartApi;
    ui?: { currentDesign?(): unknown; colorMenu?(anchor: HTMLElement, options: { mode: string }, onPick: (color: unknown) => void): void };
    bus?: { emit(name: string): void };
    clone?(value: unknown): any;
    round?(value: number, digits: number): number;
    t?(value: string, ...args: unknown[]): string;
  };
};

/** Installs the direct-host chart-editor command adapter while leaving classic DOM editing intact. */
export function installReactChartEditor(runtime: Window): () => void {
  const view = runtime as ChartRuntime;
  const L = view.L;
  const chart = L?.chart;
  if (!L || !chart || typeof chart.edit !== 'function') return () => {};

  let active: { chart: ChartRecord; shape: ChartRecord; onDone: (chart: ChartRecord) => void } | null = null;
  const installed = new Map<string, (...args: any[]) => any>();
  const originals = new Map<string, PropertyDescriptor | undefined>();
  const notify = () => L.bus?.emit('common-chart-editor');
  const translate = (value: string) => L.t?.(value) || value;
  const legends: Array<[string, string]> = [['r', 'Right'], ['b', 'Bottom'], ['t', 'Top'], ['l', 'Left'], ['none', 'None']];
  const preview = () => {
    if (!active) return '';
    const { chart: model, shape } = active;
    const width = shape?.w || 300, height = shape?.h || 190;
    if (model.v === 2) chart.syncV2(model);
    const scale = model.v === 2 ? Math.min(300 / width, 190 / height) : 1;
    const svg = chart.render(model, model.v === 2 ? width : 300, model.v === 2 ? height : 190, L.ui?.currentDesign?.());
    svg.setAttribute('width', String((L.round || ((value: number) => value))((model.v === 2 ? width : 300) * scale, 2)));
    svg.setAttribute('height', String((L.round || ((value: number) => value))((model.v === 2 ? height : 190) * scale, 2)));
    return svg.outerHTML;
  };
  const methods: Record<string, (...args: any[]) => any> = {
    edit(shape: ChartRecord, onDone: (chart: ChartRecord) => void) {
      active = { chart: L.clone ? L.clone(shape.chart || chart.sample()) : structuredClone(shape.chart || chart.sample()), shape, onDone };
      notify();
    },
    reactEditorCatalog() {
      if (!active) return null;
      const model = active.chart;
      return {
        kind: model.kind,
        kinds: chart.kinds().map(([value, label]) => ({ value, label: translate(label) })),
        title: model.title || '', legend: model.legend || 'r',
        legends: legends.map(([value, label]) => ({ value, label: translate(label) })),
        gridY: model.gridY !== false, labels: !!model.labels, cats: model.cats.slice(),
        series: model.series.map((series: ChartRecord, index: number) => ({ index, name: series.name, vals: series.vals.slice(), color: series.color || chart.PALETTE[index % chart.PALETTE.length] })),
        preview: preview(),
      };
    },
    reactEditorAction(action: string, first?: any, second?: any) {
      if (!active) return false;
      const model = active.chart;
      if (action === 'close') { active = null; notify(); return true; }
      if (action === 'commit') {
        const done = active.onDone;
        chart.syncV2(model); model.edited = true;
        active = null; notify(); done(model);
        return true;
      }
      if (action === 'kind' && chart.kinds().some(([value]) => value === first)) { const previous = model.kind; model.kind = first; chart.syncV2(model, previous); }
      else if (action === 'title' && typeof first === 'string') model.title = first.slice(0, 255);
      else if (action === 'legend' && ['r', 'b', 't', 'l', 'none'].includes(first)) model.legend = first;
      else if (action === 'grid' && typeof first === 'boolean') model.gridY = first;
      else if (action === 'labels' && typeof first === 'boolean') model.labels = first;
      else if (action === 'category' && Number.isInteger(first) && first >= 0 && first < model.cats.length && typeof second === 'string') model.cats[first] = second.slice(0, 255);
      else if (action === 'series-name' && Number.isInteger(first) && first >= 0 && first < model.series.length && typeof second === 'string') model.series[first].name = second.slice(0, 255);
      else if (action === 'value' && Number.isInteger(first) && first >= 0 && first < model.series.length && Number.isInteger(second?.category) && second.category >= 0 && second.category < model.cats.length && Number.isFinite(second.value)) model.series[first].vals[second.category] = second.value;
      else if (action === 'add-category') { model.cats.push(`Col ${model.cats.length + 1}`); model.series.forEach((series: ChartRecord) => series.vals.push(0)); }
      else if (action === 'delete-category' && model.cats.length > 1) { model.cats.pop(); model.series.forEach((series: ChartRecord) => series.vals.pop()); }
      else if (action === 'add-series') model.series.push({ name: `Series ${model.series.length + 1}`, vals: model.cats.map(() => 0), color: chart.PALETTE[model.series.length % chart.PALETTE.length] });
      else if (action === 'delete-series' && model.series.length > 1) model.series.pop();
      else if (action === 'series-color' && Number.isInteger(first) && first >= 0 && first < model.series.length && typeof second === 'string') chart.setSeriesColor(model, model.series[first], second);
      else return false;
      chart.syncV2(model);
      notify();
      return true;
    },
    reactEditorColorMenu(index: number, anchor: HTMLElement) {
      if (!active || !Number.isInteger(index) || index < 0 || index >= active.chart.series.length || !(anchor instanceof view.HTMLElement)) return false;
      const colorMenu = L.ui?.colorMenu;
      if (typeof colorMenu === 'function') colorMenu(anchor, { mode: 'plain' }, (color) => {
        const action = methods.reactEditorAction;
        if (typeof color === 'string' && typeof action === 'function') action('series-color', index, color);
      });
      return true;
    },
  };

  for (const [name, method] of Object.entries(methods)) {
    originals.set(name, Object.getOwnPropertyDescriptor(chart, name));
    installed.set(name, method);
    Object.defineProperty(chart, name, { configurable: true, enumerable: true, writable: true, value: method });
  }

  return () => {
    if (active) { active = null; notify(); }
    for (const [name, method] of installed) {
      if (chart[name] !== method) continue;
      const original = originals.get(name);
      if (original) Object.defineProperty(chart, name, original);
      else delete chart[name];
    }
  };
}
