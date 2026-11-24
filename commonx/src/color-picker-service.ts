import type { ColorMenuItem, ColorMenuSection, ColorMenuRequest } from './color-menu.js';
import { standardColorName } from './standard-colors.js';

interface SchemeColor { tip?: string; ref: string; hex: string }
interface ColorOptions { mode?: 'fill' | 'line' | 'font' | 'plain'; design?: unknown; scheme?: boolean; grid?: boolean; automatic?: boolean; effects?: boolean }
interface PaletteColorApi {
  STANDARD: string[];
  hslToRgb(hue: number, saturation: number, lightness: number): [number, number, number];
  rgbToHex(red: number, green: number, blue: number): string;
}
interface RuntimeUi {
  currentDesign?: () => unknown;
  schemeColors?: (design: unknown) => SchemeColor[];
  addRecentColor?: (color: string) => void;
  colorMenu?: (anchor: Element | { left?: number; right?: number; top?: number; bottom?: number; x?: number; y?: number }, options: ColorOptions, choose: (value: unknown) => void) => void;
  moreColors?: (initial: string | null, respond: (color: string | null) => void) => void;
  [name: string]: any;
}
interface RuntimeNamespace {
  ui?: RuntimeUi;
  color?: PaletteColorApi;
  clamp?: (value: number, min: number, max: number) => number;
  t?: (text: string) => string;
}
interface ColorRequest extends Omit<ColorMenuRequest, 'anchor'> { anchor?: Element }
interface MoreColorsRequest { initial: string; rows: string[][]; grays: string[]; respond: (color: string | null) => void }

const recentColors = new WeakMap<Window, string[]>();

/** Installs React-host color palette and custom-color dialog adapters. */
export function installColorPickerService(runtime: Window): void {
  const view = runtime as Window & {
    L?: RuntimeNamespace;
    VO?: { reactDirectHost?: { root?: HTMLElement } };
  };
  const L = view.L ?? (view.L = {});
  const ui = L.ui ?? (L.ui = {});
  const recent = recentColors.get(runtime) ?? [];
  recentColors.set(runtime, recent);
  const CustomEventConstructor = (runtime as Window & typeof globalThis).CustomEvent;
  const translate = (text: string) => L.t ? L.t(text) : text;
  const addRecentColor = (color: string) => {
    if (!color || color[0] !== '#') return;
    const index = recent.indexOf(color);
    if (index >= 0) recent.splice(index, 1);
    recent.unshift(color);
    recent.length = Math.min(recent.length, 8);
  };
  ui.addRecentColor = addRecentColor;

  ui.colorMenu = (anchor, options, onPick) => {
    const root = view.VO?.reactDirectHost?.root;
    if (!root?.isConnected) return;
    const o = options || {};
    const design = o.design || ui.currentDesign?.() || null;
    const scheme = o.scheme != null ? !!o.scheme : !!ui.schemeColors;
    const bounds = 'getBoundingClientRect' in anchor ? anchor.getBoundingClientRect() : anchor;
    const sections: ColorMenuSection[] = [];
    const separator = () => sections.push({ layout: 'separator' });
    const wide = (label: string, value?: unknown, special?: ColorMenuItem['special']) => sections.push({ layout: 'wide', item: { label: translate(label), ...(value !== undefined ? { value } : {}), ...(special ? { special } : {}), wide: true } });
    const row = (layout: 'row' | 'grid', items: ColorMenuItem[]) => sections.push({ layout, items });
    if (o.mode === 'fill' || o.mode === 'line') wide(o.mode === 'fill' ? 'No Fill' : 'No Line', { none: true });
    if (o.mode === 'font' || o.automatic) wide('Automatic', { auto: true });
    separator();
    if (scheme && ui.schemeColors) row('row', ui.schemeColors(design).map((color) => ({ label: color.tip || color.ref, color: color.hex, tip: color.tip || color.ref, value: color.ref })));
    if (recent.length) { separator(); row('row', recent.map((color) => ({ label: color, color, value: color }))); }
    if (o.mode === 'plain' || o.grid || !scheme) {
      if (scheme) separator();
      const standard = L.color?.STANDARD || [];
      row('grid', standard.map((color) => {
        const label = standardColorName(color, translate);
        return { label, color, tip: label, value: color, recent: true };
      }));
    }
    separator();
    const moreLabel = o.mode === 'fill' ? 'More Fill Colors...' : o.mode === 'line' ? 'More Line Colors...' : 'More Colors...';
    wide(moreLabel, undefined, 'more');
    if (o.effects) wide(o.mode === 'line' ? 'Patterned Lines...' : 'Fill Effects...', undefined, 'effects');
    const anchorElement = 'getBoundingClientRect' in anchor ? anchor : undefined;
    const dispatchTarget = anchorElement && typeof anchorElement.dispatchEvent === 'function' ? anchorElement : root;
    dispatchTarget.dispatchEvent(new CustomEventConstructor('vibeoffice-react-color-menu', { detail: {
      left: bounds.left || bounds.x || 0,
      bottom: bounds.bottom || bounds.y || 0,
      ...(anchorElement ? { anchor: anchorElement } : {}),
      sections,
      choose: (value: unknown, isRecent?: boolean) => { if (isRecent && typeof value === 'string') addRecentColor(value); onPick(value); },
      more: () => ui.moreColors?.(null, (color) => { if (color) { addRecentColor(color); onPick(color); } }),
    } satisfies ColorRequest, bubbles: true }));
  };

  ui.moreColors = (initial, callback) => {
    const root = view.VO?.reactDirectHost?.root;
    if (!root?.isConnected) return;
    const colorApi = L.color;
    if (!colorApi) return;
    const cur = initial?.[0] === '#' ? initial : '#3366FF';
    const clamp = L.clamp || ((value, min, max) => Math.max(min, Math.min(max, value)));
    const rows = Array.from({ length: 13 }, (_, rowIndex) => {
      const count = rowIndex < 7 ? 7 + rowIndex : 19 - rowIndex;
      const distance = Math.abs(rowIndex - 6) / 6;
      const lightness = 0.5 + (rowIndex < 6 ? 0.35 : -0.25) * distance;
      return Array.from({ length: count }, (_, index) => {
        const hue = ((index / count) * 360 + rowIndex * 23) % 360;
        const [red, green, blue] = colorApi.hslToRgb(hue / 360, 1 - distance * 0.4, clamp(lightness, 0.1, 0.92));
        return colorApi.rgbToHex(red, green, blue);
      });
    });
    const grays = Array.from({ length: 16 }, (_, index) => {
      const value = Math.round((255 * index) / 15);
      return colorApi.rgbToHex(value, value, value);
    });
    root.dispatchEvent(new CustomEventConstructor('vibeoffice-react-more-colors', { detail: {
      initial: cur,
      rows,
      grays,
      respond: (color: string | null) => { if (color) addRecentColor(color); callback(color); },
    } satisfies MoreColorsRequest, bubbles: true }));
  };
}
