import { useMemo, useState } from 'react';
import { OfficeDialog } from './dialog.js';
import { useVibeOffice } from './react-bridge.js';

export interface MoreColorsRequest {
  initial: string;
  rows: string[][];
  grays: string[];
  respond: (color: string | null) => void;
}

interface MoreColorsDialogProps {
  request: MoreColorsRequest;
  onClose: () => void;
}

function clampChannel(value: number): number {
  return Math.max(0, Math.min(255, Math.round(Number.isFinite(value) ? value : 0)));
}

function hexRgb(hex: string): [number, number, number] {
  const value = hex.replace(/^#/, '');
  return [0, 2, 4].map((offset) => Number.parseInt(value.slice(offset, offset + 2), 16)) as [number, number, number];
}

function rgbHex([red, green, blue]: [number, number, number]): string {
  return `#${[red, green, blue].map((value) => value.toString(16).padStart(2, '0')).join('')}`.toUpperCase();
}

/** Custom color picker retaining the legacy standard palette and RGB/hex editing workflow. */
export function MoreColorsDialog({ request, onClose }: MoreColorsDialogProps) {
  const { translate } = useVibeOffice();
  const initial = request.initial.toUpperCase();
  const [color, setColor] = useState(initial);
  const [rgb, setRgb] = useState<[number, number, number]>(() => hexRgb(initial));
  const [hex, setHex] = useState(initial);
  const [tab, setTab] = useState<'standard' | 'custom'>('standard');
  const update = (value: string) => {
    if (!/^#[\da-f]{6}$/i.test(value)) return;
    const next = value.toUpperCase();
    setColor(next);
    setHex(next);
    setRgb(hexRgb(next));
  };
  const setChannel = (index: number, value: number) => {
    const next = [...rgb] as [number, number, number];
    next[index] = clampChannel(value);
    setRgb(next);
    setColor(rgbHex(next));
    setHex(rgbHex(next));
  };
  const finish = (result: string | null) => {
    request.respond(result);
    onClose();
  };
  const tabs = useMemo(() => [
    { id: 'standard' as const, label: translate('Standard') },
    { id: 'custom' as const, label: translate('Custom') },
  ], [translate]);

  return <OfficeDialog title={translate('Colors')} className="vibeoffice-more-colors-dialog" style={{ width: 'min(92vw, 440px)' }} onRequestClose={() => finish(null)}>
    <form onSubmit={(event) => { event.preventDefault(); finish(color); }}>
      <div className="cc-wrap">
        <div className="tabs">
          <div className="tabs-head" role="tablist">{tabs.map((item) => <button key={item.id} className={`tab${tab === item.id ? ' on' : ''}`} type="button" role="tab" aria-selected={tab === item.id} onClick={() => setTab(item.id)}>{item.label}</button>)}</div>
          <div className="tabs-body">
            {tab === 'standard' ? <div role="tabpanel">
              <div className="cc-cap">{translate('Colors:')}</div>
              <div className="cc-honey">{request.rows.map((row, index) => <div className="cc-hrow" key={index}>{row.map((value) => <button key={value} type="button" className="cc-hex" style={{ background: value }} aria-label={value} onClick={() => update(value)} />)}</div>)}</div>
              <div className="cc-grays">{request.grays.map((value) => <button key={value} type="button" className="cc-hex" style={{ background: value }} aria-label={value} onClick={() => update(value)} />)}</div>
            </div> : <div className="cc-custom" role="tabpanel">
              <div className="field"><label htmlFor="vibeoffice-color-native">{translate('Pick:')}</label><input id="vibeoffice-color-native" type="color" value={color.toLowerCase()} onChange={(event) => update(event.currentTarget.value)} /></div>
              {(['Red', 'Green', 'Blue'] as const).map((label, index) => <div className="field" key={label}><label htmlFor={`vibeoffice-color-${label}`}>{translate(label)}:</label><input id={`vibeoffice-color-${label}`} type="number" min="0" max="255" step="1" value={rgb[index]} onChange={(event) => setChannel(index, Number(event.currentTarget.value))} /></div>)}
              <div className="field"><label htmlFor="vibeoffice-color-hex">{translate('Hex:')}</label><input id="vibeoffice-color-hex" type="text" maxLength={7} value={hex} onChange={(event) => setHex(event.currentTarget.value)} onBlur={() => update(/^#?[\da-f]{6}$/i.test(hex) ? hex.startsWith('#') ? hex : `#${hex}` : color)} onKeyDown={(event) => { if (event.key === 'Enter') { event.preventDefault(); const value = /^#?[\da-f]{6}$/i.test(hex) ? hex.startsWith('#') ? hex : `#${hex}` : color; update(value); } }} /></div>
            </div>}
          </div>
        </div>
        <aside className="cc-side">
          <div className="cc-cap">{translate('New|color')}</div><div className="cc-new" style={{ background: color }} />
          <div className="cc-cur" style={{ background: initial }} /><div className="cc-cap">{translate('Current')}</div>
        </aside>
      </div>
      <div className="vibeoffice-more-colors-dialog__actions"><button className="primary" type="submit">{translate('OK')}</button><button type="button" onClick={() => finish(null)}>{translate('Cancel')}</button></div>
    </form>
  </OfficeDialog>;
}
