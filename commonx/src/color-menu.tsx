import { useEffect, useLayoutEffect, useRef } from 'react';

export interface ColorMenuItem {
  label: string;
  color?: string;
  value?: unknown;
  tip?: string;
  special?: 'more' | 'effects';
  wide?: boolean;
  recent?: boolean;
}

export type ColorMenuSection = { layout: 'separator' } | { layout: 'row' | 'grid'; items: ColorMenuItem[] } | { layout: 'wide'; item: ColorMenuItem };

export interface ColorMenuRequest {
  left: number;
  bottom: number;
  anchor?: Element;
  sections: ColorMenuSection[];
  choose: (value: unknown, recent?: boolean) => void;
  more: () => void;
}

interface ColorMenuProps {
  request: ColorMenuRequest;
  onClose: () => void;
}

/** Shared React color palette, with color-model decisions delegated to its editor runtime. */
export function ColorMenu({ request, onClose }: ColorMenuProps) {
  const panel = useRef<HTMLDivElement>(null);

  useLayoutEffect(() => {
    if (!panel.current) return;
    const bounds = panel.current.getBoundingClientRect();
    panel.current.style.left = `${Math.max(4, Math.min(request.left, window.innerWidth - bounds.width - 4))}px`;
    panel.current.style.top = `${Math.max(4, Math.min(request.bottom, window.innerHeight - bounds.height - 4))}px`;
    panel.current.querySelector<HTMLButtonElement>('button')?.focus();
  }, [request]);

  useEffect(() => {
    const outside = (event: PointerEvent) => {
      const target = event.target;
      if (target instanceof Node && !panel.current?.contains(target) && !request.anchor?.contains(target)) onClose();
    };
    const keydown = (event: KeyboardEvent) => {
      if (event.key === 'Escape') { event.preventDefault(); event.stopPropagation(); onClose(); }
      if (event.key === 'ArrowRight' || event.key === 'ArrowDown') {
        const buttons = [...(panel.current?.querySelectorAll<HTMLButtonElement>('button') ?? [])].filter((button) => !button.disabled);
        const index = buttons.indexOf(document.activeElement as HTMLButtonElement);
        if (buttons.length) { event.preventDefault(); event.stopPropagation(); buttons[(index + 1 + buttons.length) % buttons.length]?.focus(); }
      }
      if (event.key === 'ArrowLeft' || event.key === 'ArrowUp') {
        const buttons = [...(panel.current?.querySelectorAll<HTMLButtonElement>('button') ?? [])].filter((button) => !button.disabled);
        const index = buttons.indexOf(document.activeElement as HTMLButtonElement);
        if (buttons.length) { event.preventDefault(); event.stopPropagation(); buttons[(index - 1 + buttons.length) % buttons.length]?.focus(); }
      }
    };
    document.addEventListener('pointerdown', outside, true);
    document.addEventListener('keydown', keydown, true);
    return () => {
      document.removeEventListener('pointerdown', outside, true);
      document.removeEventListener('keydown', keydown, true);
    };
  }, [onClose, request]);

  const renderItem = (item: ColorMenuItem, index: number) => {
    const action = () => {
      if (item.special === 'more') { onClose(); request.more(); }
      else if (item.special === 'effects') { onClose(); request.choose({ effects: true }); }
      else { onClose(); request.choose(item.value, item.recent); }
    };
    if (item.wide || item.special) return <button key={`${item.label}:${index}`} role="menuitem" className="cp-wide" type="button" onClick={action}>{item.label}</button>;
    return <button key={`${item.label}:${index}`} role="menuitem" className="cp-sw" type="button" style={item.color ? { background: item.color } : undefined} aria-label={item.label} title={item.tip} onClick={action} />;
  };

  return <div ref={panel} className="vibeoffice-color-menu menu color-menu" role="menu">
    <div className="color-panel">{request.sections.map((section, index) => section.layout === 'separator'
      ? <div key={`separator:${index}`} className="cp-sep" />
      : section.layout === 'wide' ? renderItem(section.item, index)
        : <div key={`${section.layout}:${index}`} className={section.layout === 'grid' ? 'cp-grid' : 'cp-row'}>{section.items.map(renderItem)}</div>)}</div>
  </div>;
}
