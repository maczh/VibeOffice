import { createPortal } from 'react-dom';
import { useEffect, useRef, useState } from 'react';

interface ScoutMount { element: HTMLElement; avatarElement: HTMLElement; balloonElement: HTMLElement; textElement: HTMLElement; choicesElement: HTMLElement; onText(text: string): void; onChoices(choices: string[], visible: boolean): void }
interface ScoutAssistantRuntime {
  t?(text: string): string;
  assistant?: { shown?(): boolean; activate?(): void; menuAt?(x: number, y: number): void; scout?(): { choose?(index: number): unknown; closeBalloon?(): unknown; stop?(): unknown; dragStart?(x: number, y: number): boolean; dragMove?(x: number, y: number): boolean; dragEnd?(commit?: boolean): { moved: boolean } | null } | null; attachReactHost?(factory: () => ScoutMount): () => void };
}
function runtime(view: Window | null): ScoutAssistantRuntime | undefined { return view ? (view as Window & { L?: ScoutAssistantRuntime }).L : undefined; }

/** React owns Scout's scoped host, avatar slot, and speech surface; the existing animation/service API remains authoritative. */
export function ScoutAssistant({ host }: { host: HTMLElement }) {
  const view = host.ownerDocument.defaultView;
  const rootRef = useRef<HTMLDivElement>(null);
  const avatarRef = useRef<HTMLButtonElement>(null);
  const balloonRef = useRef<HTMLDivElement>(null);
  const textRef = useRef<HTMLDivElement>(null);
  const choicesRef = useRef<HTMLDivElement>(null);
  const draggedRef = useRef(false);
  const [mount, setMount] = useState<HTMLElement | null>(null);
  const [speechText, setSpeechText] = useState('');
  const [choices, setChoices] = useState<string[]>([]);
  const [choicesVisible, setChoicesVisible] = useState(false);

  useEffect(() => {
    if (!view) return;
    const findMount = () => {
      const root = host.querySelector<HTMLElement>('.vibeoffice-legacy-root[data-vibeoffice-engine-root]');
      if (root) setMount(root);
    };
    const observer = new view.MutationObserver(findMount);
    observer.observe(host, { childList: true, subtree: true });
    findMount();
    return () => observer.disconnect();
  }, [host, view]);

  useEffect(() => {
    if (!mount) return;
    let disposed = false;
    let detach: (() => void) | undefined;
    const attach = () => {
      if (disposed || detach) return;
      const assistant = runtime(view)?.assistant;
      const element = rootRef.current, avatarElement = avatarRef.current, balloonElement = balloonRef.current, textElement = textRef.current, choicesElement = choicesRef.current;
      if (!assistant?.attachReactHost || !element || !avatarElement || !balloonElement || !textElement || !choicesElement) return;
      detach = assistant.attachReactHost(() => ({ element, avatarElement, balloonElement, textElement, choicesElement, onText: setSpeechText, onChoices: (items, visible) => { setChoices(items); setChoicesVisible(visible); } }));
    };
    mount.addEventListener('vibeoffice-engine-runtime-ready', attach);
    attach();
    return () => { disposed = true; mount.removeEventListener('vibeoffice-engine-runtime-ready', attach); detach?.(); };
  }, [mount, view]);

  if (!mount || !view) return null;
  return createPortal(<div ref={rootRef} className="scout-agent is-desktop" data-react-owned="scout-assistant">
    <button ref={avatarRef} className="scout-avatar" type="button" data-react-owned="scout-trigger" aria-label={runtime(view)?.t?.('Scout, the assistant. Click him for help; right-click for more.') ?? 'Scout, the assistant. Click him for help; right-click for more.'} onMouseDown={(event) => event.preventDefault()} onPointerDown={(event) => {
      if (event.button !== 0) return;
      event.currentTarget.setPointerCapture(event.pointerId);
      draggedRef.current = false;
      runtime(view)?.assistant?.scout?.()?.dragStart?.(event.clientX, event.clientY);
    }} onPointerMove={(event) => {
      if (runtime(view)?.assistant?.scout?.()?.dragMove?.(event.clientX, event.clientY)) draggedRef.current = true;
    }} onPointerUp={() => {
      const result = runtime(view)?.assistant?.scout?.()?.dragEnd?.(true);
      draggedRef.current = !!result?.moved;
    }} onPointerCancel={() => {
      runtime(view)?.assistant?.scout?.()?.dragEnd?.(false);
      draggedRef.current = false;
    }} onClick={(event) => {
      if (draggedRef.current) { event.preventDefault(); draggedRef.current = false; return; }
      runtime(view)?.assistant?.activate?.();
    }} onKeyDown={(event) => {
      if (event.key === 'ContextMenu' || (event.key === 'F10' && event.shiftKey)) {
        event.preventDefault();
        const rect = avatarRef.current?.getBoundingClientRect();
        if (rect) runtime(view)?.assistant?.menuAt?.(rect.left + rect.width / 2, rect.top + rect.height / 2);
      } else if (event.key === 'Escape') runtime(view)?.assistant?.scout?.()?.stop?.();
    }} onContextMenu={(event) => { event.preventDefault(); runtime(view)?.assistant?.menuAt?.(event.clientX, event.clientY); }} />
    <div ref={balloonRef} className="scout-balloon" role="status" onClick={(event) => { if (!(event.target as HTMLElement).closest('.scout-choices button') && !choices.length) runtime(view)?.assistant?.scout?.()?.closeBalloon?.(); }} onKeyDown={(event) => { if (event.key === 'Escape') runtime(view)?.assistant?.scout?.()?.closeBalloon?.(); }}>
      <div ref={textRef} className="scout-text" data-react-owned="scout-text">{speechText}</div>
      <div ref={choicesRef} className="scout-choices" role="group" hidden={!choicesVisible}>
        {choices.map((label, index) => <button key={`${index}:${label}`} type="button" data-react-owned="scout-choice" onClick={() => runtime(view)?.assistant?.scout?.()?.choose?.(index)}>{label}</button>)}
      </div>
    </div>
  </div>, mount);
}
