import { createPortal } from 'react-dom';
import { useEffect, useRef, useState } from 'react';

interface ScoutMount { element: HTMLElement; avatarElement: HTMLElement; balloonElement: HTMLElement; textElement: HTMLElement; choicesElement: HTMLElement; onText(text: string): void; onChoices(choices: string[], visible: boolean): void }
interface ScoutAssistantRuntime {
  assistant?: { shown?(): boolean; scout?(): { choose?(index: number): unknown; closeBalloon?(): unknown } | null; attachReactHost?(factory: () => ScoutMount): () => void };
}
function runtime(): ScoutAssistantRuntime | undefined { return (window as Window & { L?: ScoutAssistantRuntime }).L; }

/** React owns Scout's scoped host, avatar slot, and speech surface; the existing animation/service API remains authoritative. */
export function ScoutAssistant({ host }: { host: HTMLElement }) {
  const rootRef = useRef<HTMLDivElement>(null);
  const avatarRef = useRef<HTMLDivElement>(null);
  const balloonRef = useRef<HTMLDivElement>(null);
  const textRef = useRef<HTMLDivElement>(null);
  const choicesRef = useRef<HTMLDivElement>(null);
  const [mount, setMount] = useState<HTMLElement | null>(null);
  const [speechText, setSpeechText] = useState('');
  const [choices, setChoices] = useState<string[]>([]);
  const [choicesVisible, setChoicesVisible] = useState(false);

  useEffect(() => {
    let disposed = false;
    let timer = 0;
    const findMount = () => {
      if (disposed) return;
      const root = host.querySelector<HTMLElement>('.vibeoffice-legacy-root[data-vibeoffice-engine-root]');
      if (root) setMount(root);
      else timer = window.setTimeout(findMount, 25);
    };
    findMount();
    return () => { disposed = true; window.clearTimeout(timer); };
  }, [host]);

  useEffect(() => {
    if (!mount) return;
    let disposed = false;
    let timer = 0;
    let detach: (() => void) | undefined;
    const attach = () => {
      if (disposed) return;
      const assistant = runtime()?.assistant;
      const element = rootRef.current, avatarElement = avatarRef.current, balloonElement = balloonRef.current, textElement = textRef.current, choicesElement = choicesRef.current;
      if (!assistant?.attachReactHost || !element || !avatarElement || !balloonElement || !textElement || !choicesElement) { timer = window.setTimeout(attach, 25); return; }
      detach = assistant.attachReactHost(() => ({ element, avatarElement, balloonElement, textElement, choicesElement, onText: setSpeechText, onChoices: (items, visible) => { setChoices(items); setChoicesVisible(visible); } }));
    };
    attach();
    return () => { disposed = true; window.clearTimeout(timer); detach?.(); };
  }, [mount]);

  if (!mount) return null;
  return createPortal(<div ref={rootRef} className="scout-agent is-desktop" data-react-owned="scout-assistant">
    <div ref={avatarRef} className="scout-avatar" />
    <div ref={balloonRef} className="scout-balloon" role="status" onClick={(event) => { if (!(event.target as HTMLElement).closest('.scout-choices button') && !choices.length) runtime()?.assistant?.scout?.()?.closeBalloon?.(); }} onKeyDown={(event) => { if (event.key === 'Escape') runtime()?.assistant?.scout?.()?.closeBalloon?.(); }}>
      <div ref={textRef} className="scout-text" data-react-owned="scout-text">{speechText}</div>
      <div ref={choicesRef} className="scout-choices" role="group" hidden={!choicesVisible}>
        {choices.map((label, index) => <button key={`${index}:${label}`} type="button" data-react-owned="scout-choice" onClick={() => runtime()?.assistant?.scout?.()?.choose?.(index)}>{label}</button>)}
      </div>
    </div>
  </div>, mount);
}
