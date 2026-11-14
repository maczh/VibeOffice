import { useEffect, useRef, useState } from 'react';
import { createPortal } from 'react-dom';

interface NotesAdapter {
  ed: HTMLDivElement;
  editor?: HTMLDivElement;
  initialized?: boolean;
  reactOwned?: boolean;
  reactInput?(editor: HTMLDivElement): void;
  reactPaste?(event: ClipboardEvent): void;
  reactFocus?(): void;
  attachReactEditor?(editor: HTMLDivElement): () => void;
}
interface NotesRuntime {
  pres?: { slides: Array<{ notes?: string; isMaster?: boolean }> };
  ed?: { idx: number };
  panes?: { notes?: NotesAdapter };
  te?: { active(): boolean; end(): void };
  bus?: { on(event: string, listener: () => void): () => void };
}
function runtime(): NotesRuntime | undefined { return (window as Window & { L?: NotesRuntime }).L; }

/** React-owned speaker-notes editor, synchronized through the existing Notes history adapter. */
export function PptxNotesLayer({ host }: { host: HTMLElement }) {
  const [mount, setMount] = useState<HTMLElement | null>(null);
  const [text, setText] = useState('');
  const [editable, setEditable] = useState(false);
  const editorRef = useRef<HTMLDivElement>(null);
  useEffect(() => {
    let adapter: HTMLDivElement | null = null;
    let slot: HTMLElement | null = null;
    let off: (() => void) | undefined;
    const sync = () => {
      const L = runtime();
      const slide = L?.pres?.slides[L.ed?.idx ?? 0];
      setText(slide && !slide.isMaster ? slide.notes || '' : '');
      setEditable(!!slide && !slide.isMaster);
    };
    const attach = () => {
      const target = host.querySelector<HTMLElement>('#react-notes-root');
      const input = host.querySelector<HTMLDivElement>('#notes .notes-in');
      if (!target || !input || target === slot) return;
      slot = target; adapter = input;
      slot.classList.add('vibeoffice-react-owned');
      adapter.classList.add('notes-engine-adapter');
      adapter.setAttribute('aria-hidden', 'true');
      setMount(slot);
      off = runtime()?.bus?.on('pptx-notes-render', sync);
      sync();
    };
    const observer = new MutationObserver(attach);
    observer.observe(host, { childList: true, subtree: true });
    attach();
    return () => { observer.disconnect(); off?.(); slot?.classList.remove('vibeoffice-react-owned'); adapter?.classList.remove('notes-engine-adapter'); adapter?.removeAttribute('aria-hidden'); };
  }, [host]);
  useEffect(() => {
    if (!mount) return;
    let disposed = false;
    let timer = 0;
    let detach: (() => void) | undefined;
    const attachWhenReady = () => {
      const notes = runtime()?.panes?.notes;
      const editor = editorRef.current;
      if (notes?.initialized && editor?.isConnected && notes.attachReactEditor) {
        detach = notes.attachReactEditor(editor);
        return;
      }
      if (!disposed) timer = window.setTimeout(attachWhenReady, 16);
    };
    attachWhenReady();
    return () => { disposed = true; window.clearTimeout(timer); detach?.(); };
  }, [mount]);
  if (!mount) return null;
  const adapter = host.querySelector<HTMLDivElement>('#notes .notes-in');
  if (!adapter) return null;
  return createPortal(<div ref={editorRef} className="notes-in" contentEditable={editable} suppressContentEditableWarning role="textbox" aria-multiline="true" aria-label="Speaker notes" data-ph={adapter.dataset.ph || ''} spellCheck={adapter.spellcheck} onInput={(event) => runtime()?.panes?.notes?.reactInput?.(event.currentTarget)} onPaste={(event) => runtime()?.panes?.notes?.reactPaste?.(event.nativeEvent)} onFocus={() => runtime()?.panes?.notes?.reactFocus?.()}>
    {text}
  </div>, mount, 'pptx-notes');
}
