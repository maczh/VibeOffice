import { useEffect, useState } from 'react';
import { createPortal } from 'react-dom';

interface NotesAdapter { ed: HTMLDivElement }
interface NotesRuntime {
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
  useEffect(() => {
    let adapter: HTMLDivElement | null = null;
    let slot: HTMLElement | null = null;
    let off: (() => void) | undefined;
    const sync = () => {
      if (!adapter) return;
      setText(adapter.textContent || '');
      setEditable(adapter.contentEditable !== 'false');
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
  if (!mount) return null;
  const adapter = host.querySelector<HTMLDivElement>('#notes .notes-in');
  if (!adapter) return null;
  const input = (element: HTMLDivElement) => {
    adapter.innerText = element.innerText;
    adapter.dispatchEvent(new Event('input', { bubbles: true }));
  };
  return createPortal(<div className="notes-in" contentEditable={editable} suppressContentEditableWarning role="textbox" aria-multiline="true" aria-label="Speaker notes" data-ph={adapter.dataset.ph || ''} spellCheck={adapter.spellcheck} onInput={(event) => input(event.currentTarget)} onFocus={() => { const textEdit = runtime()?.te; if (textEdit?.active()) textEdit.end(); }} onPaste={(event) => { event.preventDefault(); document.execCommand('insertText', false, event.clipboardData.getData('text/plain')); }}>
    {text}
  </div>, mount, 'pptx-notes');
}
