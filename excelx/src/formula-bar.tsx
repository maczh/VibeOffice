import { useEffect, useRef, useState } from 'react';
import { createPortal } from 'react-dom';

interface FormulaEditor {
  active: boolean;
  initialized?: boolean;
  begin(mode: 'edit', text?: string | null, options?: { surface?: 'bar' }): void;
  commit(options?: { stay?: boolean; keepFocus?: boolean }): boolean;
  cancel(): void;
  reactOwned?: boolean;
  reactBarFocus?(): void;
  reactBarKeyDown?(event: KeyboardEvent): void;
  reactBarInput?(): void;
  reactBarPointerDown?(): void;
  reactBarDoubleClick?(): void;
  attachReactFormulaBar?(input: HTMLTextAreaElement): () => void;
}
interface FormulaRuntime {
  editor?: FormulaEditor;
  app?: {
    nameBoxInitialized?: boolean;
    nameBoxReactOwned?: boolean;
    reactNameBoxFocus?(input: HTMLInputElement): void;
    reactNameBoxKeyDown?(input: HTMLInputElement, event: KeyboardEvent): void;
    reactNameBoxOpenMenu?(): void;
    attachReactNameBox?(input: HTMLInputElement): () => void;
  };
  ui?: { exec(command: string): void };
  bus?: { on(event: string, listener: () => void): () => void };
}
function runtime(): FormulaRuntime | undefined { return (window as Window & { L?: FormulaRuntime }).L; }

/** React-owned Excelx formula bar controls backed by the engine's existing edit contract. */
export function ExcelxFormulaBarLayer({ host }: { host: HTMLElement }) {
  const [mount, setMount] = useState<HTMLElement | null>(null);
  const [buttonMount, setButtonMount] = useState<HTMLElement | null>(null);
  const [nameMount, setNameMount] = useState<HTMLElement | null>(null);
  const inputRef = useRef<HTMLTextAreaElement>(null);
  const valueFrame = useRef<number | null>(null);
  const [value, setValue] = useState('');
  const [editing, setEditing] = useState(false);
  const [multi, setMulti] = useState(false);
  useEffect(() => () => { if (valueFrame.current !== null) cancelAnimationFrame(valueFrame.current); }, []);
  useEffect(() => {
    let adapter: HTMLTextAreaElement | null = null;
    let adapters: HTMLElement[] = [];
    let ownedSlots: HTMLElement[] = [];
    let offs: (() => void)[] = [];
    const sync = () => {
      if (!adapter) return;
      const input = inputRef.current;
      setValue(input?.value ?? adapter.value);
      setEditing(host.querySelector('#fbar')?.classList.contains('editing') ?? false);
      setMulti((input ?? adapter).classList.contains('multi'));
    };
    const attach = () => {
      if (adapter?.isConnected) return;
      const slot = host.querySelector<HTMLElement>('#react-fbar');
      const controls = host.querySelector<HTMLElement>('#react-fbar-buttons');
      const nextAdapter = host.querySelector<HTMLTextAreaElement>('#fbinput');
      if (!slot || !controls || !nextAdapter) return;
      adapter = nextAdapter;
      setMount(slot);
      setButtonMount(controls);
      const nativeControls = host.querySelector<HTMLElement>('#fb-buttons');
      const nameSlot = host.querySelector<HTMLElement>('#react-namebox');
      const nameAdapter = host.querySelector<HTMLElement>('#namebox');
      ownedSlots = [slot, controls, ...(nameSlot ? [nameSlot] : [])];
      adapters = [nextAdapter, ...(nativeControls ? [nativeControls] : []), ...(nameAdapter ? [nameAdapter] : [])];
      ownedSlots.forEach((element) => element.classList.add('vibeoffice-react-owned'));
      adapters.forEach((element) => element.classList.add('fb-engine-adapter'));
      nativeControls?.setAttribute('aria-hidden', 'true');
      nameAdapter?.setAttribute('aria-hidden', 'true');
      if (nameSlot) setNameMount(nameSlot);
      const L = runtime();
      offs = ['selection', 'changed', 'mode'].map((event) => L?.bus?.on(event, sync)).filter((off): off is () => void => !!off);
      adapter.addEventListener('input', sync);
      sync();
    };
    const observer = new MutationObserver(() => { attach(); sync(); });
    observer.observe(host, { childList: true, subtree: true, attributes: true, attributeFilter: ['class'] });
    attach();
    return () => { offs.forEach((off) => off()); observer.disconnect(); adapter?.removeEventListener('input', sync); ownedSlots.forEach((element) => element.classList.remove('vibeoffice-react-owned')); adapters.forEach((element) => element.classList.remove('fb-engine-adapter')); [host.querySelector('#fb-buttons'), host.querySelector('#namebox')].forEach((element) => element?.removeAttribute('aria-hidden')); };
  }, [host]);
  useEffect(() => {
    if (!mount) return;
    let disposed = false;
    let timer = 0;
    let detach: (() => void) | undefined;
    const attachWhenReady = () => {
      const editor = runtime()?.editor;
      const input = inputRef.current;
      if (editor?.initialized && input?.isConnected && editor.attachReactFormulaBar) {
        detach = editor.attachReactFormulaBar(input);
        return;
      }
      if (!disposed) timer = window.setTimeout(attachWhenReady, 16);
    };
    attachWhenReady();
    return () => { disposed = true; window.clearTimeout(timer); detach?.(); };
  }, [mount]);
  if (!mount || !buttonMount) return null;
  const adapter = host.querySelector<HTMLTextAreaElement>('#fbinput');
  if (!adapter) return null;
  const editor = runtime()?.editor;
  return <>
    {createPortal(<div className="fb-btns fb-react-btns" aria-label="Formula bar controls">
      <button className="fb-x" type="button" aria-label="Cancel" title="Cancel" hidden={!editing} onPointerDown={(event) => { event.preventDefault(); editor?.cancel(); }}><svg width="13" height="13" viewBox="0 0 13 13" aria-hidden="true"><path d="M2 2l9 9M11 2l-9 9" stroke="#c0281a" strokeWidth="2.2" /></svg></button>
      <button className="fb-ok" type="button" aria-label="Enter" title="Enter" hidden={!editing} onPointerDown={(event) => { event.preventDefault(); editor?.commit({ stay: true }); }}><svg width="13" height="13" viewBox="0 0 13 13" aria-hidden="true"><path d="M1.5 7l3.5 3.5L11.5 2.5" fill="none" stroke="#1b8a2a" strokeWidth="2.2" /></svg></button>
      <button type="button" aria-label="Insert Function" title="Insert Function" onPointerDown={(event) => { event.preventDefault(); runtime()?.ui?.exec('insertFunction'); }}><svg width="18" height="14" viewBox="0 0 18 14" aria-hidden="true"><text x="1" y="11" fontFamily="Times New Roman,Tinos,serif" fontStyle="italic" fontSize="13" fill="#1a3f8a">f</text><text x="6" y="11" fontFamily="Times New Roman,Tinos,serif" fontStyle="italic" fontSize="11" fill="#1a3f8a">x</text></svg></button>
    </div>, buttonMount, 'excelx-formula-bar-controls')}
    {createPortal(
      <textarea ref={inputRef} className={`fb-input${multi ? ' multi' : ''}`} rows={1} aria-label="Formula Bar" spellCheck={false} autoComplete="off" value={value} onFocus={() => editor?.reactBarFocus?.()} onChange={(event) => setValue(event.currentTarget.value)} onInput={() => editor?.reactBarInput?.()} onKeyDown={(event) => { editor?.reactBarKeyDown?.(event.nativeEvent); const input = event.currentTarget; if (valueFrame.current !== null) cancelAnimationFrame(valueFrame.current); valueFrame.current = requestAnimationFrame(() => { valueFrame.current = null; if (input.isConnected) setValue(input.value); }); }} onPointerDown={() => editor?.reactBarPointerDown?.()} onDoubleClick={(event) => { editor?.reactBarDoubleClick?.(); setMulti(event.currentTarget.classList.contains('multi')); }} />
    , mount, 'excelx-formula-bar-input')}
    {nameMount && <ExcelxNameBoxLayer host={host} mount={nameMount} />}
  </>;
}

function ExcelxNameBoxLayer({ host, mount }: { host: HTMLElement; mount: HTMLElement }) {
  const [value, setValue] = useState('');
  const inputRef = useRef<HTMLInputElement>(null);
  const valueFrame = useRef<number | null>(null);
  useEffect(() => () => { if (valueFrame.current !== null) cancelAnimationFrame(valueFrame.current); }, []);
  useEffect(() => {
    const adapter = host.querySelector<HTMLInputElement>('#namebox');
    if (!adapter) return;
    const sync = () => setValue(inputRef.current?.value ?? adapter.value);
    const L = runtime();
    const offs = ['selection', 'sheet', 'book', 'changed'].map((event) => L?.bus?.on(event, sync)).filter((off): off is () => void => !!off);
    sync();
    return () => { offs.forEach((off) => off()); };
  }, [host]);
  useEffect(() => {
    let disposed = false;
    let timer = 0;
    let detach: (() => void) | undefined;
    const attachWhenReady = () => {
      const app = runtime()?.app;
      const input = inputRef.current;
      if (app?.nameBoxInitialized && input?.isConnected && app.attachReactNameBox) {
        detach = app.attachReactNameBox(input);
        return;
      }
      if (!disposed) timer = window.setTimeout(attachWhenReady, 16);
    };
    attachWhenReady();
    return () => { disposed = true; window.clearTimeout(timer); detach?.(); };
  }, [mount]);
  const adapter = host.querySelector<HTMLInputElement>('#namebox');
  if (!adapter) return null;
  const update = (input: HTMLInputElement) => setValue(input.value);
  return createPortal(<>
    <input ref={inputRef} aria-label="Name Box" autoComplete="off" spellCheck={false} value={value} onChange={(event) => update(event.currentTarget)} onFocus={(event) => runtime()?.app?.reactNameBoxFocus?.(event.currentTarget)} onKeyDown={(event) => { const input = event.currentTarget; runtime()?.app?.reactNameBoxKeyDown?.(input, event.nativeEvent); if (valueFrame.current !== null) cancelAnimationFrame(valueFrame.current); valueFrame.current = requestAnimationFrame(() => { valueFrame.current = null; if (input.isConnected) setValue(input.value); }); }} />
    <button type="button" aria-label="Names" title="Names" onPointerDown={(event) => { event.preventDefault(); runtime()?.app?.reactNameBoxOpenMenu?.(); }}><span className="dd-arrow" /></button>
  </>, mount, 'excelx-name-box');
}
