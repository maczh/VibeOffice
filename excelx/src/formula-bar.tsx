import { useEffect, useState, type KeyboardEvent } from 'react';
import { createPortal } from 'react-dom';

interface FormulaEditor {
  active: boolean;
  begin(mode: 'edit', text?: string | null, options?: { surface?: 'bar' }): void;
  commit(options?: { stay?: boolean; keepFocus?: boolean }): boolean;
  cancel(): void;
}
interface FormulaRuntime {
  editor?: FormulaEditor;
  ui?: { exec(command: string): void };
  bus?: { on(event: string, listener: () => void): () => void };
}
function runtime(): FormulaRuntime | undefined { return (window as Window & { L?: FormulaRuntime }).L; }

/** React-owned Excelx formula bar controls backed by the engine's existing edit contract. */
export function ExcelxFormulaBarLayer({ host }: { host: HTMLElement }) {
  const [mount, setMount] = useState<HTMLElement | null>(null);
  const [buttonMount, setButtonMount] = useState<HTMLElement | null>(null);
  const [nameMount, setNameMount] = useState<HTMLElement | null>(null);
  const [value, setValue] = useState('');
  const [editing, setEditing] = useState(false);
  const [multi, setMulti] = useState(false);
  useEffect(() => {
    let adapter: HTMLTextAreaElement | null = null;
    let adapters: HTMLElement[] = [];
    let ownedSlots: HTMLElement[] = [];
    let offs: (() => void)[] = [];
    const sync = () => {
      if (!adapter) return;
      setValue(adapter.value);
      setEditing(host.querySelector('#fbar')?.classList.contains('editing') ?? false);
      setMulti(adapter.classList.contains('multi'));
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
      const nameButton = host.querySelector<HTMLElement>('#namebox-dd');
      ownedSlots = [slot, controls, ...(nameSlot ? [nameSlot] : [])];
      adapters = [nextAdapter, ...(nativeControls ? [nativeControls] : []), ...(nameAdapter ? [nameAdapter] : []), ...(nameButton ? [nameButton] : [])];
      ownedSlots.forEach((element) => element.classList.add('vibeoffice-react-owned'));
      adapters.forEach((element) => element.classList.add('fb-engine-adapter'));
      nativeControls?.setAttribute('aria-hidden', 'true');
      nameAdapter?.setAttribute('aria-hidden', 'true');
      nameButton?.setAttribute('aria-hidden', 'true');
      if (nameSlot) setNameMount(nameSlot);
      const L = runtime();
      offs = ['selection', 'changed', 'mode'].map((event) => L?.bus?.on(event, sync)).filter((off): off is () => void => !!off);
      adapter.addEventListener('input', sync);
      sync();
    };
    const observer = new MutationObserver(() => { attach(); sync(); });
    observer.observe(host, { childList: true, subtree: true, attributes: true, attributeFilter: ['class'] });
    attach();
    return () => { offs.forEach((off) => off()); observer.disconnect(); adapter?.removeEventListener('input', sync); ownedSlots.forEach((element) => element.classList.remove('vibeoffice-react-owned')); adapters.forEach((element) => element.classList.remove('fb-engine-adapter')); [host.querySelector('#fb-buttons'), host.querySelector('#namebox'), host.querySelector('#namebox-dd')].forEach((element) => element?.removeAttribute('aria-hidden')); };
  }, [host]);
  if (!mount || !buttonMount) return null;
  const adapter = host.querySelector<HTMLTextAreaElement>('#fbinput');
  if (!adapter) return null;
  const editor = runtime()?.editor;
  const publishInput = (element: HTMLTextAreaElement) => {
    adapter.value = element.value;
    adapter.setSelectionRange(element.selectionStart, element.selectionEnd);
    adapter.dispatchEvent(new Event('input', { bubbles: true }));
    setValue(element.value);
  };
  const keyDown = (event: KeyboardEvent<HTMLTextAreaElement>) => {
    adapter.value = event.currentTarget.value;
    adapter.setSelectionRange(event.currentTarget.selectionStart, event.currentTarget.selectionEnd);
    const forwarded = new KeyboardEvent('keydown', { key: event.key, code: event.code, ctrlKey: event.ctrlKey, metaKey: event.metaKey, shiftKey: event.shiftKey, altKey: event.altKey, bubbles: true, cancelable: true });
    adapter.dispatchEvent(forwarded);
    if (forwarded.defaultPrevented) event.preventDefault();
    setValue(adapter.value);
    requestAnimationFrame(() => { if (document.activeElement === event.currentTarget) { event.currentTarget.value = adapter.value; event.currentTarget.setSelectionRange(adapter.selectionStart, adapter.selectionEnd); } });
  };
  return <>
    {createPortal(<div className="fb-btns fb-react-btns" aria-label="Formula bar controls">
      <button className="fb-x" type="button" aria-label="Cancel" title="Cancel" hidden={!editing} onPointerDown={(event) => { event.preventDefault(); editor?.cancel(); }}><svg width="13" height="13" viewBox="0 0 13 13" aria-hidden="true"><path d="M2 2l9 9M11 2l-9 9" stroke="#c0281a" strokeWidth="2.2" /></svg></button>
      <button className="fb-ok" type="button" aria-label="Enter" title="Enter" hidden={!editing} onPointerDown={(event) => { event.preventDefault(); editor?.commit({ stay: true }); }}><svg width="13" height="13" viewBox="0 0 13 13" aria-hidden="true"><path d="M1.5 7l3.5 3.5L11.5 2.5" fill="none" stroke="#1b8a2a" strokeWidth="2.2" /></svg></button>
      <button type="button" aria-label="Insert Function" title="Insert Function" onPointerDown={(event) => { event.preventDefault(); runtime()?.ui?.exec('insertFunction'); }}><svg width="18" height="14" viewBox="0 0 18 14" aria-hidden="true"><text x="1" y="11" fontFamily="Times New Roman,Tinos,serif" fontStyle="italic" fontSize="13" fill="#1a3f8a">f</text><text x="6" y="11" fontFamily="Times New Roman,Tinos,serif" fontStyle="italic" fontSize="11" fill="#1a3f8a">x</text></svg></button>
    </div>, buttonMount, 'excelx-formula-bar-controls')}
    {createPortal(
      <textarea className={`fb-input${multi ? ' multi' : ''}`} rows={1} aria-label="Formula Bar" spellCheck={false} autoComplete="off" value={value} onFocus={() => { if (!editor?.active) editor?.begin('edit', null, { surface: 'bar' }); }} onChange={(event) => publishInput(event.currentTarget)} onKeyDown={keyDown} onDoubleClick={() => { adapter.classList.toggle('multi'); setMulti(adapter.classList.contains('multi')); }} />
    , mount, 'excelx-formula-bar-input')}
    {nameMount && <ExcelxNameBoxLayer host={host} mount={nameMount} />}
  </>;
}

function ExcelxNameBoxLayer({ host, mount }: { host: HTMLElement; mount: HTMLElement }) {
  const [value, setValue] = useState('');
  useEffect(() => {
    const adapter = host.querySelector<HTMLInputElement>('#namebox');
    if (!adapter) return;
    const sync = () => setValue(adapter.value);
    const L = runtime();
    const offs = ['selection', 'sheet', 'book', 'changed'].map((event) => L?.bus?.on(event, sync)).filter((off): off is () => void => !!off);
    adapter.addEventListener('input', sync);
    sync();
    return () => { offs.forEach((off) => off()); adapter.removeEventListener('input', sync); };
  }, [host]);
  const adapter = host.querySelector<HTMLInputElement>('#namebox');
  const button = host.querySelector<HTMLButtonElement>('#namebox-dd');
  if (!adapter || !button) return null;
  const update = (input: HTMLInputElement) => { adapter.value = input.value; setValue(input.value); };
  const keyDown = (event: KeyboardEvent<HTMLInputElement>) => {
    adapter.value = event.currentTarget.value;
    const forwarded = new KeyboardEvent('keydown', { key: event.key, code: event.code, ctrlKey: event.ctrlKey, metaKey: event.metaKey, shiftKey: event.shiftKey, altKey: event.altKey, bubbles: true, cancelable: true });
    adapter.dispatchEvent(forwarded);
    if (forwarded.defaultPrevented) event.preventDefault();
    setValue(adapter.value);
  };
  return createPortal(<>
    <input aria-label="Name Box" autoComplete="off" spellCheck={false} value={value} onFocus={(event) => { if (runtime()?.editor?.active && !runtime()?.editor?.commit({ keepFocus: true })) { event.currentTarget.blur(); return; } event.currentTarget.select(); }} onChange={(event) => update(event.currentTarget)} onKeyDown={keyDown} />
    <button type="button" aria-label="Names" title="Names" onPointerDown={(event) => { event.preventDefault(); button.dispatchEvent(new PointerEvent('pointerdown', { bubbles: true, cancelable: true, pointerType: event.pointerType, clientX: event.clientX, clientY: event.clientY })); }}><span className="dd-arrow" /></button>
  </>, mount, 'excelx-name-box');
}
