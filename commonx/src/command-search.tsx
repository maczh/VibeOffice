import { useEffect, useId, useRef, useState, type ChangeEvent, type KeyboardEvent } from 'react';
import { useVibeOffice } from './react-bridge.js';

interface CommandResult {
  id?: string;
  label: string;
  where?: string;
  enabled: boolean;
  entry?: { run?: () => unknown };
}

interface SearchRequest extends CustomEvent {
  detail: {
    query: string;
    layout: string;
    respond: (results: CommandResult[]) => void;
  };
}

/** React-owned command search backed by the mounted editor's command catalog and matching rules. */
export function CommandSearch() {
  const office = useVibeOffice();
  const id = useId();
  const inputRef = useRef<HTMLInputElement>(null);
  const [query, setQuery] = useState('');
  const [results, setResults] = useState<CommandResult[]>([]);
  const [active, setActive] = useState(-1);
  const [open, setOpen] = useState(false);

  useEffect(() => {
    const focusSearch = (event: globalThis.KeyboardEvent) => {
      if (!event.altKey || event.ctrlKey || event.shiftKey || event.metaKey || (event.code !== 'KeyQ' && event.key.toLowerCase() !== 'q')) return;
      if (!inputRef.current?.closest('.vibe-editor')) return;
      event.preventDefault();
      event.stopPropagation();
      inputRef.current.focus();
    };
    document.addEventListener('keydown', focusSearch, true);
    return () => document.removeEventListener('keydown', focusSearch, true);
  }, []);

  const search = (value: string) => {
    setQuery(value);
    setActive(-1);
    const input = document.activeElement;
    const engine = input instanceof HTMLElement
      ? input.closest('.vibe-editor')?.querySelector<HTMLElement>('.vibe-editor__engine')
      : null;
    if (!engine) { setResults([]); return; }
    let response: CommandResult[] = [];
    engine.dispatchEvent(new CustomEvent('vibeoffice-react-search', {
      bubbles: true,
      detail: { query: value.trim(), layout: 'bars', respond: (items: CommandResult[]) => { response = items; } },
    }) as SearchRequest);
    setResults(response);
  };

  const run = (index: number) => {
    const result = results[index];
    if (!result?.enabled) return;
    if (!result.id) {
      result.entry?.run?.();
      setOpen(false);
      setQuery('');
      setResults([]);
      setActive(-1);
      return;
    }
    const input = document.getElementById(id);
    const engine = input?.closest('.vibe-editor')?.querySelector<HTMLElement>('.vibe-editor__engine');
    engine?.dispatchEvent(new CustomEvent('vibeoffice-react-command', { bubbles: true, detail: { type: 'vibeoffice-react-command', command: result.id } }));
    setOpen(false);
    setQuery('');
    setResults([]);
    setActive(-1);
  };

  const onChange = (event: ChangeEvent<HTMLInputElement>) => search(event.currentTarget.value);
  const onKeyDown = (event: KeyboardEvent<HTMLInputElement>) => {
    if (event.key === 'ArrowDown' && results.length) { event.preventDefault(); setActive((index) => (index + 1) % results.length); }
    else if (event.key === 'ArrowUp' && results.length) { event.preventDefault(); setActive((index) => (index <= 0 ? results.length - 1 : index - 1)); }
    else if (event.key === 'Enter') { event.preventDefault(); run(active < 0 ? results.findIndex((result) => result.enabled) : active); }
    else if (event.key === 'Escape') { setOpen(false); setQuery(''); setResults([]); }
  };

  return <div className="vibe-command-search">
    <input ref={inputRef} id={id} type="search" role="combobox" aria-label={office.translate('Search commands')} aria-expanded={open && results.length > 0} aria-controls={`${id}-results`} aria-autocomplete="list" placeholder={office.translate('Search commands')} value={query} onChange={onChange} onFocus={() => { setOpen(true); search(query); }} onBlur={() => window.setTimeout(() => setOpen(false), 100)} onKeyDown={onKeyDown} />
    {open && results.length > 0 && <ul id={`${id}-results`} className="vibe-command-search__results" role="listbox">{results.map((result, index) => <li key={`${result.id || result.label}-${index}`} role="option" aria-selected={index === active} aria-disabled={!result.enabled}><button type="button" tabIndex={-1} disabled={!result.enabled} onMouseDown={(event) => event.preventDefault()} onClick={() => run(index)}><span>{result.label}</span>{result.where && <small>{result.where}</small>}</button></li>)}</ul>}
  </div>;
}
