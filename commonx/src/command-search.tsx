import { useEffect, useId, useRef, useState, type ChangeEvent, type KeyboardEvent, type ReactNode } from 'react';
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

function highlightedLabel(label: string, query: string): ReactNode[] {
  const tokens = query.toLowerCase().split(/[^\p{L}\p{N}%$&+]+/u).filter(Boolean);
  const lower = label.toLowerCase();
  const highlighted = new Array(label.length).fill(false) as boolean[];
  for (const token of tokens) {
    const escaped = token.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const expression = new RegExp(`(^|[^a-z0-9])(${escaped})`, 'g');
    let match: RegExpExecArray | null;
    while ((match = expression.exec(lower))) {
      const prefix = match[1] || '', matched = match[2] || '';
      const start = match.index + prefix.length;
      for (let i = start; i < start + matched.length; i++) highlighted[i] = true;
    }
  }
  const parts: ReactNode[] = [];
  for (let start = 0; start < label.length;) {
    const marked = highlighted[start];
    let end = start + 1;
    while (end < label.length && highlighted[end] === marked) end++;
    const text = label.slice(start, end);
    parts.push(marked ? <b key={`match-${start}`}>{text}</b> : text);
    start = end;
  }
  return parts.length ? parts : [label];
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
      detail: { query: value.trim(), layout: document.documentElement.dataset.layout || 'bars', respond: (items: CommandResult[]) => { response = items; } },
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
    const input = inputRef.current;
    const engine = input?.closest('.vibe-editor')?.querySelector<HTMLElement>('.vibe-editor__engine');
    engine?.dispatchEvent(new CustomEvent('vibeoffice-react-command', { bubbles: true, detail: { type: 'vibeoffice-react-command', command: result.id, rememberSearch: true } }));
    setOpen(false);
    setQuery('');
    setResults([]);
    setActive(-1);
  };

  const moveActive = (direction: -1 | 1) => {
    setActive((current) => {
      let index = current < 0 ? (direction > 0 ? -1 : 0) : current;
      for (let step = 0; step < results.length; step++) {
        index = (index + direction + results.length) % results.length;
        if (results[index]?.enabled) return index;
      }
      return -1;
    });
  };

  const onChange = (event: ChangeEvent<HTMLInputElement>) => search(event.currentTarget.value);
  const onKeyDown = (event: KeyboardEvent<HTMLInputElement>) => {
    if (event.key === 'ArrowDown' && results.length) { event.preventDefault(); moveActive(1); }
    else if (event.key === 'ArrowUp' && results.length) { event.preventDefault(); moveActive(-1); }
    else if (event.key === 'Enter') { event.preventDefault(); run(active < 0 ? results.findIndex((result) => result.enabled) : active); }
    else if (event.key === 'Escape') { setOpen(false); setQuery(''); setResults([]); }
  };

  const activeResultId = active >= 0 && results[active]?.enabled ? `${id}-option-${active}` : undefined;
  return <div className="vibe-command-search">
    <input ref={inputRef} id={id} type="search" role="combobox" aria-label={office.translate('Search commands')} aria-expanded={open && results.length > 0} aria-controls={`${id}-results`} aria-activedescendant={open ? activeResultId : undefined} aria-autocomplete="list" placeholder={office.translate('Search commands')} value={query} onChange={onChange} onFocus={() => { setOpen(true); search(query); }} onBlur={() => window.setTimeout(() => setOpen(false), 100)} onKeyDown={onKeyDown} />
    {open && results.length > 0 && <ul id={`${id}-results`} className="vibe-command-search__results" role="listbox">{results.map((result, index) => <li id={`${id}-option-${index}`} key={`${result.id || result.label}-${index}`} role="option" aria-selected={index === active && result.enabled} aria-disabled={!result.enabled}><button type="button" tabIndex={-1} disabled={!result.enabled} onMouseDown={(event) => event.preventDefault()} onClick={() => run(index)}><span>{highlightedLabel(result.label, query)}</span>{result.where && <small>{result.where}</small>}</button></li>)}</ul>}
  </div>;
}
