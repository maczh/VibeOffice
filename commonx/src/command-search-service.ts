type SearchEntry = {
  id?: string;
  label: string;
  alias: string;
  keys?: string;
  path: string[] | null;
  tip: string;
  placed?: boolean;
  toggle?: boolean;
  run?: () => unknown;
  enabledFn?: () => boolean;
  own?: string[];
  also?: string[];
  icon?: string;
};

type SearchResult = { id?: string; label: string; where: string; enabled: boolean; entry: SearchEntry };
type MatcherEntry = Pick<SearchEntry, 'label' | 'alias' | 'tip' | 'path'> & { own?: string[]; also?: string[] };
interface CommandMatcher { maxResults: number; maxRecent: number; words(value: string): string[]; scoreWord(query: string, entry: MatcherEntry): number }
type SearchRuntime = Window & { L?: Record<string, any>; VO?: { commandMatcher?: CommandMatcher; reactDirectHost?: { root?: HTMLElement } } };

/** Installs the shared command catalog/matcher used by the React command-search component. */
export function installCommandSearchService(runtime: Window): void {
  const legacy = (runtime as SearchRuntime).L;
  if (!legacy?.ui) throw new Error('The shared command search requires the editor UI service');
  const matcher = (runtime as SearchRuntime).VO?.commandMatcher;
  if (!matcher) throw new Error('The shared command search requires the common command matcher');
  const L = legacy;
  const ui = L.ui;
  const normalize = (value: unknown) => ui.stripAmp(String(value || '')).replace(/(\.\.\.|…)$/, '').trim();
  function commandPlaces() {
    const menu = new Map<string, string[]>(), ribbon = new Map<string, { where: string[]; label?: string }>(), actions: SearchEntry[] = [];
    const walkMenu = (items: any, path: string[]) => {
      let list: any;
      try { list = typeof items === 'function' ? items() : items; } catch { return; }
      for (const item of list || []) {
        if (typeof item === 'string') { if (item !== '-' && !menu.has(item)) menu.set(item, path); continue; }
        if (!item || item.custom) continue;
        if (item.cmd) { if (!menu.has(item.cmd)) menu.set(item.cmd, path); continue; }
        if (Array.isArray(item.sub)) walkMenu(item.sub, path.concat(normalize(item.label)));
        else if (item.run && item.label) actions.push({ label: normalize(item.label), alias: '', path, tip: '', run: item.run, enabledFn: item.enabled, icon: item.icon });
      }
    };
    for (const item of ui.menus || []) walkMenu(item.items, [normalize(item.label)]);
    const spec = L.ribbonSpec ? L.ribbonSpec() : null;
    if (spec) {
      const put = (id: string | undefined, where: string[], label?: string) => { if (id && !ribbon.has(id)) ribbon.set(id, { where, label: label && normalize(label) }); };
      const walk = (item: any, where: string[]) => {
        if (!item || item === '|') return;
        if (typeof item === 'string') return put(item, where);
        if (item.row) return item.row.forEach((child: any) => walk(child, where));
        put(item.cmd || item.split, where, item.label);
      };
      const tabs = spec.tabs.map((tab: any) => [tab, '']).concat(...(spec.contextual || []).map((context: any) => context.tabs.map((tab: any) => [tab, L.t(context.set) + ' ▸ '])));
      for (const [tab, prefix] of tabs as [any, string][]) for (const group of tab.groups) {
        const where = [prefix + normalize(tab.label), L.t(group.label)];
        group.items.forEach((item: any) => walk(item, where));
        put(group.launcher, where);
      }
      (spec.qat || []).forEach((item: any) => walk(item, [L.t('Quick Access Toolbar')]));
      for (const item of (spec.file ? spec.file() : [])) if (typeof item === 'string' && item !== '-') put(item, [normalize('&File')]);
    }
    const toolbar = new Map<string, string[]>();
    for (const element of L.$$('.toolbar[aria-label]')) for (const button of L.$$('[data-cmd]', element)) if (!toolbar.has(button.dataset.cmd)) toolbar.set(button.dataset.cmd, [L.t('{0} toolbar', element.getAttribute('aria-label'))]);
    return { menu, ribbon, toolbar, actions };
  }

  function entries(layout: string): SearchEntry[] {
    const { menu, ribbon, toolbar, actions } = commandPlaces();
    const out: SearchEntry[] = [], seen = new Set<string>();
    for (const id of Object.keys(ui.cmds)) {
      try {
        const command = ui.cmds[id];
        if (!command.label || typeof command.label !== 'string') continue;
        if (layout === 'ribbon' && id.startsWith('tb_')) continue;
        const ribbonPlace = ribbon.get(id), menuPlace = menu.get(id);
        const label = layout === 'ribbon' && ribbonPlace?.label ? ribbonPlace.label : normalize(command.label);
        const path = (layout === 'ribbon' ? ribbonPlace?.where : menuPlace || toolbar.get(id)) || null;
        const alias = [normalize(command.label), ui.plainAmp(command.label).replace(/(\.\.\.|…)$/, ''), command.keys || '', layout !== 'ribbon' && menuPlace && menuPlace.length > 1 ? menuPlace[menuPlace.length - 1] : ''].join(' ');
        const key = label.toLowerCase() + '|' + (path || []).join('/');
        if (seen.has(key)) continue;
        seen.add(key);
        out.push({ id, label, alias, keys: (command.keys || '').toLowerCase(), path, tip: normalize(command.tip || ''), placed: !!path, toggle: id.startsWith('tb_'), icon: command.icon });
      } catch { /* Stale commands from another mounted app must not break the active search. */ }
    }
    if (layout !== 'ribbon') for (const action of actions) {
      const key = action.label.toLowerCase() + '|' + (action.path || []).join('/');
      if (!seen.has(key)) { seen.add(key); out.push({ ...action, alias: action.path && action.path.length > 1 ? action.path[action.path.length - 1] || '' : '', placed: true }); }
    }
    return out;
  }

  const isEnabled = (entry: SearchEntry) => entry.id ? ui.enabled(entry.id) : !entry.enabledFn || !!entry.enabledFn();
  const searchCommands = (query: string, layout = runtime.document.documentElement.dataset.layout || 'bars'): SearchResult[] => {
    const queryWords = matcher.words(query);
    if (!queryWords.length) return [];
    const recent: string[] = L.store ? L.store.get('searchRecent', []) : [];
    const full = queryWords.join(' '), scored: { entry: SearchEntry; score: number; enabled: boolean }[] = [];
    for (const entry of entries(layout)) {
      let score = 0;
      for (const word of queryWords) { const value = matcher.scoreWord(word, entry); if (!value) { score = 0; break; } score += value; }
      if (!score) continue;
      const label = entry.label.toLowerCase();
      if (label === full) score += 6; else if (label.startsWith(full) || (entry.keys && (` ${entry.keys}`).includes(` ${full}`))) score += 3;
      if (entry.placed) score += 0.5;
      if (entry.toggle) score -= 4;
      if (entry.run) score -= 1.5;
      if (entry.id && recent.includes(entry.id)) score += 0.5;
      const enabled = isEnabled(entry);
      if (enabled) score += 2;
      scored.push({ entry, score, enabled });
    }
    scored.sort((a, b) => b.score - a.score || a.entry.label.length - b.entry.label.length);
    return scored.slice(0, matcher.maxResults).map(({ entry, enabled }) => ({ ...(entry.id ? { id: entry.id } : {}), label: entry.label, where: entry.path ? entry.path.join(' ▸ ') : '', enabled, entry }));
  };
  const recentResults = (layout = runtime.document.documentElement.dataset.layout || 'bars'): SearchResult[] => {
    const ids: string[] = L.store ? L.store.get('searchRecent', []) : [];
    const all = entries(layout);
    return ids.map((id) => all.find((entry) => entry.id === id)).filter((entry): entry is SearchEntry => !!entry).map((entry) => ({ ...(entry.id ? { id: entry.id } : {}), label: entry.label, where: entry.path ? entry.path.join(' ▸ ') : '', enabled: isEnabled(entry), entry }));
  };
  ui.searchCommands = searchCommands;
  ui.searchResults = (query: string, layout?: string) => query ? searchCommands(query, layout) : recentResults(layout);
  ui.searchRemember = (id: string) => {
    if (!id || !L.store) return;
    L.store.set('searchRecent', [id, ...L.store.get('searchRecent', []).filter((item: string) => item !== id)].slice(0, matcher.maxRecent));
  };
  ui.searchBox = () => {
    const placeholder = runtime.document.createElement('div');
    placeholder.className = 'cmd-search cmd-search-react-owned';
    placeholder.setAttribute('aria-hidden', 'true');
    placeholder.hidden = true;
    return placeholder;
  };
  ui.searchInput = () => {
    const editor = (runtime as SearchRuntime).VO?.reactDirectHost?.root?.closest('.vibe-editor');
    return editor?.querySelector<HTMLInputElement>('.vibe-command-search input') ?? null;
  };
}
