/** Installs the shared vocabulary and ranking used by classic and React command search. */
export function installCommandMatcher(runtime: Window): void {
  type SearchEntry = { label: string; alias: string; tip: string; path: string[] | null; own?: string[]; also?: string[] };
  type Matcher = { maxResults: number; maxRecent: number; words(value: string): string[]; scoreWord(query: string, entry: SearchEntry): number; [key: symbol]: unknown };
  type Runtime = Window & { VO?: { commandMatcher?: Matcher } };
  const view = runtime as Runtime;
  const classicMarker = Symbol.for('vibeoffice.classicCommandMatcher');
  if (view.VO?.commandMatcher && !view.VO.commandMatcher[classicMarker]) return;

  const containsCjk = /[\u3040-\u30ff\u3400-\u9fff\uac00-\ud7af]/;
  const SYNONYMS: Record<string, string> = {
    smartart: 'intelliart', wordart: 'lettersart', diagram: 'intelliart|diagram', diagrams: 'intelliart|diagram', orgchart: 'intelliart|organization chart', fontwork: 'lettersart',
    image: 'picture|from file', images: 'picture|from file', pdf: 'pdf|save as', png: 'save as', html: 'web page|save as', photo: 'picture|photo album', photos: 'picture|photo album', pic: 'picture', img: 'picture',
    graph: 'chart', graphs: 'chart', plot: 'chart', colour: 'color', colours: 'color', grey: 'gray', centre: 'center', centred: 'center',
    remove: 'delete|clear', erase: 'delete|clear|eraser', strikethrough: 'strike', crossed: 'strike',
    margin: 'page setup', margins: 'page setup', orientation: 'page setup', landscape: 'page setup', portrait: 'page setup', paper: 'page setup',
    note: 'comment', notes: 'comment|notes', annotate: 'comment', toc: 'table of contents|index and tables', contents: 'index and tables|table of contents',
    sum: 'autosum', total: 'autosum|subtotal', totals: 'subtotal', redline: 'track changes', revisions: 'track changes|changes',
    password: 'protect|permission|encrypt', encrypt: 'protect|permission', lock: 'protect|freeze',
    emoji: 'symbol', special: 'symbol', character: 'symbol', characters: 'symbol|word count', link: 'hyperlink', url: 'hyperlink', links: 'hyperlink|links',
    search: 'find', locate: 'find', substitute: 'replace', swap: 'replace',
    uppercase: 'change case', lowercase: 'change case', capital: 'change case', capitals: 'change case', caps: 'change case|all caps',
    bigger: 'grow font', larger: 'grow font', enlarge: 'grow font', smaller: 'shrink font', plain: 'clear formatting|clear formats',
    brush: 'format painter', background: 'background|page color|fill', height: 'spacing|row height', list: 'bullets|numbering',
    bullet: 'bullets', numbered: 'numbering', shading: 'fill|shading', marker: 'highlight', resize: 'size|autofit', fit: 'autofit|fit',
    present: 'slide show|view show|from beginning', presentation: 'slide show|view show', play: 'slide show|view show|from beginning', slideshow: 'slide show|view show', speaker: 'notes',
    money: 'currency', dollar: 'currency', euro: 'currency', percentage: 'percent', decimals: 'decimal',
    formula: 'function|formula', vlookup: 'function', lookup: 'function', average: 'function|autosum', count: 'function|word count',
    export: 'save as|export|pdf', download: 'save as', statistics: 'word count', stats: 'word count',
    dictionary: 'spelling', spellcheck: 'spelling', spell: 'spelling', typo: 'spelling', proofread: 'spelling', synonym: 'thesaurus', synonyms: 'thesaurus',
    translate: 'language', settings: 'options', preferences: 'options', appearance: 'options', theme: 'options|theme', themes: 'options|theme',
    shortcut: 'keyboard shortcuts', shortcuts: 'keyboard shortcuts', hotkey: 'keyboard shortcuts', hotkeys: 'keyboard shortcuts',
    revert: 'undo', magnify: 'zoom', fullscreen: 'full screen', citation: 'footnote', pin: 'freeze', combine: 'merge', join: 'merge',
    alphabetize: 'sort', order: 'sort|bring to front|send to back', clone: 'duplicate', textbox: 'text box', shape: 'autoshapes|shapes',
    arrow: 'autoshapes|arrow', rectangle: 'autoshapes', circle: 'autoshapes', oval: 'autoshapes', email: 'send|mail', mail: 'mail merge|send',
    vba: 'macros', script: 'macros', quit: 'exit', blank: 'new', recent: 'open', author: 'properties', metadata: 'properties',
    info: 'properties', details: 'properties', tab: 'tabs|sheet', worksheet: 'sheet', column: 'column|columns', row: 'row|rows',
  };
  const words = (value: string) => {
    const text = String(value || ''), whole = text.toLowerCase().split(/[^\p{L}\p{N}%$&+]+/u).filter(Boolean);
    const parts = text.replace(/([a-z])([A-Z])/g, '$1 $2').toLowerCase().split(/[^\p{L}\p{N}%$&+]+/u).filter(Boolean);
    return whole.concat(parts.filter((part) => !whole.includes(part)));
  };
  const withinOneEditOrTransposition = (left: string, right: string) => {
    if (Math.abs(left.length - right.length) > 1) return false;
    let index = 0;
    while (index < left.length && left[index] === right[index]) index++;
    if (index === left.length && index === right.length) return true;
    if (left.length === right.length) return left.slice(index + 1) === right.slice(index + 1) || (left[index] === right[index + 1] && left[index + 1] === right[index] && left.slice(index + 2) === right.slice(index + 2));
    return left.length > right.length ? left.slice(index + 1) === right.slice(index) : left.slice(index) === right.slice(index + 1);
  };
  const scoreWord = (query: string, entry: SearchEntry) => {
    const own = entry.own || (entry.own = words(entry.label));
    const also = entry.also || (entry.also = words(entry.alias).filter((word) => !own.includes(word)));
    const bestScoreForTokens = (list: string[]) => {
      let best = 0;
      for (const word of list) {
        if (word === query) best = Math.max(best, 5);
        else if (word.startsWith(query)) best = Math.max(best, 4);
        else if (query.length >= 4 && word.length >= 4 && withinOneEditOrTransposition(query, word)) best = Math.max(best, 2.5);
        else if (containsCjk.test(query) && word.includes(query)) best = Math.max(best, 3.5);
      }
      const text = ` ${list.join(' ')} `;
      if (SYNONYMS[query]?.split('|').some((alt) => text.includes(` ${alt} `) || text.includes(` ${alt}s `))) best = Math.max(best, 3.5);
      return best;
    };
    const best = Math.max(bestScoreForTokens(own), bestScoreForTokens(also) - 0.5);
    if (best > 0) return best;
    return words(`${entry.tip} ${(entry.path || []).join(' ')}`).some((word) => word.startsWith(query)) ? 1 : 0;
  };
  const VO = view.VO || (view.VO = {});
  VO.commandMatcher = Object.freeze({ maxResults: 10, maxRecent: 6, words, scoreWord });
}
