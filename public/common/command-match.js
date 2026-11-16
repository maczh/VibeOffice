/* Shared command-search vocabulary and ranking, used by the classic and React search surfaces. */
(function (root) {
  'use strict';
  const CJK = /[\u3040-\u30ff\u3400-\u9fff\uac00-\ud7af]/;
  const SYNONYMS = {
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

  const words = (value) => {
    const text = String(value || ''), whole = text.toLowerCase().split(/[^\p{L}\p{N}%$&+]+/u).filter(Boolean);
    const parts = text.replace(/([a-z])([A-Z])/g, '$1 $2').toLowerCase().split(/[^\p{L}\p{N}%$&+]+/u).filter(Boolean);
    return whole.concat(parts.filter((part) => !whole.includes(part)));
  };
  function near(a, b) {
    if (Math.abs(a.length - b.length) > 1) return false;
    let i = 0;
    while (i < a.length && a[i] === b[i]) i++;
    if (i === a.length && i === b.length) return true;
    if (a.length === b.length) return a.slice(i + 1) === b.slice(i + 1) || (a[i] === b[i + 1] && a[i + 1] === b[i] && a.slice(i + 2) === b.slice(i + 2));
    return a.length > b.length ? a.slice(i + 1) === b.slice(i) : a.slice(i) === b.slice(i + 1);
  }
  function scoreWord(query, entry) {
    const own = entry.own || (entry.own = words(entry.label));
    const also = entry.also || (entry.also = words(entry.alias).filter((word) => !own.includes(word)));
    const hit = (list) => {
      let best = 0;
      for (const word of list) {
        if (word === query) best = Math.max(best, 5);
        else if (word.startsWith(query)) best = Math.max(best, 4);
        else if (query.length >= 4 && word.length >= 4 && near(query, word)) best = Math.max(best, 2.5);
        else if (CJK.test(query) && word.includes(query)) best = Math.max(best, 3.5);
      }
      const text = ` ${list.join(' ')} `;
      if (SYNONYMS[query]?.split('|').some((alt) => text.includes(` ${alt} `) || text.includes(` ${alt}s `))) best = Math.max(best, 3.5);
      return best;
    };
    const best = Math.max(hit(own), hit(also) - 0.5);
    if (best > 0) return best;
    return words(`${entry.tip} ${(entry.path || []).join(' ')}`).some((word) => word.startsWith(query)) ? 1 : 0;
  }
  const suite = root.VO || (root.VO = {});
  suite.commandMatcher = Object.freeze({ maxResults: 10, maxRecent: 6, words, scoreWord });
})(window);
