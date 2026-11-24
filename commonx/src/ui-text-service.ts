/** Installs the shared label/access-key helpers used by React-owned editor chrome. */
export function installReactUiTextService(runtime: Window): () => void {
  type Runtime = Window & {
    L?: {
      ui?: Record<string, any>;
      t?: (text: string) => string;
      esc?: (text: string) => string;
      isMac?: boolean;
    };
    VO?: { reactDirectHost?: { root?: HTMLElement | null }; i18n?: { plain?: (text: string) => string } };
  };
  const view = runtime as Runtime;
  const ui = view.L?.ui;
  const root = view.VO?.reactDirectHost?.root;
  if (!ui || !root || !view.L) return () => {};

  const originals = new Map<string, unknown>();
  const installed = new Map<string, (...args: any[]) => any>();
  const replace = (name: string, value: (...args: any[]) => any) => {
    originals.set(name, ui[name]);
    installed.set(name, value);
    ui[name] = value;
  };
  const translate = (text: string) => view.L?.t?.(text) ?? text;
  const plainAmp = (text: string) => (view.VO?.i18n?.plain?.(String(text || '')) ?? String(text || ''))
    .replace(/&&/g, '\u0000').replace(/&([^\s])/g, '$1').replace(/\u0000/g, '&');
  const stripAmp = (text: string) => plainAmp(translate(text));
  const ampHTML = (text: string) => (view.L?.esc?.(translate(text || '')) ?? String(text || ''))
    .replace(/&amp;&amp;/g, '\u0000').replace(/&amp;(\S)/g, '<u>$1</u>').replace(/\u0000/g, '&amp;');
  const accessKey = (text: string) => {
    const match = /&([^&\s])/.exec(translate(text) || '');
    return match?.[1]?.toLowerCase() ?? null;
  };
  const keyName = (part: string) => ({
    Ctrl: translate('Ctrl|key'), Shift: translate('Shift|key'), Alt: translate('Alt|key'), Del: translate('Del|key'),
    Ins: translate('Ins|key'), Esc: translate('Esc|key'), Home: translate('Home|key'), End: translate('End|key'),
    Enter: translate('Enter|key'), Tab: translate('Tab|key'), Space: translate('Space|key'), Backspace: translate('Backspace|key'),
    'Page Up': translate('Page Up|key'), 'Page Down': translate('Page Down|key'), PgUp: translate('PgUp|key'), PgDn: translate('PgDn|key'),
  } as Record<string, string>)[part] || part;
  const keyText = (value: string) => {
    if (!value) return '';
    const parts = String(value).split('+');
    if (!view.L?.isMac) return parts.map(keyName).join('+');
    const modifiers: Array<[string, string]> = [['Alt', '⌥'], ['Shift', '⇧'], ['Ctrl', '⌘']];
    return modifiers.filter(([modifier]) => parts.includes(modifier)).map(([, symbol]) => symbol).join('')
      + parts.filter((part) => !modifiers.some(([modifier]) => modifier === part)).map(keyName).join('+');
  };
  replace('plainAmp', plainAmp);
  replace('stripAmp', stripAmp);
  replace('ampHTML', ampHTML);
  replace('accessKey', accessKey);
  replace('keyText', keyText);
  return () => {
    for (const [name, value] of installed) {
      if (ui[name] !== value) continue;
      const original = originals.get(name);
      if (original === undefined) delete ui[name];
      else ui[name] = original;
    }
    installed.clear();
    originals.clear();
  };
}
