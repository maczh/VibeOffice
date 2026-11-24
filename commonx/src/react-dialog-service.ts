/** Bridges the legacy dialog request API to dialogs rendered by the React host. */
export function installReactDialogService(runtime: Window): void {
  type Runtime = Window & {
    L?: { APP?: string; VERSION?: string; t?: (text: string, ...args: unknown[]) => string; ui?: Record<string, any>; icons?: { app?: (size: number) => string }; opc?: { noticeLosses?: (doc: unknown, entries: unknown[]) => unknown[] } };
    VO?: { reactDirectHost?: { root?: HTMLElement }; LICENSES?: string; i18n?: { plain?: (text: string) => string } };
  };
  const view = runtime as Runtime;
  const L = view.L;
  const ui = L?.ui;
  if (!L || !ui) return;
  const host = () => view.VO?.reactDirectHost?.root;
  const CustomEventConstructor = (runtime as Window & typeof globalThis).CustomEvent;
  const request = <R>(type: string, detail: Record<string, unknown>, fallback: R): Promise<R> => new Promise((resolve) => {
    const root = host();
    if (!root?.isConnected) { resolve(fallback); return; }
    root.dispatchEvent(new CustomEventConstructor(type, { detail: { ...detail, respond: resolve }, bubbles: true }));
  });
  const icons: Record<string, string> = {
    info: '<svg width="32" height="32" viewBox="0 0 32 32"><circle cx="16" cy="16" r="14" fill="#2c64c8" stroke="#163b84"/><text x="12.5" y="24" font-family="Georgia,serif" font-size="19" font-weight="bold" fill="#fff">i</text></svg>',
    warn: '<svg width="32" height="32" viewBox="0 0 32 32"><path d="M16 2l14.5 26h-29z" fill="#f6cf3b" stroke="#a07c08"/><text x="13" y="25" font-family="Arial" font-size="17" font-weight="bold" fill="#222">!</text></svg>',
    question: '<svg width="32" height="32" viewBox="0 0 32 32"><circle cx="16" cy="16" r="14" fill="#2c64c8" stroke="#163b84"/><text x="10.5" y="23.5" font-family="Arial" font-size="18" font-weight="bold" fill="#fff">?</text></svg>',
    error: '<svg width="32" height="32" viewBox="0 0 32 32"><circle cx="16" cy="16" r="14" fill="#d33a2c" stroke="#8a1a10"/><path d="M10 10l12 12M22 10L10 22" stroke="#fff" stroke-width="3.4"/></svg>',
  };
  ui.linkDialog = (url: string) => ({ done: request('vibeoffice-react-link', { url, mail: /^mailto:/i.test(url) }, undefined) });
  ui.msg = (text: string, options?: Record<string, any>) => {
    const o = options || {};
    const icon = o.icon || 'info';
    const buttons: string[] = o.buttons || ['OK'];
    const plain = buttons.map((label) => (view.VO?.i18n?.plain ? view.VO.i18n.plain(String(label || '')) : String(label || '')).replace(/&&/g, '\u0000').replace(/&(\S)/g, '$1').replace(/\u0000/g, '&'));
    const escapeIndex = plain.some((label) => /cancel/i.test(label)) ? plain.findIndex((label) => /cancel/i.test(label)) : plain.findIndex((label) => /^no$/i.test(label));
    runtime.dispatchEvent(new CustomEventConstructor('vo-message', { detail: { icon } }));
    return request<number | null>('vibeoffice-react-message', {
      title: o.title || L.APP, text: String(L.t ? L.t(text) : text), icon,
      iconSvg: icons[icon] || icons.info, buttons, defaultButton: o.def || 0, escapeIndex, width: o.width || 380,
    }, null);
  };
  ui.about = (options: { name: string; text: string[] }) => ({ done: request('vibeoffice-react-about', {
    name: options.name, text: options.text, version: L.VERSION || 'dev', icon: L.icons?.app ? L.icons.app(48) : '',
    licensesUrl: view.VO?.LICENSES || '../oss-licenses.html',
  }, 0) });
  ui.prompt = (label: string, value?: string, title?: string) => request<string | null>('vibeoffice-react-prompt', {
    title: title || L.APP, label, value: value || '',
  }, null);
  ui.password = (fileName: string, wrong?: boolean) => request<string | null>('vibeoffice-react-password', {
    text: wrong ? '' : (L.t ? L.t("'{0}' is protected.", fileName) : `'${fileName}' is protected.`), wrong: !!wrong,
  }, null);
  ui.compatibility = async (doc: unknown, entries: unknown[]) => {
    const pending = L.opc?.noticeLosses?.(doc, entries) ?? [];
    if (!pending.length) return [];
    const groups = new Map<string, string[]>();
    for (const entry of pending as Array<{ what: string; place?: string }>) {
      if (!groups.has(entry.what)) groups.set(entry.what, []);
      const places = groups.get(entry.what)!;
      if (entry.place && !places.includes(entry.place)) places.push(entry.place);
    }
    const proceed = await request<boolean>('vibeoffice-react-compatibility', {
      groups: Array.from(groups, ([what, places]) => ({ what, places })),
    }, false);
    return proceed ? pending : null;
  };
}
