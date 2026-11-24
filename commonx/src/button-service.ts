/** Installs the DOM-compatible button facade backed by the React button portal. */
export function installReactButtonService(runtime: Window): void {
  type Runtime = Window & {
    L?: { ui?: Record<string, any>; esc?: (text: string) => string; h?: (tag: string, attributes?: Record<string, unknown>) => HTMLButtonElement };
    VO?: { reactDirectHost?: { root?: HTMLElement } };
  };
  const view = runtime as Runtime;
  const ui = view.L?.ui;
  const root = view.VO?.reactDirectHost?.root;
  if (!ui || !root) return;
  const CustomEventConstructor = (runtime as Window & typeof globalThis).CustomEvent;
  let nextId = 0;
  ui.button = (label: string, onClick?: (event: MouseEvent) => void, options?: Record<string, any>) => {
    const id = options?.id || `btn${++nextId}`;
    const attributes = Object.assign({ class: 'btn', type: 'button' }, options || {}, { id });
    delete attributes.html;
    const button = view.L?.h?.('button', attributes) ?? runtime.document.createElement('button');
    if (!view.L?.h) Object.assign(button, attributes);
    const buttonOptions = Object.assign({ html: ui.ampHTML(label) }, options || {}, { id, label });
    const request = { target: button, options: buttonOptions };
    button.dataset.vibeofficeReactButton = '';
    (button as HTMLButtonElement & { __vibeofficeReactButtonRequest?: typeof request }).__vibeofficeReactButtonRequest = request;
    const notify = () => root.dispatchEvent(new CustomEventConstructor('vibeoffice-react-button', { detail: request, bubbles: true }));
    Object.defineProperties(button, {
      innerHTML: { configurable: true, get: () => String(buttonOptions.html || ''), set: (value: string) => { buttonOptions.html = String(value); delete buttonOptions.textContent; notify(); } },
      textContent: { configurable: true, get: () => String(buttonOptions.textContent ?? button.querySelector('.vibeoffice-react-button-label')?.textContent ?? buttonOptions.label ?? ''), set: (value: string) => { buttonOptions.html = view.L?.esc?.(String(value)) ?? String(value); buttonOptions.label = String(value); buttonOptions.textContent = String(value); notify(); } },
    });
    notify();
    if (onClick) button.addEventListener('click', onClick);
    return button;
  };
}
