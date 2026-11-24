import type { LegacyDialogMountRequest, LegacyDialogOptions } from './legacy-dialog.js';

/** Replaces the legacy synchronous dialog facade with a request to the React dialog layer. */
export function installReactLegacyDialogService(runtime: Window): void {
  type Runtime = Window & {
    L?: { ui?: Record<string, any> };
    VO?: { reactDirectHost?: { root?: HTMLElement; dialogStack?: object[] } };
  };
  const view = runtime as Runtime;
  const ui = view.L?.ui;
  const direct = view.VO?.reactDirectHost;
  const root = direct?.root;
  if (!ui || !direct || !root) return;
  const CustomEventConstructor = (runtime as Window & typeof globalThis).CustomEvent;
  const stack = direct.dialogStack ?? (direct.dialogStack = []);
  ui.dialogOpen = () => stack.length > 0;
  ui.dialog = (options: LegacyDialogOptions) => {
    ui.closeMenus?.();
    ui.hideTip?.();
    const doc = root.ownerDocument;
    const previousFocus = doc.activeElement;
    const selection = doc.defaultView?.getSelection();
    const previousRange = previousFocus instanceof doc.defaultView!.HTMLElement
      && previousFocus.isContentEditable && selection?.rangeCount && previousFocus.contains(selection.anchorNode)
      ? selection.getRangeAt(0).cloneRange() : null;
    const buttons = (options.buttons?.length ? options.buttons : [{ label: 'OK', primary: true }]).map((button) => ({
      ...button,
      labelHtml: ui.ampHTML?.(button.label) ?? button.label,
    }));
    const token = {};
    let resolve!: (index: number | null) => void;
    let closed = false;
    const done = new Promise<number | null>((complete) => { resolve = complete; });
    const restoreFocus = () => {
      const live = previousFocus?.isConnected && previousFocus !== doc.body;
      if (live && !(previousFocus instanceof doc.defaultView!.HTMLElement && previousFocus.isContentEditable && ui.hooks?.ownsText?.(previousFocus))) {
        try { (previousFocus as HTMLElement).focus({ preventScroll: true }); } catch { /* ignore */ }
        if (previousRange?.startContainer.isConnected && previousRange.endContainer.isConnected) {
          const currentSelection = doc.defaultView?.getSelection();
          currentSelection?.removeAllRanges();
          currentSelection?.addRange(previousRange);
        }
      } else ui.hooks?.refocus?.();
    };
    const request: LegacyDialogMountRequest = {
      host: root.closest('.vibe-editor__engine') || root,
      token,
      options: { ...options, buttons },
      done,
      resolve,
      stack,
      restoreFocus: () => restoreFocus(),
      close: (index) => {
        if (closed) return;
        closed = true;
        const position = stack.indexOf(token);
        if (position >= 0) stack.splice(position, 1);
        request.unmount?.();
        options.onClose?.(index);
        resolve(index);
        restoreFocus();
      },
    };
    if (!options.modeless) stack.push(token);
    root.dispatchEvent(new CustomEventConstructor('vibeoffice-react-dialog', { detail: request, bubbles: true }));
    if (request.api) return request.api;
    const position = stack.indexOf(token);
    if (position >= 0) stack.splice(position, 1);
    return undefined;
  };
  ui.dialogKey = () => stack.length > 0;
}
