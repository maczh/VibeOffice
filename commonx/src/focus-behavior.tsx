import { useEffect } from 'react';

interface Brush {
  stop(): void;
}

interface EditorUi {
  brushOn?(): Brush | null;
  dialogOpen?(): boolean;
  menuOpen?(): boolean;
}

/** Preserve editor keyboard focus when its non-editable chrome receives a mouse press. */
export function EditorFocusBehavior({ host }: { host: HTMLElement }) {
  useEffect(() => {
    const keepEditorKeyboard = (event: MouseEvent) => {
      const view = host.ownerDocument.defaultView;
      const target = event.target;
      if (!view || !(target instanceof view.Element) || target.closest('.dlg')) return;
      const focusable = target.closest('input, select, textarea, [contenteditable=""], [contenteditable="true"], a[href], label, [tabindex], button');
      if (focusable && !focusable.matches('button, .menu, .mb-item')) return;
      event.preventDefault();
    };
    const escapeBrush = (event: KeyboardEvent) => {
      if (event.key !== 'Escape') return;
      const view = host.ownerDocument.defaultView as (Window & { L?: { ui?: EditorUi } }) | null;
      const ui = view?.L?.ui;
      const brush = ui?.brushOn?.();
      if (!brush || ui?.dialogOpen?.() || ui?.menuOpen?.()) return;
      event.preventDefault();
      event.stopPropagation();
      brush.stop();
    };
    host.addEventListener('mousedown', keepEditorKeyboard);
    host.addEventListener('keydown', escapeBrush, true);
    return () => {
      host.removeEventListener('mousedown', keepEditorKeyboard);
      host.removeEventListener('keydown', escapeBrush, true);
    };
  }, [host]);

  return null;
}
