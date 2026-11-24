/** Installs command registration/execution for React-hosted editors. */
export function installReactCommandService(runtime: Window): () => void {
  type Command = { id: string; run?: (arg?: unknown, event?: Event) => unknown; enabled?: () => boolean; checked?: () => boolean; [key: string]: unknown };
  type Runtime = Window & {
    console: Console;
    L?: {
      ui?: Record<string, any>;
      t?: (text: string, ...args: unknown[]) => string;
      uiHooks?: unknown;
    };
    CustomEvent: typeof CustomEvent;
  };
  const view = runtime as Runtime;
  const ui = view.L?.ui;
  if (!ui) return () => {};
  const originals = new Map<string, unknown>();
  const installed = new Map<string, (...args: any[]) => any>();
  const replace = (name: string, value: (...args: any[]) => any) => {
    originals.set(name, ui[name]);
    installed.set(name, value);
    ui[name] = value;
  };
  if (!ui.cmds || typeof ui.cmds !== 'object') ui.cmds = {};
  const commands = ui.cmds as Record<string, Command>;
  const tell = (type: string, detail: Record<string, unknown>) => view.dispatchEvent(new view.CustomEvent(type, { detail }));
  replace('cmd', (id: string, definition: Record<string, unknown>) => {
    const command = { id } as Command;
    Object.defineProperties(command, Object.getOwnPropertyDescriptors(definition));
    commands[id] = command;
    return command;
  });
  replace('exec', (id: string, argument?: unknown, event?: Event) => {
    const command = commands[id];
    if (!command) { view.console.warn('unknown command', id); return undefined; }
    if (command.enabled && !command.enabled()) return undefined;
    ui.hooks?.beforeExec?.();
    let result: unknown;
    try { result = command.run!(argument, event); }
    catch (error) {
      view.console.error(error);
      const message = error && typeof error === 'object' && 'message' in error ? String(error.message) : String(error);
      ui.msg?.(view.L?.t?.('That action could not be completed: {0}', message) ?? `That action could not be completed: ${message}`, { icon: 'warn' });
    }
    ui.refresh?.();
    tell('vo-command', { id, result });
    return result;
  });
  replace('enabled', (id: string) => {
    const command = commands[id];
    return !!command && (!command.enabled || !!command.enabled());
  });
  replace('checked', (id: string) => {
    const command = commands[id];
    return !!command && !!command.checked && !!command.checked();
  });
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
