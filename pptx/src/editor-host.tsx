import { createContext, useContext, type ReactNode } from 'react';

const PptxHostContext = createContext<HTMLElement | null>(null);

/** Binds React-owned presentation UI to the document that owns its mounted engine. */
export function PptxHostProvider({ host, children }: { host: HTMLElement; children: ReactNode }) {
  return <PptxHostContext.Provider value={host}>{children}</PptxHostContext.Provider>;
}

/** Resolve the legacy engine namespace from the editor instance hosting this React layer. */
export function usePptxRuntime<T>() {
  const host = usePptxHost();
  return (host.ownerDocument.defaultView as (Window & { L?: T }) | null)?.L;
}

export function usePptxHost() {
  const host = useContext(PptxHostContext);
  if (!host) throw new Error('PPTX React layers must be mounted inside PptxHostProvider');
  return host;
}
