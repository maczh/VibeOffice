import { createContext, useContext, type ReactNode } from 'react';

const ExcelxHostContext = createContext<HTMLElement | null>(null);

/** Binds React-owned spreadsheet UI to the document hosting its engine instance. */
export function ExcelxHostProvider({ host, children }: { host: HTMLElement; children: ReactNode }) {
  return <ExcelxHostContext.Provider value={host}>{children}</ExcelxHostContext.Provider>;
}

export function useExcelxHost() {
  const host = useContext(ExcelxHostContext);
  if (!host) throw new Error('Excelx React layers must be mounted inside ExcelxHostProvider');
  return host;
}

/** Resolve the spreadsheet engine namespace from this React layer's editor host. */
export function useExcelxRuntime<T>() {
  return (useExcelxHost().ownerDocument.defaultView as (Window & { L?: T }) | null)?.L;
}
