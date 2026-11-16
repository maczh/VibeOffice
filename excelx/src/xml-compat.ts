interface ExcelxXmlRuntime extends Window {
  L?: {
    xml?: unknown;
    xmlTree?: unknown;
  };
}

/** Install the one-name XML alias expected by the existing Excelx engine modules. */
export function installExcelxXmlCompatibility(runtime: Window): void {
  const L = (runtime as ExcelxXmlRuntime).L;
  if (!L || L.xmlTree == null) throw new Error('Excelx XML tree service is not ready');
  L.xml = L.xmlTree;
}
