import { OfficeDialog } from '@vibeoffice/core';
import { useEffect, useState } from 'react';

const choices = [
  ['comments', '&Comments'], ['constants', 'C&onstants'], ['formulas', '&Formulas'], ['blanks', '&Blanks'], ['region', 'Current &region'], ['array', 'Current &array'], ['objects', 'O&bjects'],
  ['rowdiff', 'Ro&w differences'], ['coldiff', 'Col&umn differences'], ['prec', '&Precedents'], ['dep', '&Dependents'], ['last', 'La&st cell'], ['visible', 'Visible cells onl&y'], ['cf', 'Condi&tional formats'], ['dv', 'Data &validation'],
] as const;
type GoToSpecialKind = typeof choices[number][0];
interface GoToSpecialState { kind: GoToSpecialKind; types: { n: boolean; t: boolean; l: boolean; e: boolean } }
interface GoToSpecialCatalog extends GoToSpecialState {}
interface GoToSpecialRuntime {
  dlg?: { reactGoToSpecialCatalog?(): GoToSpecialCatalog | null; reactGoToSpecialApply?(state: GoToSpecialState): boolean; reactGoToSpecialClose?(): void };
  bus?: { on(event: string, listener: () => void): () => void };
  t?(message: string): string;
}

function runtime(): GoToSpecialRuntime | undefined {
  return (window as Window & { L?: GoToSpecialRuntime }).L;
}

/** React option surface for Go To Special; selection and search remain in the workbook engine. */
export function ExcelxGoToSpecialDialog() {
  const [catalog, setCatalog] = useState<GoToSpecialCatalog | null>(null);
  const [kind, setKind] = useState<GoToSpecialKind>('constants');
  const [types, setTypes] = useState<GoToSpecialState['types']>({ n: true, t: true, l: true, e: true });
  useEffect(() => {
    const refresh = () => {
      const next = runtime()?.dlg?.reactGoToSpecialCatalog?.() || null;
      setCatalog(next);
      if (next) { setKind(next.kind); setTypes(next.types); }
    };
    const off = runtime()?.bus?.on('excelx-go-to-special-dialog', refresh);
    return () => {
      off?.();
      if (runtime()?.dlg?.reactGoToSpecialCatalog?.()) runtime()?.dlg?.reactGoToSpecialClose?.();
    };
  }, []);
  if (!catalog) return null;
  const translate = (message: string) => runtime()?.t?.(message) || message;
  const close = () => runtime()?.dlg?.reactGoToSpecialClose?.();
  const apply = () => runtime()?.dlg?.reactGoToSpecialApply?.({ kind, types });
  const typeNames = [['n', 'Numbers'], ['t', 'Text'], ['l', 'Logicals'], ['e', 'Errors']] as const;
  return <OfficeDialog title={translate('Go To Special')} className="excelx-go-to-special-dialog" style={{ width: 'min(92vw, 460px)' }} onRequestClose={close}>
    <div className="excelx-go-to-special-dialog__choices" role="radiogroup" aria-label={translate('Go To Special')}>
      {choices.map(([value, label]) => <div key={value}>
        <label><input type="radio" name="excelx-go-to-special" value={value} checked={kind === value} onChange={() => setKind(value)} />{translate(label).replace('&', '')}</label>
        {value === 'formulas' && kind === 'formulas' && <div className="excelx-go-to-special-dialog__types">{typeNames.map(([type, name]) => <label key={type}><input type="checkbox" checked={types[type]} onChange={(event) => { const checked = event.currentTarget.checked; setTypes((previous) => ({ ...previous, [type]: checked })); }} />{translate(name)}</label>)}</div>}
      </div>)}
    </div>
    <div className="excelx-go-to-special-dialog__actions">
      <button type="button" className="primary" onClick={apply}>{translate('OK')}</button>
      <button type="button" onClick={close}>{translate('Cancel')}</button>
    </div>
  </OfficeDialog>;
}
