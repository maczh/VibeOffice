import { OfficeDialog } from '@vibeoffice/core';
import { useEffect, useId, useState, type ChangeEvent } from 'react';

interface OptionsCatalog { options: Record<string, boolean>; optionKeys: string[]; undo: number; user: string; password: string; looks: Array<[string, string]>; languages: Array<[string, string]>; look: string; lang: string }
interface OptionsRuntime {
  dlg?: {
    reactOptionsCatalog?(): OptionsCatalog | null;
    reactOptionsAction?(action: string, values?: { options: Record<string, boolean>; undo: number; user: string; password: string; look: string; lang: string }): boolean;
    reactOptionsConfirmCatalog?(): { target: string } | null;
    reactOptionsConfirmAction?(action: string, value?: string): boolean;
  };
  bus?: { on(event: string, listener: () => void): () => void };
  t?(message: string): string;
}
function runtime(): OptionsRuntime | undefined { return (window as Window & { L?: OptionsRuntime }).L; }
const viewOptions: Array<[string, string]> = [
  ['startupPane', 'Startup &Task Pane'], ['layoutPaneOnNew', '&Slide Layout task pane when inserting new slides'], ['paneFollows', 'Task pane f&ollows the selected object'], ['statusBar', 'Stat&us bar'], ['ruler', '&Vertical ruler'],
  ['showPopup', '&Popup menu on right mouse click'], ['showPopbar', 'Show popup menu &button'], ['endBlack', '&End with black slide'],
];
const editOptions: Array<[string, string]> = [['smartQuotes', 'Replace straight &quotes with smart quotes'], ['autocorrect', '&AutoCorrect as you type'], ['capSentence', 'Capitalize first letter of &sentences']];

/** React owns PPTX application preferences and password confirmation; settings and persistence stay in the host engine. */
export function PptxOptionsDialog() {
  const [catalog, setCatalog] = useState<OptionsCatalog | null>(null);
  const [confirmTarget, setConfirmTarget] = useState<string | null>(null);
  const [tab, setTab] = useState(0);
  const [options, setOptions] = useState<Record<string, boolean>>({});
  const [undo, setUndo] = useState(20);
  const [user, setUser] = useState('');
  const [password, setPassword] = useState('');
  const [look, setLook] = useState('');
  const [lang, setLang] = useState('');
  const [confirmation, setConfirmation] = useState('');
  const id = useId().replaceAll(':', '');
  const panelId = `${id}-options-panel`;
  const tabId = (index: number) => `${id}-options-tab-${index}`;
  useEffect(() => {
    const openOptions = () => {
      const next = runtime()?.dlg?.reactOptionsCatalog?.() || null;
      setCatalog(next ? { ...next, options: { ...next.options }, looks: [...next.looks], languages: [...next.languages] } : null);
      if (next) { setTab(0); setOptions({ ...next.options }); setUndo(next.undo); setUser(next.user); setPassword(next.password); setLook(next.look); setLang(next.lang); }
    };
    const openConfirm = () => {
      const next = runtime()?.dlg?.reactOptionsConfirmCatalog?.() || null;
      setConfirmTarget(next?.target ?? null);
      if (next) setConfirmation('');
    };
    const offOptions = runtime()?.bus?.on('pptx-options-dialog', openOptions);
    const offConfirm = runtime()?.bus?.on('pptx-options-confirm-dialog', openConfirm);
    return () => { offOptions?.(); offConfirm?.(); };
  }, []);
  const t = (message: string) => runtime()?.t?.(message) || message;
  const closeOptions = () => runtime()?.dlg?.reactOptionsAction?.('close');
  const apply = () => catalog && runtime()?.dlg?.reactOptionsAction?.('apply', { options, undo, user, password, look, lang });
  const closeConfirm = () => runtime()?.dlg?.reactOptionsConfirmAction?.('close');
  const confirmPassword = () => runtime()?.dlg?.reactOptionsConfirmAction?.('confirm', confirmation);
  const toggle = (key: string) => (event: ChangeEvent<HTMLInputElement>) => {
    const checked = event.currentTarget.checked;
    setOptions((current) => ({ ...current, [key]: checked }));
  };
  const tabs = ['View', 'General', 'Edit', 'Spelling and Style', 'Security'];
  const check = ([key, label]: [string, string]) => <label key={key}><input type="checkbox" aria-label={t(label).replace('&', '')} checked={!!options[key]} onChange={toggle(key)} />{t(label).replace('&', '')}</label>;
  return <>
    {catalog && <OfficeDialog title={t('Options')} className="pptx-options-dialog" style={{ width: 'min(96vw, 560px)' }} onRequestClose={closeOptions}>
      <div role="tablist" aria-label={t('Options')} className="row">{tabs.map((label, index) => <button key={label} id={tabId(index)} type="button" role="tab" aria-controls={panelId} aria-selected={tab === index} tabIndex={tab === index ? 0 : -1} onClick={() => setTab(index)} onKeyDown={(event) => {
        const next = event.key === 'ArrowRight' ? (index + 1) % tabs.length
          : event.key === 'ArrowLeft' ? (index - 1 + tabs.length) % tabs.length
            : event.key === 'Home' ? 0
              : event.key === 'End' ? tabs.length - 1 : -1;
        if (next < 0) return;
        event.preventDefault();
        setTab(next);
        event.currentTarget.parentElement?.querySelectorAll<HTMLButtonElement>('[role="tab"]')[next]?.focus();
      }}>{t(label)}</button>)}</div>
      <section id={panelId} aria-labelledby={tabId(tab)} role="tabpanel" className="col" style={{ minHeight: 220, paddingTop: 12 }}>
        {tab === 0 && <fieldset><legend>{t('Show')}</legend><div className="col">{viewOptions.slice(0, 5).map(check)}</div><fieldset><legend>{t('Slide show')}</legend>{viewOptions.slice(5).map(check)}</fieldset></fieldset>}
        {tab === 1 && <div className="col">{(catalog.looks.length > 0 || catalog.languages.length > 0) && <fieldset><legend>{t('General')}</legend>{catalog.looks.length > 0 && <label>{t('Theme:')}<select aria-label={t('Theme')} value={look} onChange={(event) => setLook(event.currentTarget.value)}>{catalog.looks.map(([value, label]) => <option key={value} value={value}>{t(label)}</option>)}</select></label>}{catalog.languages.length > 0 && <label>{t('Language:')}<select aria-label={t('Language')} value={lang} onChange={(event) => setLang(event.currentTarget.value)}>{catalog.languages.map(([value, label]) => <option key={value} value={value}>{label}</option>)}</select></label>}</fieldset>}<fieldset><legend>{t('User information')}</legend><label>{t('&Name:').replace('&', '')}<input aria-label={t('Name')} type="text" value={user} onChange={(event) => setUser(event.currentTarget.value)} /></label></fieldset></div>}
        {tab === 2 && <div className="col"><fieldset><legend>{t('Text')}</legend>{editOptions.map(check)}</fieldset><fieldset><legend>{t('Undo')}</legend><label>{t('&Maximum number of undos:').replace('&', '')}<input aria-label={t('Maximum number of undos')} type="number" min={3} max={150} step={1} value={undo} onChange={(event) => setUndo(Number(event.currentTarget.value))} /></label></fieldset></div>}
        {tab === 3 && <fieldset><legend>{t('Spelling')}</legend>{check(['spell', 'Check spelling as you &type'])}<div className="tp-note">{t('Spelling suggestions come from your browser. Hold Shift while right-clicking a word to see them.')}</div></fieldset>}
        {tab === 4 && <fieldset><legend>{t('File encryption settings for this document')}</legend><label>{t('Password to &open:').replace('&', '')}<input aria-label={t('Password to open')} type="password" autoComplete="new-password" value={password} onChange={(event) => setPassword(event.currentTarget.value)} /></label><div className="tp-note">{t('The presentation is encrypted (AES-256) the next time you save it as a PowerPoint file, the way PowerPoint 2007 and later do. Leave the box empty to save it without a password. If you lose or forget the password, it cannot be recovered.')}</div></fieldset>}
      </section>
      <div className="vibeoffice-dialog__actions"><button className="primary" type="button" onClick={apply}>{t('OK')}</button><button type="button" onClick={closeOptions}>{t('Cancel')}</button></div>
    </OfficeDialog>}
    {confirmTarget !== null && <OfficeDialog title={t('Confirm Password')} className="pptx-options-confirm-dialog" style={{ width: 'min(96vw, 380px)' }} onRequestClose={closeConfirm}>
      <div className="col"><label>{t('Reenter password to open:')}<input autoFocus aria-label={t('Reenter password to open')} type="password" autoComplete="new-password" value={confirmation} onChange={(event) => setConfirmation(event.currentTarget.value)} /></label><div className="tp-note">{t('Caution: If you lose or forget the password, it cannot be recovered.')}</div></div>
      <div className="vibeoffice-dialog__actions"><button className="primary" type="button" onClick={confirmPassword}>{t('OK')}</button><button type="button" onClick={closeConfirm}>{t('Cancel')}</button></div>
    </OfficeDialog>}
  </>;
}
