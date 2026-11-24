import { OfficeDialog } from '@vibeoffice/core';
import { useEffect, useState } from 'react';
import { usePptxHost, usePptxRuntime } from './editor-host.js';

interface SpeakerNotesCatalog { slide: number; notes: string }
interface SpeakerNotesRuntime {
  app?: { reactSpeakerNotesCatalog?(): SpeakerNotesCatalog | null; reactSpeakerNotesAction?(action: 'save' | 'close', notes?: string): void };
  bus?: { on(event: string, listener: () => void): () => void };
  t?(message: string, ...args: unknown[]): string;
}

/** React-owned Speaker Notes dialog; the presentation engine retains the undoable notes transaction. */
export function PptxSpeakerNotesDialog() {
  const host = usePptxHost();
  const runtime = usePptxRuntime<SpeakerNotesRuntime>();
  const [catalog, setCatalog] = useState<SpeakerNotesCatalog | null>(null);
  const [notes, setNotes] = useState('');
  useEffect(() => {
    const refresh = () => {
      const next = runtime?.app?.reactSpeakerNotesCatalog?.() || null;
      setCatalog(next);
      setNotes(next?.notes || '');
    };
    const off = runtime?.bus?.on('pptx-speaker-notes-dialog', refresh);
    return () => off?.();
  }, []);
  useEffect(() => {
    if (!catalog) return;
    const cancelDraft = (event: KeyboardEvent) => {
      if (event.key !== 'Escape') return;
      event.preventDefault();
      event.stopImmediatePropagation();
      runtime?.app?.reactSpeakerNotesAction?.('close');
      setCatalog(null);
    };
    host.ownerDocument.addEventListener('keydown', cancelDraft, true);
    return () => host.ownerDocument.removeEventListener('keydown', cancelDraft, true);
  }, [catalog]);
  const close = () => { runtime?.app?.reactSpeakerNotesAction?.('close'); setCatalog(null); };
  if (!catalog) return null;
  const t = (message: string, ...args: unknown[]) => runtime?.t?.(message, ...args) || message;
  const save = () => { runtime?.app?.reactSpeakerNotesAction?.('save', notes); setCatalog(null); };
  return <OfficeDialog title={t('Speaker Notes')} className="pptx-speaker-notes-dialog" style={{ width: 'min(92vw, 420px)' }} onRequestClose={close}>
    <div className="vibeoffice-dialog__body">
      <label htmlFor="pptx-speaker-notes-text">{t('Slide {0}:', catalog.slide)}</label>
      <textarea id="pptx-speaker-notes-text" aria-label={t('Speaker Notes')} rows={10} autoFocus style={{ width: '100%' }} value={notes} onChange={(event) => setNotes(event.currentTarget.value)} />
    </div>
    <div className="vibeoffice-dialog__actions"><button type="button" className="primary" onClick={save}>{t('Close')}</button></div>
  </OfficeDialog>;
}
