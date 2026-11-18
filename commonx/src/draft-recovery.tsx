import { useEffect, useState } from 'react';
import type { LocalDraft } from './local-files.js';

export interface DraftRecoveryUpdate {
  editor: string;
  drafts: LocalDraft<Blob>[];
  restore(draft: LocalDraft<Blob>): void;
  discard(id: string): Promise<void>;
}

type DraftRecoveryWindow = Window & { __vibeofficeDraftRecoveries?: Record<string, DraftRecoveryUpdate> };
type RecoveryRuntime = { t?(text: string, ...args: unknown[]): string };

/** React-owned view of local crash-recovery drafts, independent of Recent Files. */
export function DraftRecoveryPane({ host, editor }: { host: HTMLElement; editor: string }) {
  const [update, setUpdate] = useState<DraftRecoveryUpdate | null>(null);
  const [busy, setBusy] = useState<string | null>(null);
  const [error, setError] = useState('');
  const translate = (text: string) => (window as Window & { L?: RecoveryRuntime }).L?.t?.(text) || text;

  useEffect(() => {
    const target = host.ownerDocument.defaultView as DraftRecoveryWindow | null;
    if (!target) return;
    const sync = (event?: Event) => {
      const detail = event ? (event as CustomEvent<DraftRecoveryUpdate>).detail : target.__vibeofficeDraftRecoveries?.[editor];
      if (detail?.editor === editor) setUpdate(detail);
    };
    target.addEventListener('vibeoffice-drafts-update', sync);
    sync();
    return () => target.removeEventListener('vibeoffice-drafts-update', sync);
  }, [host, editor]);

  const discard = async (id: string) => {
    if (!update || busy) return;
    setBusy(id);
    setError('');
    try { await update.discard(id); }
    catch { setError(translate('Could not discard this recovery draft.')); }
    finally { setBusy(null); }
  };

  return <div className="tp-body">
    {!update?.drafts.length && <div className="tp-note">{translate('No recovered documents.')}</div>}
    {!!update?.drafts.length && <section className="tp-sec">
      <div className="tp-h">{translate('Recovered documents')}</div>
      {update.drafts.map((draft) => <div className="tp-row" key={draft.id}>
        <span style={{ flex: 1, minWidth: 0 }}>
          <b>{draft.name}</b><br />
          <small>{new Intl.DateTimeFormat(undefined, { dateStyle: 'medium', timeStyle: 'short' }).format(draft.updatedAt)}</small>
        </span>
        <button className="btn small" type="button" onClick={() => update.restore(draft)}>{translate('Restore')}</button>
        <button className="btn small" type="button" disabled={!!busy} onClick={() => void discard(draft.id)}>{translate('Discard')}</button>
      </div>)}
    </section>}
    {error && <div className="tp-note" role="alert">{error}</div>}
  </div>;
}
