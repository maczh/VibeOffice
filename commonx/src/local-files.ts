export interface LocalDocument {
  file: File;
  handle?: FileSystemFileHandle;
  remote?: { provider: Pick<import('./storage.js').RemoteStorageProvider, 'read'> & Partial<Pick<import('./storage.js').RemoteStorageProvider, 'save' | 'create'>>; file: import('./storage.js').RemoteFile };
}

export interface LocalDraft<T = unknown> {
  id: string;
  editor: string;
  name: string;
  updatedAt: number;
  snapshot: T;
  lossState?: unknown;
}

export interface DraftStore<T = unknown> {
  list(editor?: string, signal?: AbortSignal): Promise<Array<LocalDraft<T>>>;
  put(draft: LocalDraft<T>, signal?: AbortSignal): Promise<void>;
  remove(id: string, signal?: AbortSignal): Promise<void>;
}

export type UnsavedChoice = 'save' | 'discard' | 'cancel';

export interface UnsavedChangeActions {
  isDirty(): boolean;
  choose(): Promise<UnsavedChoice>;
  save(): Promise<void>;
  discard(): Promise<void>;
}

/** Resolve a pending edit before replacing or closing a local document. A cancelled or failed save preserves it. */
export async function resolveUnsavedChanges(actions: UnsavedChangeActions): Promise<boolean> {
  if (!actions.isDirty()) return true;
  const choice = await actions.choose();
  if (choice === 'cancel') return false;
  if (choice === 'save') {
    await actions.save();
    return !actions.isDirty();
  }
  await actions.discard();
  return true;
}

export interface OpenFilePickerOptions {
  multiple?: boolean;
  types?: Array<{ description?: string; accept: Record<string, Array<string>> }>;
}

export interface SaveFilePickerOptions {
  suggestedName?: string;
}

export type FilePickerWindow = Window & {
  showOpenFilePicker?: (options?: OpenFilePickerOptions) => Promise<Array<FileSystemFileHandle>>;
  showSaveFilePicker?: (options?: SaveFilePickerOptions) => Promise<FileSystemFileHandle>;
};

const draftDatabase = 'vibeoffice-react-drafts';
const draftStore = 'documents';

function openDraftDatabase(indexedDBFactory: IDBFactory): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    const request = indexedDBFactory.open(draftDatabase, 1);
    request.onupgradeneeded = () => {
      if (!request.result.objectStoreNames.contains(draftStore)) {
        request.result.createObjectStore(draftStore, { keyPath: 'id' });
      }
    };
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error ?? new Error('Could not open local draft storage'));
  });
}

function abortError(message = 'The operation was cancelled'): DOMException { return new DOMException(message, 'AbortError'); }

function withStore<T>(database: IDBDatabase, mode: IDBTransactionMode, run: (store: IDBObjectStore) => IDBRequest<T>, signal?: AbortSignal): Promise<T> {
  if (signal?.aborted) return Promise.reject(abortError());
  return new Promise((resolve, reject) => {
    const transaction = database.transaction(draftStore, mode);
    const request = run(transaction.objectStore(draftStore));
    let result: T;
    const cleanup = () => signal?.removeEventListener('abort', onAbort);
    const onAbort = () => { try { transaction.abort(); } catch { /* already completed */ } };
    signal?.addEventListener('abort', onAbort, { once: true });
    request.onsuccess = () => { result = request.result; };
    request.onerror = () => { cleanup(); reject(request.error ?? new Error('Local draft storage request failed')); };
    transaction.oncomplete = () => { cleanup(); resolve(result); };
    transaction.onerror = () => { cleanup(); reject(transaction.error ?? new Error('Local draft storage transaction failed')); };
    transaction.onabort = () => { cleanup(); reject(signal?.aborted ? abortError() : transaction.error ?? new Error('Local draft storage transaction failed')); };
  });
}

export function createIndexedDbDraftStore<T = unknown>(factory: IDBFactory = indexedDB): DraftStore<T> {
  let database: Promise<IDBDatabase> | undefined;
  const db = () => database ??= openDraftDatabase(factory);

  return {
    async list(editor, signal) {
      const entries = await withStore(await db(), 'readonly', (store) => store.getAll(), signal) as Array<LocalDraft<T>>;
      return entries.filter((entry) => !editor || entry.editor === editor).sort((a, b) => b.updatedAt - a.updatedAt);
    },
    async put(draft, signal) { await withStore(await db(), 'readwrite', (store) => store.put(draft), signal); },
    async remove(id, signal) { await withStore(await db(), 'readwrite', (store) => store.delete(id), signal); },
  };
}

export async function openLocalDocument(accept?: Record<string, Array<string>>, target: FilePickerWindow = window, signal?: AbortSignal): Promise<LocalDocument> {
  const abort = () => new DOMException('File selection was cancelled', 'AbortError');
  if (signal?.aborted) throw abort();
  if (target.showOpenFilePicker) {
    const pending = target.showOpenFilePicker({ multiple: false, ...(accept ? { types: [{ description: 'Office documents', accept }] } : {}) });
    const handles = signal ? await new Promise<Array<FileSystemFileHandle>>((resolve, reject) => {
      const onAbort = () => { signal.removeEventListener('abort', onAbort); reject(abort()); };
      signal.addEventListener('abort', onAbort, { once: true });
      pending.then((value) => { signal.removeEventListener('abort', onAbort); resolve(value); }, (error: unknown) => { signal.removeEventListener('abort', onAbort); reject(error); });
    }) : await pending;
    const [handle] = handles;
    if (!handle) throw new Error('No file was selected');
    if (signal?.aborted) throw abort();
    return { file: await handle.getFile(), handle };
  }

  const file = await new Promise<File>((resolve, reject) => {
    const input = document.createElement('input');
    input.type = 'file';
    if (accept) input.accept = Object.values(accept).flat().join(',');
    let onChange: EventListener;
    let onCancel: EventListener;
    const finish = (error?: Error, selected?: File) => {
      signal?.removeEventListener('abort', onAbort);
      input.removeEventListener('change', onChange);
      input.removeEventListener('cancel', onCancel);
      input.remove();
      if (error) reject(error); else if (selected) resolve(selected); else reject(new Error('No file was selected'));
    };
    const onAbort = () => finish(abort());
    onChange = () => finish(undefined, input.files?.[0]);
    onCancel = onAbort;
    input.addEventListener('change', onChange, { once: true });
    input.addEventListener('cancel', onCancel, { once: true });
    signal?.addEventListener('abort', onAbort, { once: true });
    input.click();
  });
  return { file };
}

export async function saveLocalDocument(
  document: LocalDocument,
  blob: Blob,
  suggestedName: string,
  target: FilePickerWindow = window,
  signal?: AbortSignal,
): Promise<LocalDocument> {
  if (signal?.aborted) throw abortError();
  let handle = document.handle;
  if (!handle && target.showSaveFilePicker) {
    const pending = target.showSaveFilePicker({ suggestedName });
    handle = signal ? await new Promise<FileSystemFileHandle>((resolve, reject) => {
      const onAbort = () => { signal.removeEventListener('abort', onAbort); reject(abortError('File save was cancelled')); };
      signal.addEventListener('abort', onAbort, { once: true });
      pending.then((value) => { signal.removeEventListener('abort', onAbort); resolve(value); }, (error: unknown) => { signal.removeEventListener('abort', onAbort); reject(error); });
    }) : await pending;
    if (signal?.aborted) throw abortError('File save was cancelled');
  }

  if (handle) {
    const writable = await handle.createWritable();
    const onAbort = () => { void writable.abort().catch(() => {}); };
    signal?.addEventListener('abort', onAbort, { once: true });
    try {
      if (signal?.aborted) throw abortError('File save was cancelled');
      await writable.write(blob);
      if (signal?.aborted) throw abortError('File save was cancelled');
      await writable.close();
    } catch (error) {
      if (signal?.aborted) { try { await writable.abort(); } catch { /* the stream may already be closed */ } }
      throw error;
    } finally { signal?.removeEventListener('abort', onAbort); }
    return { file: await handle.getFile(), handle };
  }

  if (signal?.aborted) throw abortError('File save was cancelled');
  const url = URL.createObjectURL(blob);
  const anchor = window.document.createElement('a');
  anchor.href = url;
  anchor.download = suggestedName;
  anchor.hidden = true;
  window.document.body.append(anchor);
  anchor.click();
  anchor.remove();
  setTimeout(() => URL.revokeObjectURL(url), 0);
  return { file: new File([blob], suggestedName, { type: blob.type, lastModified: Date.now() }) };
}
