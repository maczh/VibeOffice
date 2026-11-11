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
  list(editor?: string): Promise<Array<LocalDraft<T>>>;
  put(draft: LocalDraft<T>): Promise<void>;
  remove(id: string): Promise<void>;
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

function withStore<T>(database: IDBDatabase, mode: IDBTransactionMode, run: (store: IDBObjectStore) => IDBRequest<T>): Promise<T> {
  return new Promise((resolve, reject) => {
    const transaction = database.transaction(draftStore, mode);
    const request = run(transaction.objectStore(draftStore));
    let result: T;
    request.onsuccess = () => { result = request.result; };
    request.onerror = () => reject(request.error ?? new Error('Local draft storage request failed'));
    transaction.oncomplete = () => resolve(result);
    transaction.onerror = transaction.onabort = () => reject(transaction.error ?? new Error('Local draft storage transaction failed'));
  });
}

export function createIndexedDbDraftStore<T = unknown>(factory: IDBFactory = indexedDB): DraftStore<T> {
  let database: Promise<IDBDatabase> | undefined;
  const db = () => database ??= openDraftDatabase(factory);

  return {
    async list(editor) {
      const entries = await withStore(await db(), 'readonly', (store) => store.getAll()) as Array<LocalDraft<T>>;
      return entries.filter((entry) => !editor || entry.editor === editor).sort((a, b) => b.updatedAt - a.updatedAt);
    },
    async put(draft) { await withStore(await db(), 'readwrite', (store) => store.put(draft)); },
    async remove(id) { await withStore(await db(), 'readwrite', (store) => store.delete(id)); },
  };
}

export async function openLocalDocument(accept?: Record<string, Array<string>>, target: FilePickerWindow = window): Promise<LocalDocument> {
  if (target.showOpenFilePicker) {
    const [handle] = await target.showOpenFilePicker({ multiple: false, ...(accept ? { types: [{ description: 'Office documents', accept }] } : {}) });
    if (!handle) throw new Error('No file was selected');
    return { file: await handle.getFile(), handle };
  }

  const file = await new Promise<File>((resolve, reject) => {
    const input = document.createElement('input');
    input.type = 'file';
    if (accept) input.accept = Object.values(accept).flat().join(',');
    input.addEventListener('change', () => input.files?.[0] ? resolve(input.files[0]) : reject(new Error('No file was selected')), { once: true });
    input.addEventListener('cancel', () => reject(new DOMException('File selection was cancelled', 'AbortError')), { once: true });
    input.click();
  });
  return { file };
}

export async function saveLocalDocument(
  document: LocalDocument,
  blob: Blob,
  suggestedName: string,
  target: FilePickerWindow = window,
): Promise<LocalDocument> {
  let handle = document.handle;
  if (!handle && target.showSaveFilePicker) {
    handle = await target.showSaveFilePicker({ suggestedName });
  }

  if (handle) {
    const writable = await handle.createWritable();
    await writable.write(blob);
    await writable.close();
    return { file: await handle.getFile(), handle };
  }

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
