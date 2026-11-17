import type { OnlineSource, SaveTarget } from './online/types';
import type { Look } from './ui/look';
export interface OfficeProps {
    /** a local File/Blob, an http(s) URL, or an online source to open on mount */
    src?: string | File | Blob | OnlineSource;
    fileName?: string;
    look?: Look;
    readonly?: boolean;
}
export interface OfficeEmits {
    ready: [];
    change: [dirty: boolean];
    /** the user asked to save: the component produced the file; the host may keep it */
    save: [blob: Blob, name: string];
    error: [err: Error];
    /** the user chose File > Open / Save As and the host should show its own picker */
    openRequest: [];
    saveRequest: [];
}
export interface OfficeApi {
    /** open a local file, a URL or an online source */
    open(input: File | Blob | string | OnlineSource, name?: string): Promise<void>;
    /** start a blank document */
    newDocument(): void;
    /** the current document as a file blob (same format it was opened in) */
    getBlob(): Promise<Blob>;
    /** save to the source it came from when that is writable; else emits saveRequest */
    save(): Promise<void>;
    saveAs(name?: string): Promise<Blob>;
    saveTo(target: SaveTarget, name?: string): Promise<void>;
    readonly name: string;
    readonly dirty: boolean;
}
export type { OnlineSource, SaveTarget, Look };
