export declare function crc32(buf: Uint8Array): number;
export declare function inflateRawJS(data: Uint8Array, sizeHint?: number): Uint8Array;
export declare function inflateRaw(data: Uint8Array, size?: number): Promise<Uint8Array>;
export interface ZipEntry {
    name: string;
    method: number;
    csize: number;
    usize: number;
    loff: number;
    crc?: number;
    bytes(): Promise<Uint8Array>;
    text(): Promise<string>;
}
export interface ZipFiles extends Map<string, ZipEntry> {
    repaired?: boolean;
}
export interface ZipError extends Error {
    code?: string;
    encrypted?: boolean;
    bytes?: Uint8Array;
}
export declare function readZip(buffer: Uint8Array | ArrayBuffer): Promise<ZipFiles>;
export interface ZipWriteFile {
    name: string;
    data: Uint8Array | string;
    store?: boolean;
}
export declare function writeZip(files: ZipWriteFile[], mime?: string): Promise<Blob>;
export declare const encode: (s: string) => Uint8Array;
export declare const decode: (b: Uint8Array) => string;
