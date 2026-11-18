import { type El } from './xml';
export declare const NS: Readonly<{
    rel: "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    strictRel: "http://purl.oclc.org/ooxml/officeDocument/relationships";
    pkg: "http://schemas.openxmlformats.org/package/2006/relationships";
    ct: "http://schemas.openxmlformats.org/package/2006/content-types";
    mc: "http://schemas.openxmlformats.org/markup-compatibility/2006";
    w: "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    p: "http://schemas.openxmlformats.org/presentationml/2006/main";
    a: "http://schemas.openxmlformats.org/drawingml/2006/main";
    s: "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
}>;
export interface FormatInfo {
    ext: string;
    macro: boolean;
    template: boolean;
    family: string;
    contentType: string;
    mime: string;
}
export declare const formats: Record<string, FormatInfo>;
export declare const uuid: () => string;
export interface Rel {
    id: string;
    type: string;
    target: string;
    mode?: string;
}
/** resolve a relationship target against the part that owns it */
export declare function resolveTarget(part: string, target: string): string;
export declare class Package {
    /** raw bytes of every part, keyed by part name (no leading slash) */
    parts: Map<string, Uint8Array<ArrayBufferLike>>;
    /** parts replaced by the editor */
    edited: Set<string>;
    private textCache;
    private xmlCache;
    private relsCache;
    defaults: Map<string, string>;
    overrides: Map<string, string>;
    private ctDirty;
    private relsDirty;
    main: string;
    repaired: boolean;
    static open(bytes: Uint8Array | ArrayBuffer): Promise<Package>;
    has(part: string): boolean;
    bytes(part: string): Uint8Array | undefined;
    text(part: string): string;
    xml(part: string): El;
    type(part: string): string;
    /** relationships of a part ('' = the package root) */
    relsOf(part: string): Rel[];
    rel(part: string, id: string): Rel | undefined;
    /** parts related to `part` by relationship type (resolved names) */
    related(part: string, type: string): string[];
    setPart(part: string, data: string | Uint8Array): void;
    addPart(part: string, data: string | Uint8Array, contentType?: string): void;
    removePart(part: string): void;
    /** add a relationship; returns its new Id (rIdN, as Office numbers them) */
    addRel(part: string, type: string, target: string, mode?: string): string;
    /** next free media part name, as Office numbers them (image1.png, image2.png, ...) */
    nextMediaName(ext: string, stem?: string): string;
    variant(fallback: string): string;
    private writeContentTypes;
    private writeRels;
    save(mime?: string): Promise<Blob>;
}
export declare function contentTypeForExt(ext: string): string;
/** a fresh package with a main part and the officeDocument relationship */
export declare function newPackage(mainPart: string, mainContent: string, format: FormatInfo, extraDefaults?: Record<string, string>): Package;
/** serialize a parsed tree into a full XML part with declaration */
export declare function xmlPart(root: El): string;
