export interface PreservedPart {
    name: string;
    bytes: Uint8Array;
    contentType?: string;
}
export declare function preserveUnknownPart(parts: Map<string, Uint8Array>, name: string, bytes: Uint8Array): Map<string, Uint8Array>;
export declare function clonePreservedParts(parts: Iterable<PreservedPart>): PreservedPart[];
