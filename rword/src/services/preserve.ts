export interface PreservedPart { name: string; bytes: Uint8Array; contentType?: string }
export function preserveUnknownPart(parts: Map<string, Uint8Array>, name: string, bytes: Uint8Array): Map<string, Uint8Array> { if (!parts.has(name)) parts.set(name, bytes); return parts }
export function clonePreservedParts(parts: Iterable<PreservedPart>): PreservedPart[] { return Array.from(parts, part => ({ ...part, bytes: part.bytes.slice() })) }
