export declare function unescape(s: string): string;
export declare const esc: (s: string) => string;
export declare const escText: (s: string, keepCR?: boolean) => string;
export declare const unx: (s: string) => string;
export interface XmlAttr {
    name: string;
    prefix: string | null;
    localName: string;
    value: string;
}
export declare class Txt {
    data: string;
    parentNode: El | null;
    constructor(data: string, parent: El | null);
    get nodeType(): number;
    get textContent(): string;
    get nodeValue(): string;
    get localName(): null;
    remove(): void;
}
export type XmlNode = El | Txt;
export declare class El {
    nodeName: string;
    prefix: string | null;
    localName: string;
    attrs: Record<string, string> | null;
    children: El[];
    childNodes: XmlNode[];
    parentNode: El | null;
    constructor(name: string, parent: El | null);
    get nodeType(): number;
    get tagName(): string;
    getAttribute(q: string): string | null;
    hasAttribute(q: string): boolean;
    setAttribute(q: string, v: string | number): void;
    removeAttribute(q: string): void;
    get attributes(): XmlAttr[];
    lookupNamespaceURI(prefix: string | null): string | null;
    get namespaceURI(): string | null;
    getAttributeNS(ns: string | null, local: string): string | null;
    get textContent(): string;
    set textContent(v: string);
    get firstElementChild(): El | null;
    get firstChild(): XmlNode | null;
    get parentElement(): El | null;
    getElementsByTagName(name: string): El[];
    getElementsByTagNameNS(ns: string, local: string): El[];
    appendChild(n: XmlNode): XmlNode;
    remove(): void;
    replaceWith(...nodes: XmlNode[]): void;
    cloneNode(deep?: boolean): El;
}
export declare function attrs(s: string): Record<string, string>;
export interface ParseOpts {
    keepWhitespace?: boolean;
    strict?: boolean;
}
/** parse a document; returns the document element */
export declare function parseXml(text: string, opts?: ParseOpts): El;
/** character data between two tags as text: entities resolved, CDATA taken literally, line ends normalised */
export declare function chars(s: string): string;
/** Scan tags between [from, to): fn(local, kind, attrText, textBefore); kind 1 open, 2 close, 3 self-closing. */
export declare function scanXml(text: string, from: number, to: number, fn: (local: string, kind: number, attrText: string, textBefore: string) => void): void;
/** bytes → string honouring a UTF-16 byte-order mark */
export declare function decodeBytes(u8: Uint8Array): string;
/** serialize an element tree back to XML, carrying the in-scope namespace bindings */
export declare function serializeXml(el: El | null): string;
/** create a detached element with a text child */
export declare function mkEl(name: string, attrs?: Record<string, string | number>, text?: string): El;
