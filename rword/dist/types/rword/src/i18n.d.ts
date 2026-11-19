declare const LANGUAGES: readonly ["en", "fr", "de", "zh-CN", "zh-Hant", "ja", "ko", "qps"];
type Language = (typeof LANGUAGES)[number];
export declare function normalizeLanguage(tag: string | null | undefined): Language | null;
export declare function initLanguage(): Language;
/** Same contract as VO.t / L.t: a missing translation returns the English source. */
export declare function t(text: string, ...args: (string | number)[]): string;
/** Strip an access-key ampersand from translated text, as common/ui.js does for tooltips. */
export declare function plainAmp(text: string): string;
export {};
