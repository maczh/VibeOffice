declare const api: {
    raw: Record<string, string>;
    get(name: string, size?: number): string;
    /** the application's own icon; each app sets L.icons.app in its app.js (this is the suite's) */
    app(size?: number): string;
    /** more icons: { name: '<svg body, 16×16 user units>' } */
    add(icons: Record<string, string>): void;
    /** drawing helpers and the Luna palette, for apps that add icons */
    kit: {
        page: (x: number, y: number, w: number, hh: number, fill?: string) => string;
        txt: (x: number, y: number, t: string, size: number, col?: string, weight?: string, style?: string) => string;
        lines: (x: number, y: number, w: number | ((i: number) => number), n: number, gap: number, col?: string) => string;
        arrow: (d: string, col?: string) => string;
        pg: (extra?: string) => string;
        tl: (x: number, y: number, w: number, n: number, gap: number) => string;
        cellGrid: (x: number, y: number, w: number, hh: number, n: number, m: number, col?: string) => string;
        B: string;
        B2: string;
        LB: string;
        Y: string;
        YD: string;
        R: string;
        G: string;
        K: string;
        GR: string;
    };
};
export declare const icons: {
    raw: Record<string, string>;
    get(name: string, size?: number): string;
    /** the application's own icon; each app sets L.icons.app in its app.js (this is the suite's) */
    app(size?: number): string;
    /** more icons: { name: '<svg body, 16×16 user units>' } */
    add(icons: Record<string, string>): void;
    /** drawing helpers and the Luna palette, for apps that add icons */
    kit: {
        page: (x: number, y: number, w: number, hh: number, fill?: string) => string;
        txt: (x: number, y: number, t: string, size: number, col?: string, weight?: string, style?: string) => string;
        lines: (x: number, y: number, w: number | ((i: number) => number), n: number, gap: number, col?: string) => string;
        arrow: (d: string, col?: string) => string;
        pg: (extra?: string) => string;
        tl: (x: number, y: number, w: number, n: number, gap: number) => string;
        cellGrid: (x: number, y: number, w: number, hh: number, n: number, m: number, col?: string) => string;
        B: string;
        B2: string;
        LB: string;
        Y: string;
        YD: string;
        R: string;
        G: string;
        K: string;
        GR: string;
    };
};
export type IconKit = typeof api.kit;
export declare const icon: (name: string, size?: number) => string;
export {};
