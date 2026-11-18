export interface RPr {
    bold?: boolean;
    italic?: boolean;
    underline?: boolean;
    size?: number;
    color?: string;
    font?: string;
}
export interface Run {
    text: string;
    rPr?: RPr;
}
export interface Para {
    align?: 'left' | 'center' | 'right';
    runs: Run[];
    bullet?: boolean;
    lvl?: number;
}
export interface Fill {
    color: string;
    alpha?: number;
}
export interface Line {
    color: string;
    width?: number;
    dash?: string;
}
export interface Shape {
    id: string;
    kind: 'sp' | 'pic' | 'cxnSp' | 'graphicFrame';
    /** EMU */
    x: number;
    y: number;
    w: number;
    h: number;
    rot?: number;
    preset?: string;
    fill?: Fill | null;
    line?: Line | null;
    text?: Para[];
    image?: {
        part: string;
        rId: string;
        url: string;
    };
    raw?: string;
    name?: string;
}
export interface Slide {
    id: string;
    shapes: Shape[];
    notes?: string;
    background?: string;
    part?: string;
    raw?: string;
}
export interface Presentation {
    slideW: number;
    slideH: number;
    slides: Slide[];
    format: string;
}
export declare const uid: (p?: string) => string;
export declare const EMU_PER_PX: number;
export declare const px: (emu: number) => number;
export declare const emu: (pxv: number) => number;
export declare const newSlide: () => Slide;
export declare const newPresentation: () => Presentation;
export declare function newTextShape(x: number, y: number, w: number, h: number, text?: string): Shape;
export declare const clonePresentation: (p: Presentation) => Presentation;
