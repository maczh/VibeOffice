import { type ReactNode } from 'react';
/** The Office 2003 standard colour grid (public/common/core.js → L.color.STANDARD). */
export declare const STANDARD_COLORS: string[];
/** The highlight colours (public/quire/js/render.js → R.HIGHLIGHT). */
export declare const HIGHLIGHT_COLORS: [string, string][];
/** The name a colour shows as its tooltip, in the UI's language. */
export declare function colorName(hex: string): string;
export interface ColorPick {
    none?: boolean;
    auto?: boolean;
    hex?: string;
}
/** The colour panel (ui.colorMenu): a wide automatic/none button, the standard grid. */
export declare function QColorPanel({ mode, onPick, onMore }: {
    mode: 'plain' | 'fill' | 'line' | 'font';
    onPick: (value: ColorPick) => void;
    onMore?: () => void;
}): ReactNode;
/** The highlight panel: the fifteen highlight colours plus a None button. */
export declare function QHighlightPanel({ onPick }: {
    onPick: (name: string) => void;
}): ReactNode;
/** The border grid: the thirteen border presets the original offers. */
export declare const BORDER_ITEMS: [string, string, string][];
export declare function QBorderGrid({ onPick }: {
    onPick: (key: string) => void;
}): ReactNode;
/** The table grid picker: a growing grid whose label shows the size. */
export declare function QTableGridPicker({ onPick }: {
    onPick: (rows: number, cols: number) => void;
}): ReactNode;
/** The column picker: a row of columns; the label counts them. */
export declare function QColumnPicker({ onPick }: {
    onPick: (columns: number) => void;
}): ReactNode;
/** A line-style list (the original's listPanel): a row of previews with labels. */
export declare function QStyleList({ rows }: {
    rows: {
        html: string;
        label?: string;
        run: () => void;
    }[];
}): ReactNode;
/** The shadow grid: eight shadow presets plus "no shadow". */
export declare function QShadowGrid({ onPick }: {
    onPick: (offset: [number, number] | null) => void;
}): ReactNode;
/** A line preview used by the line-weight and dash-style lists. */
export declare function lineSVG(width: number, dash?: string): string;
/** The AutoShapes gallery: one group of the original's shapes, drawn as they are.
 * Each button carries the original's own preview markup, so the gallery's
 * pictures are the original's, pixel for pixel. */
export declare function QShapeGallery({ group, onPick }: {
    group: string;
    onPick: (geom: string) => void;
}): ReactNode;
