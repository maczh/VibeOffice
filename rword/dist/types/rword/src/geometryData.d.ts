export interface ShapePart {
    d: string;
    fill: string;
    stroke: boolean;
    rule: string;
}
export interface ShapePreset {
    /** the preset name (the OOXML prstGeom vocabulary) */
    geom: string;
    /** the original's English label, shown as the gallery's tooltip */
    label: string;
    /** a line preset draws a stroke only */
    isLine: boolean;
    /** the gallery group the original lists it under */
    group: string;
    /** the original's own preview markup for this preset, at its gallery size */
    preview: string;
    /** the SVG path parts, in a 100×100 box */
    parts: ShapePart[];
}
/** the gallery's groups, in the original's order */
export declare const SHAPE_GROUPS: string[];
export declare const SHAPE_PRESETS: ShapePreset[];
export declare const SHAPE_BY_GEOM: Record<string, ShapePreset>;
/** the presets of one gallery group, in the original's order */
export declare function shapesInGroup(group: string): ShapePreset[];
