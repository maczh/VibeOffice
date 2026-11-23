import { type ImageRef } from '../engine/model';
import { type CommandEnv, type CommandFn } from './documentCommands';
/** The pen the border buttons draw with: style, width in points and colour. */
export interface BorderPen {
    val: string;
    sz: number;
    color: string;
}
export declare const borderPen: BorderPen;
/** Apply a border preset to the paragraph at the caret, or to the table cell. */
export declare function applyBorder(env: CommandEnv, key: string): string | undefined;
/** The shading colour of the paragraph or cell at the caret. */
export declare function applyShading(env: CommandEnv, hex: string | null): string | undefined;
/** Apply a character format to every run of the paragraph at the caret. */
export declare function formatRuns(env: CommandEnv, patch: Record<string, unknown>): string | undefined;
/** The font colour of the caret: `auto` restores the inherited colour. */
export declare function applyFontColor(env: CommandEnv, hex: string | null): string | undefined;
/** The highlight colour of the caret; `none` removes it. */
export declare function applyHighlight(env: CommandEnv, name: string): string | undefined;
/** The line spacing of the caret, as the line-spacing menu shows it. */
export declare function currentLineSpacing(env: CommandEnv): number;
/** Set the line spacing of the caret. */
export declare function setLineSpacing(env: CommandEnv, value: number): string | undefined;
/** Insert a table of the given size at the caret. */
export declare function insertTableAt(env: CommandEnv, rows: number, cols: number): string | undefined;
/** Set the number of newspaper-style columns in the section. */
export declare function setColumns(env: CommandEnv, count: number): string | undefined;
/** The picture the drawing and picture panels act on. */
export declare function selectedImage(env: CommandEnv): ImageRef | null;
/** The text-wrapping modes the original offers, with their labels. */
export declare const WRAP_MODES: [string, string][];
/** The LettersArt shapes the original offers, with their labels. */
export declare const WORDART_SHAPES: [string, string][];
/** The border styles the original offers, with the dash pattern each draws. */
export declare const BORDER_STYLES: [string, string | null][];
/** The line weights the original offers, in points. */
export declare const LINE_WEIGHTS: number[];
/** The dash styles the original offers for drawing objects. */
export declare const DASH_STYLES: [string, string | null][];
/** The AutoText entries the original offers in the header and footer toolbar.
 *  The names are the original's own (public/quire/js/autocorrect.js → AUTOTEXT['Header/Footer']),
 *  so they translate through the same dictionary; the text is the field sequence they insert. */
export declare const HEADER_FOOTER_AUTOTEXT: [string, string][];
/** Insert an AutoText entry into the header or the footer. */
export declare function insertAutoText(env: CommandEnv, text: string): string | undefined;
export declare const toolbarCommands: Record<string, CommandFn>;
