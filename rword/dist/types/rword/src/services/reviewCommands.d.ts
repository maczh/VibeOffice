import { type Doc, type Para } from '../engine/model';
import { type CommandFn } from './documentCommands';
export declare const reviewCommands: Record<string, CommandFn>;
/** Called from the component when a tracked insertion is typed. */
export declare function recordInsertion(doc: Doc, blockId: string | undefined, text: string): void;
export declare function ensureEmptyParagraph(doc: Doc): Para;
