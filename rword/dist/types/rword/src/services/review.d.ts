import type { Doc, Para, Run } from '../engine/model';
export interface Revision {
    id: string;
    kind: 'insert' | 'delete' | 'format';
    author: string;
    date: string;
    text: string;
    accepted?: boolean;
}
export interface Comment {
    id: string;
    author: string;
    date: string;
    text: string;
    resolved?: boolean;
}
export declare function addRevision(run: Run, revision: Revision): Run;
export declare function acceptRevision(revision: Revision): Revision;
export declare function rejectRevision(revision: Revision): Revision;
export declare function documentComments(doc: Doc): Comment[];
export declare function plainParagraph(text: string): Para;
