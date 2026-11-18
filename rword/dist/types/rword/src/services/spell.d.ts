export interface SpellIssue {
    word: string;
    start: number;
    end: number;
    kind: 'spelling' | 'grammar';
    suggestions: string[];
    message: string;
}
export declare function tokenizeWords(text: string): Array<{
    word: string;
    start: number;
    end: number;
}>;
export declare function grammarIssues(text: string): SpellIssue[];
export declare function spellingIssues(text: string, dictionary: ReadonlySet<string>): SpellIssue[];
export declare function levenshtein(a: string, b: string): number;
export declare function suggest(word: string, dictionary: Iterable<string>, limit?: number): string[];
