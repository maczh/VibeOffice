export interface MarkdownDocument {
    source: string;
    text: string;
}
export declare function readMarkdown(source: string): MarkdownDocument;
export declare function writeMarkdown(text: string): string;
