export interface MarkdownDocument { source: string; text: string }
export function readMarkdown(source: string): MarkdownDocument { return { source, text: source.replace(/^\uFEFF/, '').replace(/\r\n?/g, '\n') } }
export function writeMarkdown(text: string): string { return text.replace(/\r\n?/g, '\n').replace(/\n?$/, '\n') }
