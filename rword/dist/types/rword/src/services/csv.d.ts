export interface DelimitedData {
    fields: string[];
    rows: string[][];
}
export declare function parseDelimited(text: string): DelimitedData;
export declare function csvEscape(value: string): string;
export declare function writeDelimited(data: DelimitedData, delimiter?: string): string;
