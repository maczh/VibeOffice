export interface FieldInstruction {
    name: string;
    args: string[];
    switches: Record<string, string | true>;
}
export declare function parseField(instruction: string): FieldInstruction;
export interface FieldContext {
    page?: number;
    pages?: number;
    title?: string;
    author?: string;
    fileName?: string;
    now?: Date;
}
export declare function evaluateField(instruction: string, context?: FieldContext): string;
