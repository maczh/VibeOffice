export interface CommandContext {
    [key: string]: unknown;
}
export interface Command {
    id: string;
    label: string;
    enabled?: (context: CommandContext) => boolean;
    run: (context: CommandContext) => void | Promise<void>;
}
export declare class CommandRegistry {
    private readonly commands;
    register(command: Command): void;
    get(id: string): Command | undefined;
    list(): Command[];
    execute(id: string, context?: CommandContext): Promise<boolean>;
}
export declare class History<T> {
    private readonly limit;
    private past;
    private future;
    constructor(limit?: number);
    push(value: T): void;
    undo(current: T): T | undefined;
    redo(current: T): T | undefined;
    clear(): void;
}
