export interface CommandMeta {
    id: string;
    label?: string;
    icon?: string;
    tip?: string;
    key?: string;
}
export declare const COMMAND_META: CommandMeta[];
export declare const META_BY_ID: Record<string, CommandMeta>;
