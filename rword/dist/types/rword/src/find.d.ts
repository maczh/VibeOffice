export interface FindOptions {
    text: string;
    replace?: string;
    matchCase?: boolean;
    wholeWord?: boolean;
    wildcards?: boolean;
    allForms?: boolean;
    soundsLike?: boolean;
}
export interface FindHit {
    index: number;
    end: number;
    value: string;
    groups: string[];
}
export declare function findAll(value: string, options: FindOptions): FindHit[];
export declare function replaceAll(value: string, options: FindOptions): {
    value: string;
    count: number;
};
