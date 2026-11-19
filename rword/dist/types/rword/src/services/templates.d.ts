export interface DocumentTemplate {
    id: string;
    name: string;
    description?: string;
    format?: 'docx' | 'md' | 'html';
    content?: string;
}
export declare class TemplateCatalog {
    private readonly templates;
    constructor(templates?: DocumentTemplate[]);
    list(): DocumentTemplate[];
    get(id: string): DocumentTemplate | undefined;
    create(id: string): string;
}
