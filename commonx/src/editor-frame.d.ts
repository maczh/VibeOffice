import type { HTMLAttributes, ReactNode } from 'react';
export type EditorKind = 'wordx' | 'excelx' | 'pptx';
export interface EditorFrameProps extends Omit<HTMLAttributes<HTMLElement>, 'children'> {
    appName: string;
    editor: EditorKind;
    children?: ReactNode;
}
export declare function EditorFrame({ appName, editor, className, children, ...attributes }: EditorFrameProps): import("react").JSX.Element;
//# sourceMappingURL=editor-frame.d.ts.map