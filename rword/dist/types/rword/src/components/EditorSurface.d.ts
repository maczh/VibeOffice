export interface EditorSurfaceProps {
    value: string;
    readOnly?: boolean;
    onChange: (value: string) => void;
    onContextMenu?: (x: number, y: number) => void;
}
export declare function EditorSurface({ value, readOnly, onChange, onContextMenu }: EditorSurfaceProps): import("react").JSX.Element;
