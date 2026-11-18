export interface TaskPaneProps {
    text: string;
    onNew: () => void;
    onOpen: () => void;
    onClose: () => void;
}
/** React counterpart of Quire's Getting Started, Help and Search task panes. */
export declare function TaskPane({ text, onNew, onOpen, onClose }: TaskPaneProps): import("react").JSX.Element;
