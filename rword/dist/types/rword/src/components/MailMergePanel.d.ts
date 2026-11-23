export interface MailMergePanelProps {
    onClose?: () => void;
    onSelect?: (row: Record<string, string>) => void;
}
export declare function MailMergePanel({ onClose, onSelect }: MailMergePanelProps): import("react").JSX.Element;
