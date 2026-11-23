import type { RPr } from '../engine';
export interface FormattingPanelProps {
    onApply?: (style: RPr) => void;
    onClose?: () => void;
}
export declare function FormattingPanel({ onApply, onClose }: FormattingPanelProps): import("react").JSX.Element;
