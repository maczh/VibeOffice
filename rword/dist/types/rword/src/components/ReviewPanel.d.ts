import type { Comment, Revision } from '../engine';
export interface ReviewPanelProps {
    revisions?: Revision[];
    comments?: Comment[];
    onRevision?: (revision: Revision) => void;
    onClose?: () => void;
}
export declare function ReviewPanel({ revisions, comments, onRevision, onClose }: ReviewPanelProps): import("react").JSX.Element;
