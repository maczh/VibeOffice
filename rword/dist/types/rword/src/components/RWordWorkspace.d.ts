import { type RWordProps } from '../index';
export type WorkspacePanel = 'none' | 'task' | 'mailmerge' | 'review';
export interface RWordWorkspaceProps extends RWordProps {
    initialPanel?: WorkspacePanel;
}
/** Composed React workspace for applications that need Quire's editor and side panels together. */
export declare function RWordWorkspace({ initialPanel, ...props }: RWordWorkspaceProps): import("react").JSX.Element;
