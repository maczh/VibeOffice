import { type OfficeApi, type OfficeProps, type OpenInput, type SaveTarget } from '@voffice/react-shared';
import '../../public/common/looks.css';
import '../../public/common/ui.css';
import '../../public/fonts/fonts.css';
import './quire.css';
export { quireScriptMigration } from './quire-manifest';
export { TaskPane, type TaskPaneProps } from './components/TaskPane';
export { EditorSurface, type EditorSurfaceProps } from './components/EditorSurface';
export { Ruler, type RulerProps } from './components/Ruler';
export { MailMergePanel, type MailMergePanelProps } from './components/MailMergePanel';
export { ReviewPanel, type ReviewPanelProps } from './components/ReviewPanel';
export { RWordWorkspace, type RWordWorkspaceProps, type WorkspacePanel } from './components/RWordWorkspace';
export { OutlinePanel, type OutlinePanelProps } from './components/OutlinePanel';
export { FormattingPanel, type FormattingPanelProps } from './components/FormattingPanel';
export { CommandSurface, type CommandSurfaceProps } from './components/CommandSurface';
export type { OfficeApi, OfficeProps, OpenInput, SaveTarget };
export interface RWordProps extends OfficeProps {
    onReady?: () => void;
    onChange?: (dirty: boolean) => void;
    onSave?: (blob: Blob, name: string) => void;
    onError?: (error: Error) => void;
}
/** An action the interface can run; some take a value, such as a colour or a size. */
export type ActionFn = (arg?: string) => void;
export declare const RWord: import("react").ForwardRefExoticComponent<RWordProps & import("react").RefAttributes<OfficeApi>>;
export default RWord;
