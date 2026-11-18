import { type OfficeApi, type OfficeProps, type OpenInput, type SaveTarget } from '@voffice/react-shared';
import '../../rshared/src/style.css';
export type { OfficeApi, OfficeProps, OpenInput, SaveTarget };
export interface RPptxProps extends OfficeProps {
    onReady?: () => void;
    onChange?: (dirty: boolean) => void;
    onSave?: (blob: Blob, name: string) => void;
    onError?: (error: Error) => void;
}
export declare const RPptx: import("react").ForwardRefExoticComponent<RPptxProps & import("react").RefAttributes<OfficeApi>>;
export default RPptx;
