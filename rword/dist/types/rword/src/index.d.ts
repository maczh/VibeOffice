import { type OfficeApi, type OfficeProps, type OpenInput, type SaveTarget } from '@voffice/react-shared';
import '../../rshared/src/style.css';
export type { OfficeApi, OfficeProps, OpenInput, SaveTarget };
export interface RWordProps extends OfficeProps {
    onReady?: () => void;
    onChange?: (dirty: boolean) => void;
    onSave?: (blob: Blob, name: string) => void;
    onError?: (error: Error) => void;
}
export declare const RWord: import("react").ForwardRefExoticComponent<RWordProps & import("react").RefAttributes<OfficeApi>>;
export default RWord;
