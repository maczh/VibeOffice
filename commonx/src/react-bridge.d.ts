import { type ReactNode } from 'react';
export type LookId = 'classic' | 'paper' | 'paper2016';
export type LanguageId = 'auto' | 'en' | 'fr' | 'de' | 'zh-CN' | 'zh-Hant' | 'ja' | 'ko';
export interface Choice<T extends string> {
    id: T;
    name: string;
}
export interface VibeOfficeServices {
    looks: readonly Choice<LookId>[];
    languages: readonly Choice<LanguageId>[];
    look: LookId;
    language: LanguageId;
    translate: (message: string, ...values: Array<string | number>) => string;
    setLook: (look: LookId) => void;
    setLanguage: (language: LanguageId) => void;
    execute: (command: string) => void;
}
interface LegacyVibeOffice {
    t?: (message: string, ...values: Array<string | number>) => string;
    look?: {
        set: (look: LookId) => void;
    };
    lang?: {
        set: (language: LanguageId) => void;
    };
    reactDirectHost?: {
        root?: HTMLElement | null;
    };
}
declare global {
    interface Window {
        VO?: LegacyVibeOffice;
    }
}
export interface VibeOfficeProviderProps {
    children: ReactNode;
    services?: Partial<Pick<VibeOfficeServices, 'translate' | 'setLook' | 'setLanguage' | 'execute'>>;
}
export declare function VibeOfficeProvider({ children, services }: VibeOfficeProviderProps): import("react").JSX.Element;
export declare function useVibeOffice(): VibeOfficeServices;
export {};
//# sourceMappingURL=react-bridge.d.ts.map
