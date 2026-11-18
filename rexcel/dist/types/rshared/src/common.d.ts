import { type ReactNode } from 'react';
export type SuiteApp = 'word' | 'excel' | 'pptx';
export interface RecentFile {
    name: string;
    app: SuiteApp;
    url?: string;
    openedAt: number;
}
export declare function useRecentFiles(): {
    recent: RecentFile[];
    remember: (file: Omit<RecentFile, "openedAt">) => void;
    remove: (file: RecentFile) => void;
};
export type Look = 'classic' | 'paper' | 'paper2016';
export declare function useLook(initial?: Look): {
    look: Look;
    setLook: import("react").Dispatch<import("react").SetStateAction<Look>>;
    layout: string;
};
export declare function useDirtyGuard(dirty: boolean, message?: string): void;
export declare function useI18n(): {
    lang: string;
    setLanguage: (next: string) => void;
    t: (s: string, ...args: (string | number)[]) => string;
};
export declare function usePwa(script?: string): {
    updateAvailable: boolean;
    applyUpdate: () => void;
};
export interface AssistantChoice {
    label: string;
    run?: () => void;
}
export declare function useAssistant(): {
    show: (prompt?: string, choices?: AssistantChoice[]) => void;
    hide: () => void;
    choose: (choice: AssistantChoice) => void;
    visible: boolean;
    prompt: string;
    choices: AssistantChoice[];
};
export declare function LanguageSelect(): import("react").JSX.Element;
export declare function FileDropZone({ accept, onFiles, children }: {
    accept?: string;
    onFiles: (files: File[]) => void;
    children: ReactNode;
}): import("react").JSX.Element;
export declare function StartCenter({ recent, onOpen, onCreate, onRecent, onRemove }: {
    recent: RecentFile[];
    onOpen: () => void;
    onCreate: (app: SuiteApp) => void;
    onRecent: (file: RecentFile) => void;
    onRemove: (file: RecentFile) => void;
}): import("react").JSX.Element;
export declare function PwaUpdate({ script }: {
    script?: string;
}): import("react").JSX.Element | null;
export declare function Assistant({ visible, prompt, choices, onHide, onChoose }: {
    visible: boolean;
    prompt: string;
    choices: AssistantChoice[];
    onHide: () => void;
    onChoose: (choice: AssistantChoice) => void;
}): import("react").JSX.Element | null;
export interface Command {
    id: string;
    label: string;
    keywords?: string[];
    shortcut?: string;
    disabled?: boolean;
    run: () => void | Promise<void>;
}
export interface CommandMenuItem extends Command {
    sep?: boolean;
    sub?: CommandMenuItem[];
}
export declare function commandsFromMenus(menus: {
    items: CommandMenuItem[];
}[]): Command[];
export declare function CommandSearch({ commands, placeholder, onRun }: {
    commands: Command[];
    placeholder?: string;
    onRun?: (command: Command) => void;
}): import("react").JSX.Element;
export type OnlineDialogValue = {
    kind: 'webdav';
    url: string;
    username?: string;
    password?: string;
    token?: string;
} | {
    kind: 's3';
    endpoint: string;
    bucket: string;
    key: string;
    region?: string;
    accessKeyId: string;
    secretAccessKey: string;
    sessionToken?: string;
    forcePathStyle?: boolean;
} | {
    kind: 'presigned';
    url: string;
    putUrl?: string;
} | {
    kind: 'post';
    url: string;
    field?: string;
    fields?: Record<string, string>;
};
export declare function OnlineSourceDialog({ mode, fileName, onSubmit, onClose }: {
    mode: 'open' | 'save';
    fileName?: string;
    onSubmit: (value: OnlineDialogValue) => void;
    onClose: () => void;
}): import("react").JSX.Element;
