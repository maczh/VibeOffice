import { createContext, useContext, useEffect, useMemo, useState, type ReactNode } from 'react';

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
  N?: (message: string) => string;
  reactClose?: () => Promise<boolean>;
  look?: { get?: () => LookId; set: (look: LookId) => void };
  lang?: { set: (language: LanguageId) => void };
  reactDirectHost?: { root?: HTMLElement | null };
}

declare global {
  interface Window { VO?: LegacyVibeOffice }
}

export interface VibeOfficeProviderProps {
  children: ReactNode;
  services?: Partial<Pick<VibeOfficeServices, 'translate' | 'setLook' | 'setLanguage' | 'execute'>>;
}

export type VibeOfficeServiceOverrides = VibeOfficeProviderProps['services'];

const looks: readonly Choice<LookId>[] = [
  { id: 'classic', name: 'Classic' },
  { id: 'paper', name: 'Paper' },
  { id: 'paper2016', name: 'Paper 2016' },
];

const languages: readonly Choice<LanguageId>[] = [
  { id: 'auto', name: 'Auto' },
  { id: 'en', name: 'English' },
  { id: 'fr', name: 'Français' },
  { id: 'de', name: 'Deutsch' },
  { id: 'zh-CN', name: '简体中文' },
  { id: 'zh-Hant', name: '繁體中文' },
  { id: 'ja', name: '日本語' },
  { id: 'ko', name: '한국어' },
];

const Context = createContext<VibeOfficeServices | null>(null);

function currentLook(): LookId {
  const value = typeof window !== 'undefined' ? window.VO?.look?.get?.() : undefined;
  return looks.some(({ id }) => id === value) ? value as LookId : 'classic';
}

function currentLanguage(): LanguageId {
  const value = typeof document === 'undefined' ? undefined : document.documentElement.lang;
  return languages.some(({ id }) => id === value) ? value as LanguageId : 'en';
}

function defaultTranslate(message: string, ...values: Array<string | number>): string {
  const translator = typeof window === 'undefined' ? undefined : window.VO?.t;
  return typeof translator === 'function' ? translator(message, ...values) : message.replace(/\{(\d+)\}/g, (match, index: string) => String(values[Number(index)] ?? match));
}

function defaultSetLook(look: LookId): void {
  if (typeof window === 'undefined' || typeof document === 'undefined') return;
  if (window.VO?.look) window.VO.look.set(look);
  else {
    document.documentElement.dataset.look = look;
    document.documentElement.dataset.layout = look === 'paper2016' ? 'ribbon' : 'bars';
    window.dispatchEvent(new Event('vo-look'));
  }
}

function defaultSetLanguage(language: LanguageId): void {
  if (typeof window === 'undefined') return;
  if (window.VO?.lang) window.VO.lang.set(language);
  else {
    try { localStorage.setItem('vibeoffice.lang', language); } catch { /* Host storage may be unavailable. */ }
    window.dispatchEvent(new CustomEvent('vo-language', { detail: language }));
  }
}

function defaultExecute(command: string): void {
  if (typeof window !== 'undefined') window.dispatchEvent(new CustomEvent('vo-command', { detail: command }));
}

export function VibeOfficeProvider({ children, services }: VibeOfficeProviderProps) {
  const [look, setLookState] = useState<LookId>(() => currentLook());
  const [language, setLanguageState] = useState<LanguageId>(() => currentLanguage());

  useEffect(() => {
    if (typeof window === 'undefined') return;
    const refreshLook = () => setLookState(currentLook());
    const refreshLanguage = () => setLanguageState(currentLanguage());
    window.addEventListener('vo-look', refreshLook);
    window.addEventListener('storage', refreshLook);
    window.addEventListener('vo-language', refreshLanguage);
    window.addEventListener('storage', refreshLanguage);
    return () => {
      window.removeEventListener('vo-look', refreshLook);
      window.removeEventListener('storage', refreshLook);
      window.removeEventListener('vo-language', refreshLanguage);
      window.removeEventListener('storage', refreshLanguage);
    };
  }, []);

  const value = useMemo<VibeOfficeServices>(() => ({
    looks,
    languages,
    look,
    language,
    translate: services?.translate ?? defaultTranslate,
    setLook: (next) => (services?.setLook ?? defaultSetLook)(next),
    setLanguage: (next) => (services?.setLanguage ?? defaultSetLanguage)(next),
    execute: services?.execute ?? defaultExecute,
  }), [look, language, services]);

  return <Context.Provider value={value}>{children}</Context.Provider>;
}

export function useVibeOffice(): VibeOfficeServices {
  const value = useContext(Context);
  if (!value) throw new Error('useVibeOffice must be used within VibeOfficeProvider');
  return value;
}
