import { installSharedI18n, installSharedLanguage, installSharedLook } from '../../commonx/src/runtime-services.js';

declare global {
  interface Window {
    checkWorkspaceRuntimeServices(): Promise<Record<string, unknown>>;
    __VO_HOST__?: Record<string, unknown>;
    __TAURI_INTERNALS__?: { invoke: (...args: unknown[]) => Promise<unknown> };
    __VO_SETTINGS__?: Record<string, string>;
    VO?: any;
  }
}

window.checkWorkspaceRuntimeServices = async () => {
  let desktopInvocations = 0;
  window.__VO_HOST__ = {};
  window.__TAURI_INTERNALS__ = { invoke: async () => { desktopInvocations++; throw new Error('workspace must not invoke desktop APIs'); } };
  window.__VO_SETTINGS__ = { 'vibeoffice.look': 'paper' };
  installSharedLook(window);
  installSharedLanguage(window);
  installSharedI18n(window, new URL('/common/i18n.js', location.href));
  const key = 'vibeoffice.workspace-runtime-test';
  localStorage.removeItem(key);
  const stored = window.VO!.settings.set(key, 'browser-storage');
  await window.VO!.settings.kept();
  const settings = {
    stored,
    value: localStorage.getItem(key),
    look: window.VO!.look.get(),
    desktopInvocations,
    desktopSettingsUntouched: window.__VO_SETTINGS__?.['vibeoffice.look'] === 'paper',
  };
  const brand = window.VO!.t('Welcome to Quire / Ledger / Lectern');
  localStorage.removeItem(key);
  return { settings, brand };
};
