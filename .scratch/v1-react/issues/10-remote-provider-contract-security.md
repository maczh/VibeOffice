# 10: Remote storage provider contract and security boundary

**What to build:** A host can connect an editor to a remote file service through a typed provider contract with explicit authentication, origin, revision, and conflict behavior.

**Blocked by:** 01: Vite workspace and editor package contract; 02: Shared React interface foundation; 03: Local file access and independent draft recovery.

**Status:** complete

- [x] A provider can list, read, create, and save files and can report a revision when the service supports one.
- [x] A host can supply authentication through a callback; editor code does not persist credentials.
- [x] Standalone WebDAV credentials can be entered for the session and remain in memory only.
- [x] Deployments configure allowed remote origins and the build's Content Security Policy permits only those configured connections.
- [x] Provider documentation describes the CORS requirements for browser requests.
- [x] Conditional saves detect a changed remote revision; providers without revision metadata report that concurrent changes cannot be detected.
- [x] Fake-provider tests can exercise provider behavior without connecting to production services.

## Comments

- 2026-11-24 security follow-up: Built-in URL and WebDAV fetches now set `redirect: 'error'`, so an allowed endpoint cannot transparently redirect the browser request outside the configured-origin boundary. Added an isolated regression covering both providers. Typecheck, core build, the focused test, and scoped diff check pass; custom host providers remain responsible for their own transport policy.
- 2026-11-24 security follow-up: WebDAV path validation now rejects encoded or multiply encoded separators and dot-traversal segments before issuing fetch, while preserving ordinary encoded filenames. Redirect and encoded-path regressions, the existing workspace/provider suite (35/35), typecheck, core build, and scoped diff check pass.
- 2026-11-24 lifecycle follow-up: WebDAV credentials callbacks now receive the active AbortSignal; the request checks cancellation both before and after credential resolution, so unmounted editors do not start transport afterward. Focused cancellation/path/redirect checks and the workspace/provider suite pass (36/36), along with typecheck, core build, and scoped diff check.
- 2026-11-24 write-target follow-up: WebDAV `create` rejects empty, dot, traversal, slash, backslash, and control-character names; `save` rejects the configured collection root and trailing-slash collection targets. Tests confirm these are rejected before transport while a normal encoded filename can still be created. The focused provider set/workspace suite passes 36/36; typecheck, core build, and scoped diff check pass.
- 2026-11-12: Extended URL and WebDAV security regression coverage: credential-bearing URLs and WebDAV base URLs are rejected, and a sibling path such as `/docs-escape/` cannot pass a `/docs/` base-path boundary check. Workspace package tests pass (17/17), all workspace builds pass, and `git diff --check` is clean.
