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

- 2026-11-12: Extended URL and WebDAV security regression coverage: credential-bearing URLs and WebDAV base URLs are rejected, and a sibling path such as `/docs-escape/` cannot pass a `/docs/` base-path boundary check. Workspace package tests pass (17/17), all workspace builds pass, and `git diff --check` is clean.
