# 11: URL, WebDAV, custom POST, and S3 provider integrations

**What to build:** Users and host applications can use the supported remote storage routes without adding a VibeOffice server.

**Blocked by:** 10: Remote storage provider contract and security boundary.

**Status:** complete

- [x] A user can open a document from an allowed URL; URL loading is read-only unless a writable provider is supplied.
- [x] A user can browse a WebDAV location, open a document, create a document, and save it through the provider.
- [x] A host can implement its own service listing, authentication, and POST contract through the provider interface.
- [x] A host can provide S3 object listing and short-lived signed read/write requests without exposing AWS credentials to the editor.
- [x] Provider operations report failures and revision conflicts to the editor in a form that supports user recovery.
