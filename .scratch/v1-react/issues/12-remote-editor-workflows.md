# 12: Remote document workflows in all editors

**What to build:** Users can open, create, edit, and save remote documents through the existing file workflows in Wordx, Excelx, and PPTX.

**Blocked by:** 05: Wordx full feature and format parity; 07: Excelx full feature and format parity; 09: PPTX full feature and format parity; 11: URL, WebDAV, custom POST, and S3 provider integrations; 20: Web-wide legacy brand name cleanup.

**Status:** ready-for-agent

- [ ] Each editor can open documents through URL and configured storage providers.
- [ ] Each editor can create a blank document or template-based document and save it to a selected remote destination.
- [ ] Each editor can save changes back to the current remote file or use Save As for another destination.
- [ ] Each editor prompts before overwriting when a provider detects a changed remote revision.
- [ ] Provider authentication and remote service errors are surfaced without exposing credentials.
- [ ] Remote file operations work when launched through the Start Center and when the editor is embedded by a host.
