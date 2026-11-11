# VibeOffice

VibeOffice edits Office documents in the browser, either as a standalone suite or as editors embedded in another application.

## Files and storage

**Local file**:
A document the user selects from, or saves to, storage on their own device.
_Avoid_: browser file, downloaded file (when referring to the source document)

**Remote file**:
A document the user opens or saves through a URL or a configured remote storage provider.
_Avoid_: cloud file (unless the provider is specifically cloud storage)

**Storage provider**:
An integration that lists, opens, creates, and saves documents in a storage location, and may report a revision for conflict detection.
_Avoid_: storage backend (when referring to the browser-side integration)

## Editors

**Editor package**:
A complete React editor for one document family that can run by itself or be embedded in a host application.
_Avoid_: app component (when referring to the distributable editor)

**Document engine**:
The existing editor-specific model and file-format logic that reads, edits, and writes documents.
_Avoid_: React editor (when referring to document-format logic)
