# Issue tracker: Local Markdown

Issues and specs for this repo live as markdown files in `.scratch/`.

## Conventions

- One feature per directory: `.scratch/<feature-slug>/`
- The spec is `.scratch/<feature-slug>/spec.md`
- Implementation issues are one file per ticket at `.scratch/<feature-slug>/issues/<NN>-<slug>.md`, numbered from `01`, never a single combined tickets file
- Triage state is recorded as a `Status:` line near the top of each issue file (see `triage-labels.md` for the role strings)
- Comments and conversation history append to the bottom of the file under a `## Comments` heading

## When a skill says "publish to the issue tracker"

Create a new file under `.scratch/<feature-slug>/` (creating the directory if needed).

## When a skill says "fetch the relevant ticket"

Read the file at the referenced path. The user will normally pass the path or the issue number directly.

## Wayfinding operations

Used by `/wayfinder`. The map is `.scratch/<effort>/map.md`, with one child file per ticket.

- Child tickets live in `.scratch/<effort>/issues/NN-<slug>.md`
- Each ticket has a `Type:` line and a `Status:` line
- Blocking is recorded with `Blocked by: NN, NN`
- The first open, unblocked, unclaimed ticket is the frontier
- Claim by setting `Status: claimed`
- Resolve by appending an `## Answer` section and setting `Status: resolved`
