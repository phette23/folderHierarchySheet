# Folder hierarchy sheet

Google Apps Script for generating a Drive folder inventory in Google Sheets. Each run creates a new tab with an indented, collapsible folder tree with columns about user access. No content other than the open sheet is modified.

## Setup

1. Open a Google Sheet. Choose **Extensions → Apps Script** to create a spreadsheet-bound script.
2. Replace `Code.gs` with [`src/Code.js`](src/Code.js). The `.js` source works unchanged in the editor's `.gs` file.
3. In **Project Settings**, enable **Show "appsscript.json" manifest file in editor**, then replace that file with [`src/appsscript.json`](src/appsscript.json). Save.
4. Reload the Sheet and select **Folder hierarchy → Create folder report…**. Authorize the script when prompted; rerun the action if authorization interrupts it. For an unverified-app warning, review your installed code before proceeding through **Advanced**. Organization-level blocks require a Workspace administrator.

### Permissions

| Scope | Purpose |
| --- | --- |
| `drive.readonly` | Read folder metadata, child folders, and available sharing information. Applies across accessible Drive content, not just the starting folder. |
| `spreadsheets.currentonly` | Create and format report tabs in the bound spreadsheet. |
| `script.container.ui` | Display prompts and alerts. |

## Usage

The menu action asks for a **folder ID or Drive URL**, then a **recursion depth**: `0` includes only the starting folder, `1` adds immediate children, `2` adds grandchildren, etc. Use the full URL when it includes a `resourcekey`. Collapse or expand descendants with the row-group controls beside each parent; groups start expanded. Sheets supports eight nested group levels; deeper folders remain indented. Files and folder shortcuts are excluded; siblings follow Drive's ordering.

For an optional on-sheet button, use **Insert → Drawing**, then **Assign script → `promptFolderHierarchy`** (without parentheses).

To update an installation, replace `Code.gs`, save, and reload the Sheet; local edits do not sync automatically. The `clasp` tool can be setup to update apps script on the CLI but it requires more setup and a Google Cloud project.

If the depth prompt does not appear after clicking OK, check for an outdated installation or errors in Apps Script's **Executions** page.

## Limitations

- Sharing labels distinguish public, link, organization, and restricted access. Restricted folders may still be shared with specific accounts. Missing owners and empty account lists are blank. Blank account cells do not prove a folder is private: Google can hide emails or return empty lists without edit access. Viewers include commenters.
- Group membership, inherited access, and shared-drive permissions are not fully enumerated. Unavailable fields and child-listing failures appear in Notes. See Google's [Folder API reference](https://developers.google.com/apps-script/reference/drive/folder).
- Runs stop at 3,000 folders or approximately four minutes and mark the report **PARTIAL**. Use smaller subtrees for larger inventories. Hard timeouts or quota errors can leave unfinished tabs; there is no automatic resume.
- Not built to update a previous run, but only to perform idempotent runs.

## Development and tests

```sh
pnpm install
pnpm lint       # ESLint recommended rules, Apps Script globals
pnpm test
pnpm validate   # lint + tests
```

The 23 tests load the actual `src/Code.js` functions into an isolated Node VM with mocked Apps Script services. No live Google APIs, credentials, or network access are used; time-limit tests use a fake clock.

## License

[CC0 Public Domain](https://creativecommons.org/publicdomain/zero/1.0/)
