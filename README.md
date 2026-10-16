# Rubric to Canvas CSV

A small, private, in-browser tool that turns the rubric in a finished Rubric-Based Assessment (RBA) Word document into a CSV that Canvas can import as a rubric.

It does one job: load an RBA, check it, export the CSV, and remind you what to fix back in the RBA.

## Privacy

Everything happens in your browser tab.

- No uploads, accounts, analytics or telemetry.
- No third-party scripts, fonts, CDNs or services.
- The page ships with a Content Security Policy that blocks all network requests (`default-src 'none'`, `connect-src 'none'`, `form-action 'none'`).
- The automated tests fail the build if the HTML contains network APIs (`fetch`, `XMLHttpRequest`, `WebSocket`, `sendBeacon`, and so on), browser storage, or any external resource reference.
- The app does not save your rubric anywhere. Close the tab and it is gone. It never changes your Word file.

You can confirm this yourself: open the browser's network panel, load a file, and watch. There should be no requests after the page itself loads.

## How to use it

**Offline (recommended for the team):** download `rubric-to-canvas.html` from the latest release, save it anywhere, and double-click it. It works from disk (`file://`) with no internet.

**Online:** open the GitHub Pages site for this repository. It is the same file.

1. **Load** the finished RBA: choose or drop the `.docx`, paste the whole document, paste just the rubric table, or load a CSV/TSV of the rubric table.
2. **Check** the summary and fix anything flagged. Text, points and the rubric name are editable.
3. **Export**: Download CSV (or copy it), then hand it to AIT. Read the **"Before you hand off"** list and make the same fixes in the RBA.

Supported browsers: current Chrome, Edge, Firefox and Safari. The app uses the built-in `DecompressionStream` API to read `.docx` files (Chrome/Edge 103+, Firefox 113+, Safari 16.4+).

## What the app expects from the RBA

The app finds tables by their header row, never by position.

| Table | Recognized by | Used for |
|---|---|---|
| Metadata | `Field` / `Value` | `Assessment Name (learner facing)` (the **rubric name**) and `Assessment Points` (the expected total) |
| Rubric | first header is `Criterion Name` | One Canvas row per criterion |
| Approval Record | `Role` / `Name` | Ignored |

The rubric table columns are: `Criterion Name`, `Points`, `Competent`, `Approaching Competency`, `Not Competent`, `Not Submitted`.

### Assessment Points

The criterion points should add up to the **Assessment Points** value in Metadata. Every RBA is normally 100, so **100 is used when the value is blank, missing, or not a number** (not-a-number also warns). If the RBA says something else, such as 80 for a legacy RBA, the app uses that number and notes it. A total that does not match only warns. The expected total is also an editable box in the app.

### How rating points work

The `Points` column is each criterion's maximum. Rating points come from that number:

| Rating | Points | Example: 20 | Example: 15 |
|---|---|---|---|
| Competent | the criterion's points | 20 | 15 |
| Approaching Competency | half, rounded up | 10 | 8 |
| Not Competent | 0 | 0 | 0 |
| Not Submitted | 0 | 0 | 0 |

The scoring rule lives in `src/scoring.ts` and nowhere else. Rating levels are matched by name (case and spacing do not matter). A rating column with any other name cannot be scored and blocks export until it is renamed.

Other rules:

- Rating column headers are plain names (`Competent`, not `Competent (10 points)`). An old-style `(10 points)` suffix is removed and noted.
- Criterion names are kept exactly as written.
- Criteria Description is left blank in the CSV.
- Status is free text and is never checked. The app assumes the file is final.
- Keep the rubric table plain: one row per criterion, no merged cells, no nested tables. Canvas descriptions are plain text, so lists and formatting are dropped.

## Keeping the RBA in sync

The CSV is built from what is on screen, so any edit made in the app is not in the RBA. At export the app lists what to fix:

- **Edits made in the app.** The app remembers the rubric exactly as loaded and compares it with the current rubric. It lists:
  - a rubric name typed or changed
  - an expected total changed
  - a rating column renamed
  - criteria added, removed, renamed or reordered
  - points changed
  - description text filled in, edited or cleared (showing the old and new text)

  It compares end states, so an edit that is typed and then undone does not appear.
- **Parsing fix only.** Each edit has a tick box for "Parsing fix only (the RBA is fine)", for example when you fixed a column that a paste misaligned. Ticked edits no longer count as things to fix in the RBA.
- **Still flagged.** Warnings that are still open at export (a total that does not match, placeholders, unresolved comments and so on) are listed too.
- **Copy these notes** copies the whole list as plain text for a ticket, an email or the RBA's Development Notes.

After Download or Copy, the status line says how many edits are not yet in the RBA. The CSV never contains the notes. The Criteria Enable Range setting is export-only, so it is not reported.

## Checks

**Blocks export** (Canvas could not use the file):

- Rubric name is blank
- No criteria, or no rating columns
- A criterion has no name
- A criterion's points are blank or not a number (0 or higher)
- A rating column has a non-standard name, so its points can't be worked out

**Warns** (you can continue):

- Criterion points do not add up to Assessment Points (100 by default)
- Assessment Points is not a number
- Unresolved comments or tracked changes in the `.docx` (inserted text is kept, deleted text is ignored)
- Template placeholders remain (`#.`, `Title of Sub-Prompt`, "goes here")
- A criterion is worth 0 points, or its points are not a whole number
- Duplicate criterion names, merged cells, nested tables, more than one rubric table, no `Points` column, plain-text paste

**Notes** (informational): empty rating descriptions, Assessment Points that is not 100 or was not found, 1-point criteria (Competent and Approaching Competency are both 1), very long text.

## Canvas CSV format

One row per criterion. Header: `Rubric Name, Criteria Name, Criteria Description, Criteria Enable Range`, then repeating groups of `Rating Name, Rating Description, Rating Points`. `Criteria Enable Range` is lowercase `true` or `false` (default `false`). UTF-8, no byte-order mark, `\n` line endings, trailing newline, standard CSV quoting.

### Not yet verified against a real Canvas import

See [docs/canvas-import-test.md](docs/canvas-import-test.md). Until it is run, treat these as assumptions:

- Canvas accepts two ratings with the same points (Not Competent and Not Submitted are both 0, and a 1-point criterion has Competent and Approaching Competency both at 1).
- Canvas accepts blank Rating Description and blank Criteria Description values.
- Canvas accepts a rubric whose total is not 100.
- Canvas has no length limit that this app's 255-character warnings miss or overstate.

## Development

Requires Node 22.13 or newer (Node 24 recommended). No `npm install` is needed: there are no dependencies.

```
node --test "tests/**/*.test.ts"   # run the tests
node build.mjs                      # build dist/rubric-to-canvas.html (and dist/index.html)
```

Source is TypeScript using only syntax Node can strip natively. `build.mjs` strips the types, inlines the modules into one script and writes a single HTML file. Optional environment variables: `REPO_URL` (footer link), `VERSION`.

- `src/` parsing, checks, scoring, edit tracking (`changes.ts`) and export are pure functions. `app.ts` is the only file that touches the page.
- `tests/` unit tests, a round-trip test, a cross-input equivalence test, and a smoke test that runs the built file against a small fake DOM.
- `samples/` a sample RBA and a sample CSV.
- `src/rba.ts` `extractMetadata()` is the one place that knows where the rubric name and Assessment Points live. Change it there if the template moves them.

A `tsconfig.json` is included for contributors who want `tsc --noEmit`. It has not been run as part of this project's tests.

## Deploying to GitHub Pages

1. In the repository settings, set **Pages > Source** to **GitHub Actions**.
2. Push to `main`. `.github/workflows/pages.yml` runs the tests, builds, and deploys `dist/`.
3. To publish the offline file, push a tag such as `v0.3.0`. `.github/workflows/release.yml` attaches `rubric-to-canvas.html` and the samples to a release.

Pages sites are public. This one is safe to host publicly because the page contains no organization data, only the tool and a clearly labelled sample. Rubrics you load never leave your browser.

## License

To be chosen by the repository owner. Add a `LICENSE` file before publishing.
