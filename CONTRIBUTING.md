# Contributing

Keep it small. If a change does not help turn an RBA into a Canvas rubric CSV, it probably does not belong here.

**Rules that never bend**

- No network requests, analytics, storage or third-party code. The tests check for this.
- No dependencies unless there is a strong reason.
- Parsing, checks, scoring, edit tracking and export stay in pure functions, separate from the page code.
- All source files are inlined into one script by `build.mjs`, so **top-level names must be unique across files**.

**Before you open a pull request**

1. `node --test "tests/**/*.test.ts"` passes.
2. Add a test for anything you change. Build `.docx` fixtures in code with `tests/helpers/docxbuilder.ts` where you can.
3. If you change the RBA template contract or the scoring rule, update the README tables, `src/rba.ts` and `src/scoring.ts` together.
4. Check by hand in a real browser: load the sample, load a pasted Word table, edit a cell, export, read the hand-off list, and open the file from disk.

**Style:** erasable TypeScript only (no enums, namespaces or parameter properties), named exports, one module per concern, plain-language messages that say what happened and how to fix it.
