# Canvas import test (manual, required before relying on the tool)

Do this once in a Canvas **test course** (not a live course) and record the results here.

## Steps

1. Load `samples/sample-rba.docx` in the app. Export the CSV (`Sample Reflective Journal.csv`).
2. In the test course, import the CSV as a rubric.
3. Open the new rubric and compare it to the sample RBA.

## Check and record

| Check | Expected | Result |
|---|---|---|
| Rubric appears with the exported name | `Sample Reflective Journal` | |
| Number of criteria | 10 | |
| Criterion names | Match the RBA exactly | |
| Each criterion has four ratings, in order | Competent, Approaching Competency, Not Competent, Not Submitted | |
| Points for A. Sample Prompt One (15) | 15, 8, 0, 0 | |
| Points for C. Sample Prompt Three (10) | 10, 5, 0, 0 | |
| Points for F. Professional Comportment (5) | 5, 3, 0, 0 | |
| Two ratings with 0 points (Not Competent, Not Submitted) | Accepted | |
| Empty Rating Description cells | Accepted. Edit a copy of the CSV to blank a whole row of descriptions and try again | |
| Blank Criteria Description | Accepted | |
| Multi-line descriptions | Line breaks kept | |
| Curly quotes and non-ASCII text | Show correctly | |
| Total points | 100 | |
| A 1-point criterion (Competent 1, Approaching Competency 1) | Accepted. Edit a copy of the CSV to try it | |
| A total other than 100 (legacy RBA) | Accepted. Load an RBA whose Assessment Points is not 100 | |

## If something fails

- Note which check failed and the error Canvas showed.
- Tell the tool owner. If Canvas rejects equal points or blank descriptions, the matching check in `src/validate.ts` should change from a note to a block.
