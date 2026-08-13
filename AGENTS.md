# AGENTS.md

Conventions for any AI agent working in this repository.

## Git

- **Never `git push` unless the human explicitly asks for it in that message.** Committing locally is fine and expected; publishing is always their decision.
- **Never `git add -A` / `git add .` blindly.** Stage named paths. If a bulk add seems unavoidable, run `git status` first and report what would be staged.
- **Never commit real scan exports.** Files like `test11.txt` and any `*_result.xlsx` contain machine IDs, timestamps and EN numbers belonging to real people. They are gitignored — do not add exceptions, do not paste their contents into code, docs, issues, or commit messages. `sample_data.txt` (synthetic, generated) is the only data file that belongs here.
- **Never commit the `.exe`.** It is a 30 MB PyInstaller artifact, rebuildable from source.
- Commit author identity comes from global git config (`NW <olay097056@gmail.com>`). Do not override it per-repo or per-commit.
- Do not rewrite history that already exists on `origin/main`.

## Code

- The record format is fixed-width and exact: 20 characters as unit(2) year(4) month(2) day(2) time(4) EN(6). Changing the layout changes the contract with the scanner hardware — do not "improve" it.
- **The two validation layers are separate on purpose.** Hard rules catch structurally impossible values (red); per-unit frequency catches values that are legal but abnormal *for that machine* (yellow). Do not merge them into one score or one colour.
- **Frequency analysis runs on year/month/day only.** Time and EN have high cardinality — applying it there floods the output with false positives and buries the real findings. Do not extend `FREQ_FIELDS` without a concrete reason and a test showing the false-positive rate.
- Control characters are stripped before parsing, because openpyxl cannot write them. Keep that step ahead of any new parsing path.

## Docs

- `README.md` and `README.th.md` are kept in sync. Updating one means updating the other.
