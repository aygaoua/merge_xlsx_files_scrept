# Excel Merge Script (TypeScript + ExcelJS)

Merge salaries from a user Excel file into a base Excel file while preserving the original theme/formatting and enforcing a specific column layout.

## Features

- Preserves workbook theme and formatting using a template or the base file
- Idempotent layout: inserts SALAIRE NET next to FONCTION and keeps R/S/T stable
- Robust header detection (accents/spacing/punctuation normalized)
- CIN matching is case-insensitive and trimmed
- Clears duplicate "SALAIRE NET" columns (keeps only Q)
- Does not overwrite inputs; writes to `output_merged.xlsx`

## File layout (columns)

- P = FONCTION
- Q = SALAIRE NET (inserted next to FONCTION)
- R = CATEGORIE
- S = DATE EMB
- T = DATE SORTIE

If any of these columns exist elsewhere, the script copies them to the target columns and updates headers accordingly.

## Requirements

- Node.js 16+ recommended

## Getting started

1) Place your Excel files in the project root:
	 - Base: `fulldbbase.xlsx`
	 - User: `userdb.xlsx`
	 - Optional template (for theme): `emptydb.xlsx`

2) Install dependencies (first time only):

```sh
npm install
```

3) Build and run:

```sh
npm run build
npm run start
```

The result is written to `output_merged.xlsx` in the project root.

## How it works

- Copies `emptydb.xlsx` to `output_merged.xlsx` if present; otherwise copies the base file
- Copies headers and data from base into the output
- Detects the `FONCTION` column; inserts a new column immediately to its right for `SALAIRE NET`
- Ensures headers P..T are exactly: `FONCTION, SALAIRE NET, CATEGORIE, DATE EMB, DATE SORTIE`
- Builds a CIN → salaire map from the user file and fills `SALAIRE NET` (Q) in the output
- Clears any duplicate `SALAIRE NET` columns outside Q (e.g., one that already existed at D)

## Customizing file names

Edit the constants at the top of `src/script.ts`:

```ts
const baseFile = 'fulldbbase.xlsx';
const userFile = 'userdb.xlsx';
const templateFile = 'emptydb.xlsx';
const outputFile = 'output_merged.xlsx';
```

## Troubleshooting

- No salary filled for some rows:
	- Check CIN values exist in both files and match after trimming (case-insensitive)
	- Ensure the user file has a salary column detected as `SALAIRE` or `NET` in the header
- SALAIRE NET appears in another column (e.g., D):
	- The script will clear duplicates automatically and keep only Q
- Headers not detected:
	- Header normalization removes accents and non-alphanumerics; ensure the first row contains headers

## Development

- Tech stack: TypeScript + ExcelJS
- Scripts:

```sh
npm run build   # compile TypeScript to dist/
npm run start   # run compiled script
npm run dev     # (optional) run ts-node if configured
```

- Main entry: `src/script.ts`
- Output: `output_merged.xlsx`

## Roadmap / ideas

- CLI flags to pass custom file paths
- More flexible CIN normalization (strip non-alphanumerics, handle leading zeros)
- Unit tests for header detection and column layout
