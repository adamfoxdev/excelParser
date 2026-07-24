# Excel Template Extractor

A browser app for pulling structured data out of spreadsheets. Open a workbook, mark the
ranges that matter, save that as a reusable **template**, then apply it to future files.

Files are parsed entirely in the browser — nothing is uploaded, and templates are stored
locally in IndexedDB.

## Running it

```bash
npm install
npm run dev
```

| Script            | What it does                                     |
| ----------------- | ------------------------------------------------ |
| `npm run dev`     | Dev server on http://localhost:5173              |
| `npm run build`   | Production build into `dist/`                    |
| `npm test`        | Unit tests for the extraction engine             |
| `npm run lint`    | oxlint                                           |
| `npm run samples` | Regenerates the sample workbooks in `samples/`   |

## How it works

1. **Open workbooks** — `.xlsx`, `.xlsm`, `.xls`, `.csv`, `.tsv`. Drop one or many at once.
2. **Select a range** — drag across cells, or click and shift-click to extend.
3. **Add field from selection** — names itself from the label above the range where it can.
4. **Save** the template. It persists across reloads and can be exported as JSON.
5. **Open more files, load the template** — the Results tab shows what it extracted from
   every one of them, exportable as JSON or CSV.

## Batch processing

Drop several files and the template runs against all of them. The file bar above the grid
switches which file the grid and the by-field view show; **Flat table** shows the whole batch
at once, one row per file.

- One unreadable file does not sink the batch — it gets an *unreadable* chip with the parser
  error on hover, and the rest still load.
- Each file chip shows a red count of fields that failed to resolve in that file.
- Exports cover every loaded file: the flat CSV gets one row each, and the JSON export an
  entry per file.
- Editing a field re-extracts the active file immediately; the full batch recomputes at lower
  priority (`useDeferredValue`), so typing an anchor stays responsive with many files open. A
  brief *updating…* tag appears while the batch catches up.

Batch is where anchored fields earn their keep. Running a template built on `invoice-a.xlsx`
across all four samples, with the total field pinned to a **fixed** cell:

| File | Invoice total (fixed) | Invoice total (anchored) |
| ---- | --------------------- | ------------------------ |
| invoice-a.xlsx | *(empty)* | `231.66` |
| invoice-b-shifted.xlsx | `231.66` | `231.66` |
| invoice-c.xlsx | *(empty)* | `118.8` |
| invoice-d.xlsx | `128.8` ← wrong row | `402.15` |

The fixed selector does not error on `invoice-d` — it returns a real number from the wrong
row. That silence is the argument for anchoring anything you intend to run over a batch.

### Fixed vs. anchored fields

Each field resolves its data one of two ways, switchable in the field editor.

**Fixed** stores a literal location (`Invoice!A9:E13`). Predictable, and right when every
file is generated identically. It reads whatever now sits at those coordinates, so inserting
rows above the table silently shifts it onto the wrong data.

**Anchored** finds a label first, then reads a region relative to it:

| Setting              | Meaning                                                         |
| -------------------- | --------------------------------------------------------------- |
| Anchor text          | Label to search for (exact / contains / regex)                   |
| Sheet                | A specific sheet, or *Any sheet* to search all of them in order  |
| Occurrence           | Which match to use when the label repeats                        |
| Offset rows / cols   | Where the region starts relative to the anchor cell              |
| Height / width       | Region size; blank means **auto**                                |

Auto height grows downward until a fully blank row ends the region; auto width grows right
until a blank cell in the region's first row does. So a field anchored to `"SKU"` with auto
sizing captures the whole line-items table however far down the sheet it has moved.

### Output shapes

- **Table** — rows and columns, optionally taking the first row as headers.
- **List** — the region flattened into one column, blanks dropped.
- **Single value** — the first non-blank cell. Use with a `+1` column offset to read the
  number sitting next to a label.

"First row is headers" only applies to **Table**; switching a field to List or Single value
clears it, since a header row would otherwise eat the first item of a list.

### Flat table

The Results tab toggles between **By field** — each field of the active file rendered as its
own block — and **Flat table**, which collapses the batch into one row per file: one column
per field, with any field spanning multiple rows JSON-encoded into its single cell.

| File | Line Items | Invoice total | SKUs |
| ---- | ---------- | ------------- | ---- |
| invoice-a.xlsx | `[{"SKU":"A-1001","Qty":12}, ...]` | `231.66` | `["A-1001","A-1002"]` |
| invoice-d.xlsx | `[{"SKU":"B-2010","Qty":60}, ...]` | `402.15` | `["B-2010","C-3300"]` |

- **Table** fields become a JSON array of objects, or arrays if the field has no headers.
- **List** fields become a JSON array.
- **Single value** fields stay bare scalars — no JSON.
- Failed fields become an empty cell.

The shape of a column is fixed by the field's output type, never by how much data a given
file happened to hold: a one-row table is still `[{...}]` and an empty one is still `[]`. A
column that were a bare number in one file and an array in the next would be unusable
downstream. If you want a bare value in a column, set that field's output to **Single value**.

In flat mode the CSV export switches to the same shape (embedded JSON is quote-escaped, so it
survives a CSV round trip); in by-field mode it exports the active file as labelled blocks.
The JSON export always nests under a `files` array — one entry or twenty, a consumer never has
to branch on batch size. The flat/by-field toggle is saved with the template.

## Try it

`samples/` holds four invoices sharing a layout at different vertical offsets, with different
row counts and values:

| File | Line items at | Rows | Total |
| ---- | ------------- | ---- | ----- |
| `invoice-a.xlsx` | `A9:E13` | 4 | 231.66 |
| `invoice-b-shifted.xlsx` | `A13:E17` | 4 | 231.66 |
| `invoice-c.xlsx` | `A11:E13` | 2 | 118.8 |
| `invoice-d.xlsx` | `A18:E22` | 4 | 402.15 |

Build a template against the first with an anchored field on `"SKU"`, then drop all four in
at once and switch to **Flat table**. Every row resolves. Point the same field at a fixed
range instead and only the file you built it on stays correct — pinned down in
`src/lib/extract.test.ts` and `src/lib/flatten.test.ts`.

## Layout

```
src/
  lib/
    types.ts       Template / field / selector shapes
    range.ts       A1 notation <-> row-col indices
    workbook.ts    Files -> dense cell matrices (SheetJS, loaded on demand)
    extract.ts     Selector resolution and extraction — the core
    flatten.ts     Extractions -> one flat row per file, multi-row fields as JSON
    storage.ts     TemplateStore interface + IndexedDB implementation
    download.ts    JSON / CSV export
  components/
    SheetGrid.tsx  Virtualised grid with drag selection
    FileBar.tsx    Batch file chips: switch, remove, per-file error counts
    FieldEditor.tsx
    ResultsPanel.tsx
```

Templates go through the `TemplateStore` interface in `src/lib/storage.ts`, so swapping the
IndexedDB implementation for a server-backed one to share templates between users means
writing one new implementation and leaving the UI alone.

## Known issue: SheetJS version

The newest `xlsx` on the npm registry is `0.18.5`, which carries two open advisories
(prototype pollution, ReDoS). SheetJS publishes patched builds only from their own CDN.
To move to a patched release:

```bash
npm install https://cdn.sheetjs.com/xlsx-0.20.3/xlsx-0.20.3.tgz
```

The API used here is unchanged between the two, so nothing else needs to change. Left as
your call since it installs from outside the npm registry.
