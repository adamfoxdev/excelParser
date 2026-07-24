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

1. **Open a workbook** — `.xlsx`, `.xlsm`, `.xls`, `.csv`, `.tsv`. Each sheet gets a tab.
2. **Select a range** — drag across cells, or click and shift-click to extend.
3. **Add field from selection** — names itself from the label above the range where it can.
4. **Save** the template. It persists across reloads and can be exported as JSON.
5. **Open a different file, load the template** — the Results tab shows what it extracted,
   exportable as JSON or CSV.

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

## Try it

`samples/` holds two invoices with identical content at different vertical offsets:

- `invoice-a.xlsx` — line items at `A9:E13`
- `invoice-b-shifted.xlsx` — same table, pushed down to `A13:E17`

Build a template against the first with an anchored field on `"SKU"`, then load the second:
it still resolves. A fixed field on `A9:E13` does not — that difference is the point of the
feature, and it's pinned down in `src/lib/extract.test.ts`.

## Layout

```
src/
  lib/
    types.ts       Template / field / selector shapes
    range.ts       A1 notation <-> row-col indices
    workbook.ts    File -> dense cell matrices (SheetJS, loaded on demand)
    extract.ts     Selector resolution and extraction — the core
    storage.ts     TemplateStore interface + IndexedDB implementation
    download.ts    JSON / CSV export
  components/
    SheetGrid.tsx  Virtualised grid with drag selection
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
