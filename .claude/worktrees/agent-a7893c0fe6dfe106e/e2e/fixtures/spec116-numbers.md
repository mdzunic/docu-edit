# spec116-numbers.xlsx / spec116-numbers.numbers

Human-generated SPEC-116 fixtures: `scripts/make-numbers-fixture.applescript`
run on 2026-10-09 with **Numbers 14.4** (build `M14.4-7043.0.93-4`, file
format 14.4.1) on macOS 27, locale `en_US@rg=dezzzz` (English, German region
formats). The `.numbers` file is the document the `.xlsx` was exported from
("Export To ▸ Excel", default options).

## What Numbers wrote into the .xlsx

Real producer shapes no emulated fixture has — the reader must take them as
they are:

- Three worksheets, not two: Numbers prepends **"Export Summary"** (a sheet of
  explanatory text and a sheet/table/worksheet name map; row 3 is 0.05 pt
  high) and names each table's worksheet `"<sheet> - <table>"`:
  "Summary - Figures", "Second - Notes".
- Each table's name is a merged title row (`A1:C1`, `A1:B1`); the header row
  is row 2, so the data in the script's `A1..C5` lands at `A2..C6`. The
  `<dimension>` starts at `A2` although `A1` holds the title.
- **"Summary - Figures"**: `B3` = 1250.5, `B4` = 400, `C3` = the date
  2024-03-15 cached as serial **45366.041666666664** (the script set local
  midnight; Numbers wrote 01:00, the CET offset) with number format
  `dd.mm.yy hh:mm` (numFmtId 59, locale-stamped); `B5` is the formula
  `=B2-B3` exported as a single-cell dynamic-array formula —
  `<c cm="1"><f t="array" ref="B5">B3-B4</f><v>850.5</v>` — backed by
  `xl/metadata.xml` (`XLDAPR`); `B6` is the text "Plain text, with a comma".
- **"Second - Notes"**: `B3` = 42; `B4` is the boolean `true`, exported as the
  **number 1** (no `t="b"`).
- Every row has `r`, `ht` and `customHeight`; columns are declared out to
  16384; panes are frozen at `B3`; gridlines off; every part is deflated.
  No `calcChain.xml`, no `date1904`, and `docProps/app.xml` is an empty
  `<Properties/>` (no `TitlesOfParts`).

## What Numbers wrote into the .numbers (SPEC-116 detection)

A single ZIP file (not a package directory) whose `Index/` holds
`AnnotationAuthorStorage.iwa`, `CalculationEngine.iwa`, `Document.iwa`,
`DocumentMetadata.iwa`, `DocumentStylesheet.iwa`, `Metadata.iwa`,
`ViewState.iwa` and `Index/Tables/{DataList,HeaderStorageBucket,Tile}[-<id>].iwa`
— and **no `Index/Slide…` entry**. The root archive (identifier 1) of
`Index/Document.iwa` has message type 1.

Compared on the same Mac and day (14.4 apps), for the detection rule:

| Producer | `Index/CalculationEngine.iwa` (exact name) | `Index/Slide…` entries | `Document.iwa` root type |
| --- | --- | --- | --- |
| Numbers 14.4 (this file) | yes | none | 1 |
| Keynote 14.4 (`mini-deck.key`) | yes | `Slide.iwa`, `Slide-<id>.iwa` | 1 |
| Pages 14.4 (blank document) | **no** — `Index/CalculationEngine-<id>.iwa` | none | 10000 |

So "exact `Index/CalculationEngine.iwa` and no `Index/Slide` prefix" tells
Numbers from both: a *prefix* match on `Index/CalculationEngine` would call a
Pages document Numbers.

Regenerate both files together (`open -a Numbers` first) and update this note
from the script's report; the cached date and the export's shape are
producer- and locale-dependent, so tests pin values read from the committed
file, not the script's intent.
