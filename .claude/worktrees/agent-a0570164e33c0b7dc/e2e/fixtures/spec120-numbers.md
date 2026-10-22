# spec120-numbers.xlsx

Human-generated SPEC-120 fixture, the Numbers oracle:
`scripts/make-numbers-calc-fixture.applescript` run on 2026-10-09 with
**Numbers 14.4** (build 7043.0.93) on macOS 27.0, launched as
`open -a Numbers --args -AppleLocale en_US` so it parsed and computed the
formulas as en-US (the Mac's own locale is `en_US@rg=dezzzz`, German region
formats). Exported with "Export To ▸ Excel", default options. Nothing in the
file was produced by this repo's engine: every cached value is Numbers'.

## Layout

One worksheet, **"Calc - Oracle"**. With a single table Numbers writes no
"Export Summary" sheet (unlike `spec116-numbers.xlsx`).

- Row 1 is the table name, "Oracle", merged over `A1:E1`; row 2 is the header
  (`a`, `b`, `c`, `formula`, `Numbers`). The script's Numbers row *n* is the
  sheet's row *n*+1. Numbers rewrote every reference in the exported formulas
  to match, but **column D is the text as typed into Numbers**, so D's
  references sit one row above the sheet's (`D8` reads `A7+B7*C7`; the formula
  in `E8` is `A8+B8*C8`). D is for the report only. Read the formula from E.
- Rows 3–7 are a data block: `A` "apple", "banana", "cherry", "date", "Banana";
  `B` 10–50; `C` 1, 2, 3, empty, "x". The aggregate, criteria and lookup rows
  reference it absolutely (`$B$3:$B$7` in the file).
- Rows 3–149: one formula per row in `E` (147 formulas), each with its own
  inputs in `A`–`C` where it has any. Text inputs such as "5", "3" and "1" are
  shared strings, not numbers. No input is a boolean, because Numbers exports a
  checkbox as the number 1.

## What every formula cell looks like

**Every one of the 147 formula cells is a single-cell dynamic-array formula**:
`<c r="E8" s="8" cm="1"><f t="array" ref="E8">A8+B8*C8</f><v>7</v></c>`,
backed by `xl/metadata.xml` (`XLDAPR`, `fDynamic="1"`). This includes plain
arithmetic. SPEC-119's reader marks each one `dynamicArray: true`. Under
SPEC-120's "Array mode", a dynamic-array formula whose range is its own cell
computes, and is uncomputable only if its result would spill. So every row
here computes, and the oracle needs no special case.

Cached values by type:

- number: `<v>` with up to 16 significant digits (`0.333333333333333`,
  `2.71828182845905`, `1462.520833333333`)
- text: `t="str"`
- boolean: `t="b"` with 1/0 (a formula's boolean *result* stays a boolean;
  only a checkbox *input* becomes a number)
- date: a serial number with numFmt 14, or numFmt 59 (`h:mm" "AM/PM`, with a
  U+202F narrow space) for TIME
- **error: no `<v>` and no `t="e"`**. Numbers writes no error codes, so a
  formula with no cached value is one Numbers evaluated to an error. Fourteen
  cells: `E16 E17 E27 E29 E30 E31 E35 E39 E40 E51 E79 E81 E83 E115`.

The workbook has no `calcPr`, no `calcChain.xml` and no `date1904`.
`docProps/app.xml` is an empty `<Properties/>`. Every part is deflated.

## Where Numbers differs from Excel

Every row was checked against SPEC-120's conformance table and Excel's
documented behaviour. Every value not listed here agrees with them,
including errors: Numbers errs exactly where Excel returns an error code,
apart from the three error rows below. These nine rows are the ones
`numbersOracle.spec.ts` lists, with these reasons:

| Cell | Formula | Numbers | Excel / SPEC-120 | Why |
| --- | --- | --- | --- | --- |
| `E35` | `"z"<TRUE` | error | TRUE | Numbers does not compare text with a boolean |
| `E51` | `CEILING(-2.5,2)` | error | −2 | Numbers' CEILING needs both arguments to have the same sign (Excel 2007's rule) |
| `E79` | `SUM("3",TRUE)` | error | 4 | Numbers' SUM rejects a boolean argument (`SUM(TRUE,1)` errs too; `SUM("3")` is 3) |
| `E110` | `PROPER("it's")` | "It's" | "It'S" (pinned) | Numbers does not capitalise after an apostrophe |
| `E121` | `LEN("😀")` | 1 | 2 (pinned) | Numbers counts characters, not UTF-16 code units |
| `E124` | `DATE(108,1,2)` | text `"1/2/08"` | 2008-01-02 (39449) | Numbers takes year 108 literally (no +1900). No serial holds that date, so the export writes the displayed text |
| `E125` | `DATE(1900,2,29)` | 61 | 60 | Numbers' calendar has no 29 Feb 1900 (it rolls over to 1 Mar). Excel keeps the 1900 leap-year bug |
| `E129` | `TIME(12,30,0)` | 1462.520833333333 | 0.520833333333333 | Numbers anchors a time of day to 1 Jan 1904 (serial 1462) |
| `E135` | `WEEKDAY(DATE(1900,1,1))` | 2 | 1 | Same calendar: 1 Jan 1900 was a Monday, while Excel's leap-bug arithmetic makes it a Sunday |

## What Numbers 14.4 lacks

IFNA, XOR, DAYS, ISNA, ISERR, ISLOGICAL and TEXT. Numbers accepts these as
formulas but cannot compute them, even with trivial arguments. The script
leaves them out. In the export the first three disappear entirely (an empty
styled cell). The last four keep their formula with no cached value, which
looks exactly like an error. AppleScript reads both cases as `missing value`,
so the script checks the exported file: every case must arrive as an `<f>`.

## Regenerating

Quit Numbers. Run `open -a Numbers --args -AppleLocale en_US`, then
`osascript scripts/make-numbers-calc-fixture.applescript` from the repo root.
The script reports `0 FAIL` and lists the cells with no cached value. If the
Numbers version changes, update this note, especially the dynamic-array
shape, the error encoding and the table above. Tests pin values read from the
committed file, not the script's intent.
