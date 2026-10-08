---
id: SPEC-118
title: Number formats and dates
status: draft
area: formatting
complexity: complex
story_points: 13
depends_on: [SPEC-066, SPEC-116, SPEC-117]
---
## Why
After SPEC-117 a workbook looks right until you read it. The invoice date in
`B4` says 46303. The margin column says 0.125. The price list's accounting
column says -1234.5 where the author saw "(1,234.50)" in red, the timesheet's
total says 1.5 where it meant 36 hours, and a German user sees full stops
where they expect commas. Every one of those cells carries its display rule
in the file — `numFmtId`/`<numFmts>` in `xl/styles.xml`, `number:*` data
styles in the ODS — and SPEC-116 says it out loud: "a date shows as its
serial until SPEC-118". Dates are serials by design (SPEC-116's fourth
invariant); the number format is what makes 46303 a Thursday.

There is no format engine in this repo and none may be added
(`package.json` is frozen). `core/dateTimeFormats.ts` renders nine closed
date patterns for the word processor's fields; it is not a format-code
language. This spec writes one: a parser and renderer for Excel's
format-code grammar (the lingua franca — ODS number styles are mapped onto
it both ways), the built-in formats Excel refers to by id, a small locale
layer, and General's width-fitting.

**It is the normative source for the format-code grammar and its rendering
(`parseFormatCode`, `formatValue`), the built-in formats 0–49, the mapping
between ODS `number:*` styles and codes (`odsNumberStyleToCode`,
`odsNumberStyleXml`), `SpreadsheetLocale` and its six locales, General's
width-fitting and the "########" overflow, and `editText` — what the formula
bar and the cell editor show for a value.** Two rules are pinned here for the
whole queue: the locale is a **per-user preference** (SPEC-066), never a
property of a file; and **formulas are always entered and shown in the
invariant en-US syntax** — the locale changes how values look and (SPEC-122)
how typed values are read, never formula text and never what is saved.

## Acceptance criteria
- [ ] `spreadsheet/numberFormat/parse.ts` exports `parseFormatCode(code): ParsedFormat` per "The grammar": up to four `;`-separated sections, conditions, colours, `[$…]` currency/locale tags, elapsed `[h]`/`[m]`/`[s]`, digit placeholders `0 # ?`, `.`, `,` (grouping and trailing-comma scaling), `%`, `E+ E- e+ e-`, fractions, `@`, quoted literals, `\x`, `_x`, `*x`, the `General` token and every date/time token; results are cached (`FORMAT_CACHE_LIMIT = 1024`, cleared on overflow); structural errors throw `FormatCodeError` with an offset, and a code using an unsupported calendar (`B1`/`B2` prefixes, `e`/`g` era tokens, `[DBNum*]`/`[NatNum*]`, a calendar byte in an LCID) parses as `{ unsupported: true }`
- [ ] `spreadsheet/numberFormat/render.ts` exports `formatValue(value, code, ctx): FormattedValue` per "Rendering numbers", "Rendering dates and times" and "Text, booleans and errors" — never throws (an invalid or unsupported code renders as General), section choice and the sign rule per "Sections", 15-significant-digit decimal rounding half away from zero (`2.675` with `0.00` → "2.68", `1.005` → "1.01", `-0.001` → "-0.00"), the locale's decimal and group characters substituted for the code's invariant `.` and `,`
- [ ] Dates and times render per "Rendering dates and times": the minute rule for `m`/`mm`, rounding to the smallest shown unit with carry (`0.999999` with `h:mm` → "0:00"), elapsed `[h]:mm:ss` (`1.5` → "36:00:00"), `AM/PM` and `A/P` keeping their case from the code, month and day names from `Intl.DateTimeFormat` for the locale (or the section's LCID), the genitive month form when the section also shows a day, Excel's fictitious 1900-02-29 (serial 60 with `d-mmm-yyyy` → "29-Feb-1900"), serial 0 as "1900-01-00", and negative or post-9999 serials as "########" for an xlsx-born workbook while an ods-born one renders negative serials as proleptic Gregorian dates
- [ ] `formatGeneral(value, ctx)` subsumes SPEC-116's `generalText` per "General": at full width its output is `generalText`'s for every value in `general.spec.ts` (which passes unchanged — `generalText` now delegates); with `ctx.widthPx` it shrinks decimals, then switches to scientific, then fills with `#` (64 px, Calibri 11: `123456789012` → "1.23E+11", `1234.56789` → "1234.568", `123456789` → "1.23E+08"; 26 px: `12345` → "###"); and it uses the locale's decimal separator
- [ ] A non-General number, date or time whose text does not fit `ctx.widthPx` renders as `#` repeated to the width (`overflow: "hash"`); text never shrinks or hashes; with no `widthPx` nothing is fitted
- [ ] `spreadsheet/numberFormat/builtin.ts` exports `BUILTIN_NUMBER_FORMATS: Record<number, string>` (ids 0–49 per "Built-in formats", en-US spellings), `LOCALE_DEPENDENT_BUILTINS` (5–8, 14, 22, 41–44), `builtinFormatCode(id, locale)` and `formatCodeOf(style, locale)` — a style whose `builtinFormatId` is locale-dependent renders with that locale's version (id 14: "10/8/2026" in en-US, "08.10.2026" in de-DE); ids 23–36 and 50–81 render as General
- [ ] `spreadsheet/numberFormat/locale.ts` exports `SpreadsheetLocaleTag`, `SPREADSHEET_LOCALE_TAGS`, `SpreadsheetLocale`, `SPREADSHEET_LOCALES`, `DEFAULT_SPREADSHEET_LOCALE = "en-US"` and `localeNames(tag)` with every field of "Locales" — including what SPEC-122's input parser needs (decimal, group, date order and separator, short/long date and time codes, AM/PM strings, currency symbol, position and code) — and month/day names built once per tag with `Intl.DateTimeFormat(tag, { timeZone: "UTC" }).formatToParts`; unit tests pin the en-US and de-DE name tables and the cs-CZ genitive ("8. října 2026") as fixed strings
- [ ] The xlsx reader fills `CellStyle.numberFormat` from `cellXfs/xf@numFmtId` per "Reading xlsx number formats": an explicit `<numFmts>/<numFmt>` wins for any id (Excel writes 5–8 and 41–44 explicitly), otherwise the built-in; a locale-dependent built-in also sets `builtinFormatId`; General leaves both absent
- [ ] The ODS reader fills `CellStyle.numberFormat` from the flattened style's `style:data-style-name` by `odsNumberStyleToCode` per "ODS number styles": `number:number-style`, `scientific-number`, `fraction`, `percentage-style`, `currency-style`, `date-style` (incl. textual and long variants, day-of-week, quarter, week-of-year), `time-style` (`number:truncate-on-overflow="false"` → elapsed), `boolean-style`, `text-style` with `number:text-content`, `style:map` conditions → sections (LibreOffice's `value()>=0` positive map → the conventional unconditioned two sections), `fo:color` → a named colour; a `number:era` style reads as General
- [ ] `odsNumberStyleXml(code, name): string` in `numberFormat/odsNumberStyles.ts` is the reverse mapping SPEC-122's ODS writer calls: it returns the `number:*-style` element(s) for a code — sub-styles `${name}P0`, `${name}P1`… for multi-section codes plus the main `${name}` with its `style:map`s — and `odsNumberStyleToCode(parse(odsNumberStyleXml(c, "N1")))` returns `c` for every code in the round-trip corpus ("Tests")
- [ ] Values the source marks as dates or percentages but leaves unformatted get an implied format per "Implied formats" (an ODS `office:value-type="date"` cell without a data style, an xlsx `t="d"` cell with a General style), so a date never arrives as a bare serial because its producer skipped the style
- [ ] The grid renders every value through `formatValue(value, formatCodeOf(style, locale), ctx)` with `widthPx` = the cell's (or merge's) width and `digitWidthPx` = SPEC-117's `maxDigitWidthPx` of the cell font: `FormattedValue.color` overrides the font colour, `alignHint` drives `general` alignment, `fill` expands as a flexible run of its character, and hashed cells never overflow into neighbours
- [ ] `editText(cell, ctx)` (render.ts; `ctx.style` = the cell's effective `CellStyle`, absent = Normal) is what the read-only formula bar shows, per "editText": dates as the locale's short date (+ long time when the serial has a time part), time-only values as the long time, percent-formatted numbers ×100 with `%`, other numbers in plain notation up to 15 significant digits with the locale decimal separator and no grouping, `TRUE`/`FALSE`, error codes, and text (with a leading `'` when `quotePrefix`); `formula-bar-content` shows it (en-US: "10/8/2026 1:45:00 PM"; de-DE: "08.10.2026 13:45:00")
- [ ] The status-bar Sum and Average (SPEC-117) render with the format of the first numeric cell of the selection in row-major order; Count stays a plain integer
- [ ] The locale is a preference per "The preference": `Preferences.spreadsheetLocale` (default `"en-US"`, validated against `SPREADSHEET_LOCALE_TAGS`) in `src/lib/preferences.ts`, a "Spreadsheets" section in `PreferencesDialog.vue` with "Number and date format" (`preferences-spreadsheet-locale`, each option labelled with a live sample), and a `sheet-locale` status-bar control with one `sheet-locale-<tag>` radio item per locale; changing it re-renders every open sheet immediately, marks nothing dirty and changes no stored or downloaded byte
- [ ] Booleans and error codes display in English ("TRUE", "#DIV/0!") in every locale, and formula text (SPEC-119) is never localised — asserted in `render.spec.ts` for de-DE and fr-FR
- [ ] `buildPreview(model, locale)` (SPEC-116's library card) formats with `formatValue` at full width in the locale it is given; the app layer passes the current preference, and a later locale change leaves stored previews as they are until the workbook's next preview rebuild
- [ ] `src/lib/spreadsheet/numberFormat/` imports only spreadsheet siblings and core (`boundary.spec.ts`), never `src/lib/preferences.ts` — the view resolves the preference and passes the locale; `src/views/spreadsheet/useSpreadsheetLocale.ts` is the one place that does
- [ ] Both readers' extensions bump `SPREADSHEET_READER_VERSION` (to the next integer) and append a `ReaderUpgrade` per "Reader version": a workbook stored by an earlier reader reopens with its formats
- [ ] Fixtures per "Fixtures": `e2e/fixtures/spec118-formats.xlsx` (Excel-shaped), `spec118-generated.xlsx` (openpyxl-shaped) and `spec118-formats.ods` (LibreOffice-shaped), built by `spreadsheet/fixtures/build.ts`, and `roundTrip.spec.ts` keeps all three byte-identical
- [ ] Unit tests: `numberFormat/parse.spec.ts`, `render.spec.ts` (every example in this Reference as a table-driven test), `builtin.spec.ts`, `locale.spec.ts`, `odsNumberStyles.spec.ts` (both directions and the corpus round trip), `general.spec.ts` (width cases added), `xlsx/styles.spec.ts` and `ods/styles.spec.ts` (number-format links), `open.spec.ts` (upgrade), `src/lib/preferences.spec.ts` (the new field's validation), `FormulaBar.spec.ts` and `SpreadsheetStatusBar.spec.ts` (formatted output)
- [ ] e2e (`e2e/SPEC-118.spec.ts`): open `spec118-formats.xlsx` → `A2` (id 14) reads "10/8/2026", `A3` (id 22) "10/8/2026 13:45", `A4` (`[$-F800]`) "Thursday, October 8, 2026", `B2` "(1,234.50)" in red, `C2` "12.50%", `D2` "36:00:00" → the narrow column F shows "1.23E+11" and the narrow date in G shows only `#` → click `A4` → `formula-bar-content` reads "10/8/2026 1:45:00 PM" → choose `sheet-locale-de-DE` → `A2` reads "08.10.2026", `A4` "Donnerstag, 8. Oktober 2026", `B5` "1.234,50", the formula bar "08.10.2026 13:45:00" → reload → still de-DE → download → every part `equalBytes` the fixture's → open `spec118-formats.ods` → the LibreOffice styles render (`-1,00` in red, "3 14/99", "Q4 2026", and the unstyled 1899-12-25 date cell as "1899-12-25")

## Out of scope
Typing values and the automatic format a typed value earns ("1/2" → a date,
"12%" → `0%`) — SPEC-122's `numberFormat/input.ts`, which consumes the locale
fields pinned here. Choosing or writing formats: the Format Cells number tab,
the custom-format editor and the toolbar's format buttons (SPEC-123, which
shows codes in the invariant syntax this spec stores). Writing number formats
into files (SPEC-122 calls `odsNumberStyleXml` and the built-in table). The
`TEXT()` function (SPEC-120/127 call `formatValue`). Conditional-format
colours (SPEC-129). Non-Gregorian calendars, eras, Buddhist years, `[DBNum]`/
`[NatNum]` numerals and the CJK built-ins 23–36/50–81 — they render as
General, tested. Localised booleans, error codes, function names or
argument separators (backlog, with locale-specific formula syntax). Locales
beyond the six. Detecting the locale from the browser (the default is
en-US, deliberately). Exact glyph-width padding for `_x` and `?`.
Right-to-left numerals.

## Reference

### Dependencies restated
- `specs/116-workbooks-in-the-library.md` — dates are serials (no date type);
  `serial.ts`: `serialToParts(serial, system)` → `{ year, month, day, hour,
  minute, second, ms }`, 1900 system with the fictitious serial 60 =
  1900-02-29, serials ≥ 61 = days since 1899-12-30, serials before 1 counting
  backwards (1899-12-31 = 0, 1899-12-30 = −1); 1904 system with 0 =
  1904-01-01. `general.ts`: `generalText(value)` — `null` → "", strings as
  is, "TRUE"/"FALSE", error codes, numbers by the 11-character rule (fixed
  notation for `1e-9 ≤ |n| < 1e11` rounded to ≤ 15 significant digits,
  trailing zeros dropped, decimals cut to fit 11 characters; otherwise
  scientific with upper-case `E` and a two-digit signed exponent).
  `preview.ts`: `buildPreview` — `generalText` of the active sheet's top-left
  12 × 6. Reader versions (`open.ts`): bump `SPREADSHEET_READER_VERSION`,
  append a `ReaderUpgrade { from; upgradeDirtySheet(model, sheet, parts) }`;
  older clean records are re-read from `opaque`.
- `specs/117-sheet-grid-styles-sizes-and-selection.md` — `model.styles:
  CellStyle[]`, resolved by `effectiveStyleId(sheet, row, col)` (cell → row →
  column → 0) and `cellStyle(model, id)`; `CellStyle.numberFormat?: string`
  (absent = General) and `builtinFormatId?: number` are declared there and
  populated here; `quotePrefix?: true`; `sourceXf`/`sourceStyleName`. The ODS
  style flattening (default style → parents → automatic style) and the xlsx
  `cellXfs` → `styles` one-to-one mapping. `geometry.ts`: `maxDigitWidthPx
  (family, sizePt)`, `CELL_PADDING_PX = 2` (+1 px gridline: 5 px per cell),
  and the text-width estimate (per-character weight × MDW; 0.5 for `.` `,`
  `:` and space, 1 otherwise). `INDEXED_COLORS` (`styles.ts`). The read-only
  `FormulaBar.vue` (`formula-bar-content`), `SpreadsheetStatusBar.vue`
  (`sheet-status-sum`/`-average`/`-count`) and `cellRender.ts`.
- `specs/066-preferences.md` — `src/lib/preferences.ts`: `Preferences`,
  `PREFERENCE_DEFAULTS`, per-field validators (a bad field falls back to its
  default alone), `loadPreferences()`, `savePreference(key, value)`,
  `onPreferencesChange(listener)` (fires across tabs); the reactive wrapper
  `src/composables/usePreferences.ts`; `PreferencesDialog.vue` built on
  `BaseDialog.vue` with headed sections and `preferences-*` testids. The rule
  it established: a setting about how this person wants the application to
  behave is a preference; a setting another application opening the file
  would need belongs to the file. Number display locale is the former.

### API
```ts
// numberFormat/render.ts
export interface FormatContext {
  locale: SpreadsheetLocale;
  dateSystem: "1900" | "1904";
  widthPx?: number;            // the cell's width (column or merge); absent = full width, no fitting
  digitWidthPx?: number;       // MDW of the cell font; default 7
  dateQuirks?: "excel" | "proleptic";   // default "excel"; the view passes "proleptic" for format === "ods"
  palette?: readonly string[]; // [ColorN] palette; default INDEXED_COLORS (model.indexedColors when set)
}
export interface FormattedValue {
  text: string;
  color?: string;                          // "#rrggbb" from [Red] / [ColorN]
  alignHint?: "left" | "right" | "center"; // for horizontal "general"
  overflow?: "hash";                       // text is the # fill
  fill?: { char: string; index: number };  // *x: expand `char` at `index` to fill the width
}
export function formatValue(value: CellValue, code: string | undefined, ctx: FormatContext): FormattedValue;
export function formatGeneral(value: CellValue, ctx: FormatContext): FormattedValue;
export interface EditTextContext extends FormatContext { style?: CellStyle }   // the cell's effective style
export function editText(cell: Cell, ctx: EditTextContext): string;
export function plainText(fv: FormattedValue): string;   // text with one fill char inserted (CSV, clipboard, preview)

// numberFormat/parse.ts
export function parseFormatCode(code: string): ParsedFormat;     // cached; throws FormatCodeError
export class FormatCodeError extends Error { readonly offset: number }
export interface ParsedFormat {
  sections: FormatSection[];     // 1..4
  unsupported?: true;            // calendar/numeral system out of scope → render as General
}
export interface FormatSection {
  condition?: { op: "<" | "<=" | ">" | ">=" | "=" | "<>"; value: number };
  color?: string; lcid?: number; kind: "number" | "fraction" | "scientific" | "date" | "text" | "general" | "literal";
  tokens: FormatToken[];         // the parsed body (shape internal to parse/render)
}
// numberFormat/odsNumberStyles.ts
export function odsNumberStyleToCode(style: Element, lookup: (name: string) => Element | undefined): string | undefined;
export function odsNumberStyleXml(code: string, name: string): string;
```
`generalText(value)` (SPEC-116) stays exported from `general.ts` as
`formatGeneral(value, { locale: SPREADSHEET_LOCALES["en-US"], dateSystem:
"1900" }).text`.

### The grammar
```
code     := section ( ";" section ){0,3}
section  := bracket* body
bracket  := "[" color "]" | "[" condition "]" | "[$" symbol? ( "-" lcid )? "]" | "[" elapsed "]"
color    := Black | Blue | Cyan | Green | Magenta | Red | White | Yellow | "Color" 1..56   (case-insensitive)
condition:= ( "<" | "<=" | ">" | ">=" | "=" | "<>" ) number
lcid     := 1–8 hex digits | "x-sysdate" | "x-systime"
elapsed  := "h"+ | "m"+ | "s"+                                    (case-insensitive)
body     := ( placeholder | "." | "," | "%" | exponent | "/" | "@" | "General"
            | datetoken | ampm | literal )*
literal  := '"' … '"' | "\" char | "_" char | "*" char | any other char
```
- Date/time tokens (case-insensitive): `y`/`yy` (two-digit year), `yyy`+
  (four-digit); `m`, `mm`, `mmm`, `mmmm` (also 6+), `mmmmm` (first letter);
  `d`, `dd`, `ddd`, `dddd`+; `h`, `hh`+; `s`, `ss`+; `.0`/`.00`/`.000` after
  a seconds token (fractional seconds); `AM/PM`, `am/pm`, `A/P`, `a/p`; the
  LibreOffice-only `Q`/`QQ` (quarter) and `WW` (ISO week) that ODS styles map
  to. A section with any date/time token is a date section; `0#?` there are
  literals except the fractional-second zeros.
- **The minute rule.** `m`/`mm` is minutes when the nearest preceding
  date/time token is an hour (`h`, `hh`, `[h]`) or the nearest following one
  is a second (`s`, `ss`, `[s]`); otherwise it is the month.
- `[$-F800]` and `[$-x-sysdate]` replace the section's date with the locale's
  `longDate`; `[$-F400]` and `[$-x-systime]` with its `longTime`. An LCID whose
  value exceeds `0xFFFF` (a calendar or numeral-system byte), `B1`/`B2`
  prefixes, `e`/`ee`/`g`/`gg`/`ggg`, `[DBNum1-4]` and `[NatNum*]` make the
  code `unsupported`.
- Codes are stored and parsed in the invariant syntax: `.` is the decimal
  point and `,` the grouping/scaling comma whatever the locale; `/` and `:`
  are literal characters (no date-separator substitution).
- Errors (`FormatCodeError`): an unterminated quote or bracket, a fifth
  section, an unknown bracket, a section mixing digit placeholders and `@`,
  an `E` with no exponent digits. Unknown bare letters are lenient literals,
  not errors.

### Sections
- Without conditions: one section → every number; two → (≥ 0; < 0); three →
  (> 0; < 0; = 0); a fourth (or, with fewer, a final section containing `@`)
  → text.
- With conditions on the first and/or second section: the first whose
  condition holds wins; the next unconditioned numeric section is the
  "else"; no match → `overflow: "hash"`.
- **Sign.** The automatic `-` is dropped only in section 2 of a code whose
  first two sections carry no explicit conditions (the classic negative
  section, which writes its own `-` or parentheses); in every other section a
  negative value gets the automatic `-`. `0.00;0.00` on −1 → "1.00";
  `[>100][Red]0;[Blue]0` on −5 → "-5" in blue.
- An empty section renders "" (`;;;` hides everything); a section of only
  literals renders the literals.

### Rendering numbers
1. Normalise: |v| × 100 per `%`, ÷ 1000 per trailing comma (a `,` directly
   after the last integer placeholder, before `.` or the end), then to a
   decimal string of 15 significant digits (`toPrecision(15)`), so `0.1 + 0.2`
   renders as 0.3 everywhere.
2. Round the decimal string half away from zero to the number of fraction
   placeholders (or, for scientific, the mantissa's).
3. Integer part: digits fill the integer placeholders right to left; extra
   leading digits precede the first placeholder; unfilled `0` → "0", `#` →
   nothing, `?` → U+2007 FIGURE SPACE; a `,` between integer placeholders
   groups every three digits with `locale.group`; literals between
   placeholders stay in position (`000-00-0000` on 123456789 →
   "123-45-6789"). An integer part of 0 under `#` only renders nothing
   (`#.##` on 0.5 → ".5").
4. Fraction part: `0` keeps zeros, `#` drops trailing zeros, `?` turns them
   into figure spaces; the point is `locale.decimal` and is shown whenever the
   code has one (`#.##` on 5 → "5.").
5. Scientific (`E+`/`E-`): with n integer placeholders containing a `#` and n
   > 1, the exponent is a multiple of n (`##0.0E+0` on 12345 → "12.3E+3");
   otherwise one integer digit (`0.00E+00` on 12345 → "1.23E+04"); exponent
   padded to its `0` count; `E+` always signs, `E-` only negatives.
6. Fractions: a `/` between placeholders. An integer part exists when a
   placeholder group and a space precede the numerator (`# ?/?`). The
   denominator is fixed when written as digits (`?/8` on 0.3 → "2/8", never
   reduced) or bounded by 10^k − 1 for k placeholders, choosing the best
   rational approximation (continued-fraction convergents and
   semiconvergents; ties → the smaller denominator): `# ??/??` on π →
   "3 14/99", `# ?/?` on 1.5 → "1 1/2", `?/?` on 1.5 → "3/2". A zero
   fraction with an integer part renders the numerator, slash and denominator
   as spaces of the same length; exact 0 → "0" plus those spaces.
7. `General` inside a numeric section renders `formatGeneral` of |v| (sign
   per the rule above) at full width in place: `[Red]General`.
8. `_x` → one space; `*x` → `fill` (the last one in a section wins, the
   others are dropped); `\x` and quoted text → literal.
9. `alignHint` "right". Colours: `[Black]` `#000000`, `[Blue]` `#0000ff`,
   `[Cyan]` `#00ffff`, `[Green]` `#00ff00`, `[Magenta]` `#ff00ff`, `[Red]`
   `#ff0000`, `[White]` `#ffffff`, `[Yellow]` `#ffff00`; `[ColorN]` →
   `palette[N + 7]`.

### Rendering dates and times
- Let k = the section's fractional-second digits (0–3). Total = serial ×
  86 400 seconds rounded half away from zero to 10^−k; days = floor(total /
  86 400) and the non-negative remainder gives h/m/s/fraction — so carries
  reach the date
  (`0.999999` with `h:mm` → "0:00"; `yyyy-mm-dd hh:mm` on 46303.9999999 →
  "2026-10-09 00:00").
- The date part comes from `serialToParts(days, system)`. With
  `dateQuirks: "excel"`: a negative serial in a date or time section → hash;
  serial 0 → year 1900, month 1, day 0 ("1/0/1900"); weekday in the 1900
  system = (days + 6) mod 7 (0 = Sunday; serial 1 is a Sunday, serial 60 a
  Wednesday — Excel's calendar), in the 1904 system (days + 5) mod 7. With
  `"proleptic"`: negative serials are real proleptic Gregorian dates and the
  weekday is the calendar date's. Either way a year above 9999 or below 1 →
  hash.
- `[h]`/`[m]`/`[s]` are totals (floor; `[hh]` pads to two); after `[h]`,
  `mm` is minutes of the hour; a negative value under an elapsed token → hash.
- `h` with `AM/PM`/`A/P` in the section is 12-hour (0 → 12); the marker
  prints in the code's case ("AM"/"PM", "am"/"pm", "A"/"P", "a"/"p") in every
  locale.
- Names: `localeNames(tag)` — `monthsLong`, `monthsGenitive`, `monthsShort`,
  `daysLong`, `daysShort` — from `Intl.DateTimeFormat(tag, { timeZone: "UTC",
  … }).formatToParts` over 2001-mm-15 (months) and 2001-01-07 + d (Sunday
  first), built once per tag. `mmm` → `monthsShort`; `mmmm` →
  `monthsGenitive` when the section also has a `d`/`dd` token, else
  `monthsLong`; `mmmmm` → the first character of `monthsLong`; `ddd`/`dddd` →
  `daysShort`/`daysLong`. A section LCID of 409, 809, 407, 40C, C0A/40A or
  405 takes that locale's names; any other LCID the user's.
- Pinned names (Node full-ICU, the unit suite): en-US "January" … "Sep" …
  "Sunday", "Thu"; de-DE "Januar", "März", "Oktober", short "Jan", "Mär",
  "Okt", "Donnerstag", "Do"; cs-CZ `d. mmmm yyyy` on 46303 → "8. října 2026",
  `mmmm yyyy` → "říjen 2026". Names are CLDR's, not Windows': German `mmm`
  for March is "Mär" where German Excel shows "Mrz" — accepted and stated.
- `Q` → "Q" + quarter, `QQ` the same (the long form is kept for the ODS
  writer only); `WW` → the ISO 8601 week number.
- `alignHint` "right".

### Text, booleans and errors
A string renders through the text section (`@` → the string; `"Name: "@` →
"Name: Bob"), or as itself when there is none; `alignHint` "left". A number
under a code whose only section is a text section renders as General. Booleans
render "TRUE"/"FALSE" and errors their code under any code, centred, in every
locale. `null` → "".

### General
`formatGeneral(value, ctx)`: non-numbers as above. A number: first the
full-width text by SPEC-116's 11-character rule with `locale.decimal`. With
`widthPx`: available = `widthPx − 2 × CELL_PADDING_PX − 1`; a text fits when
its estimated width (SPEC-117's weights × `digitWidthPx`) ≤ available.
If the full text does not fit: (1) fixed notation — drop decimals one at a
time (rounding half away from zero) while the integer part still fits;
(2) scientific — mantissa decimals from 5 down to 0 (`1.23457E+11` …
`1E+11`); (3) `#` × max(1, floor(available / digitWidthPx)) with `overflow:
"hash"`. Worked examples at 64 px, MDW 7 (8.43 digits): `123456789012` →
"1.23E+11"; `1234.56789` → "1234.568"; `123456789` → "1.23E+08";
`0.000123456` → "0.000123"; at 26 px: `12345` → "###".

Non-General: a number or date text that does not fit → `#` × max(1,
floor(available / digitWidthPx)), `overflow: "hash"`. A `fill` counts one
fill character toward the fit.

### Built-in formats
`BUILTIN_NUMBER_FORMATS` (en-US spellings; ECMA-376 §18.8.30 lists `14` as
`mm-dd-yy`, but Excel renders and SpreadsheetML consumers expect the
locale's short date — this table holds what en-US Excel displays):
```
0  General            9  0%                 18 h:mm AM/PM        37 #,##0 ;(#,##0)
1  0                  10 0.00%              19 h:mm:ss AM/PM     38 #,##0 ;[Red](#,##0)
2  0.00               11 0.00E+00           20 h:mm              39 #,##0.00;(#,##0.00)
3  #,##0              12 # ?/?              21 h:mm:ss           40 #,##0.00;[Red](#,##0.00)
4  #,##0.00           13 # ??/??            22 m/d/yyyy h:mm     45 mm:ss
14 m/d/yyyy           15 d-mmm-yy           16 d-mmm             46 [h]:mm:ss
17 mmm-yy             47 mmss.0             48 ##0.0E+0          49 @
5  "$"#,##0_);\("$"#,##0\)                  41 _(* #,##0_);_(* \(#,##0\);_(* "-"_);_(@_)
6  "$"#,##0_);[Red]\("$"#,##0\)             42 _("$"* #,##0_);_("$"* \(#,##0\);_("$"* "-"_);_(@_)
7  "$"#,##0.00_);\("$"#,##0.00\)            43 _(* #,##0.00_);_(* \(#,##0.00\);_(* "-"??_);_(@_)
8  "$"#,##0.00_);[Red]\("$"#,##0.00\)       44 _("$"* #,##0.00_);_("$"* \(#,##0.00\);_("$"* "-"??_);_(@_)
```
`builtinFormatCode(id, locale)`: 14 → `locale.shortDate`; 22 →
`locale.shortDateTime`; 5–8 and 41–44 → the table above with `"$"` replaced
by the locale's quoted symbol for a prefix locale, or this suffix template
for a suffix locale (S = the quoted symbol): 5 `#,##0 S;\-#,##0 S`, 6 `#,##0
S;[Red]\-#,##0 S`, 7 `#,##0.00 S;\-#,##0.00 S`, 8 `#,##0.00 S;[Red]\-#,##0.00
S`, 41 `_-* #,##0_-;\-* #,##0_-;_-* "-"_-;_-@_-`, 42 `_-* #,##0 S_-;\-* #,##0
S_-;_-* "-" S_-;_-@_-`, 43 `_-* #,##0.00_-;\-* #,##0.00_-;_-* "-"??_-;_-@_-`,
44 `_-* #,##0.00 S_-;\-* #,##0.00 S_-;_-* "-"?? S_-;_-@_-`. Every other id is
locale-independent. `formatCodeOf(style, locale)` = `builtinFormatCode(
style.builtinFormatId, locale)` when that is set, else `style.numberFormat`.

### Locales
| Tag | LCID | decimal | group | list | order | date sep | shortDate | longDate | shortTime | longTime | shortDateTime | currency |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| en-US | 409 | `.` | `,` | `,` | MDY | `/` | `m/d/yyyy` | `dddd, mmmm d, yyyy` | `h:mm AM/PM` | `h:mm:ss AM/PM` | `m/d/yyyy h:mm` | `$` prefix, USD |
| en-GB | 809 | `.` | `,` | `,` | DMY | `/` | `dd/mm/yyyy` | `dd mmmm yyyy` | `hh:mm` | `hh:mm:ss` | `dd/mm/yyyy hh:mm` | `£` prefix, GBP |
| de-DE | 407 | `,` | `.` | `;` | DMY | `.` | `dd.mm.yyyy` | `dddd, d. mmmm yyyy` | `hh:mm` | `hh:mm:ss` | `dd.mm.yyyy hh:mm` | `€` suffix, EUR |
| fr-FR | 40C | `,` | U+202F | `;` | DMY | `/` | `dd/mm/yyyy` | `dddd d mmmm yyyy` | `hh:mm` | `hh:mm:ss` | `dd/mm/yyyy hh:mm` | `€` suffix, EUR |
| es-ES | C0A | `,` | `.` | `;` | DMY | `/` | `dd/mm/yyyy` | `dddd, d "de" mmmm "de" yyyy` | `h:mm` | `h:mm:ss` | `dd/mm/yyyy h:mm` | `€` suffix, EUR |
| cs-CZ | 405 | `,` | U+00A0 | `;` | DMY | `.` | `dd.mm.yyyy` | `dddd d. mmmm yyyy` | `h:mm` | `h:mm:ss` | `dd.mm.yyyy h:mm` | `Kč` suffix, CZK |

```ts
export interface SpreadsheetLocale {
  tag: SpreadsheetLocaleTag; lcid: number;
  decimal: string; group: string; listSeparator: string;     // listSeparator: CSV (SPEC-126); never formulas
  dateOrder: "MDY" | "DMY" | "YMD"; dateSeparator: string; timeSeparator: ":";
  shortDate: string; longDate: string; shortTime: string; longTime: string; shortDateTime: string;  // codes
  amPm: [string, string];                                    // "AM", "PM" for all six (input accepts either case)
  currencySymbol: string; currencyPosition: "prefix" | "suffix"; currencyCode: string;  // suffix = after a space
}
```
SPEC-122's input parser additionally accepts U+0020, U+00A0 and U+202F as
the group separator for fr-FR and cs-CZ, since users type a plain space.

### editText
`editText(cell, ctx)` with `code = formatCodeOf(ctx.style ?? {}, ctx.locale)`
— every caller (the formula bar, SPEC-119's `formulaBarText`, SPEC-122's
editor, SPEC-134's search) passes the cell's effective style in `ctx.style`:
- string → itself, prefixed `'` when `ctx.style?.quotePrefix`; boolean →
  "TRUE"/"FALSE"; error → its code; `null` → "".
- number under a code with a date token → `shortDate`, + " " + `longTime`
  when the time part rounds to a non-zero second; under a time-only code
  (elapsed included) → `longTime` when the serial is below 1, else
  `shortDate` + " " + `longTime` (`[h]:mm:ss` on 1.5 → "1/1/1900 12:00:00
  PM", as Excel's bar shows); hash cases (negative xlsx dates) → plain
  notation.
- number under a code with `%` outside quotes → (n × 100) in plain notation +
  "%" (`0.125` → "12.5%").
- any other number → plain notation: up to 15 significant digits, no
  grouping, `locale.decimal`, scientific (`1.23456789012346E+17`) only when
  |n| ≥ 1e15 or 0 < |n| < 1e−9.
SPEC-119 extends the formula bar to show `=` + the formula for a formula
cell; `editText` itself never reads `Cell.formula`.

### Reading xlsx number formats
`xl/styles.xml` `numFmts/numFmt@numFmtId` + `@formatCode` (attribute-decoded)
are read into a map; each `cellXfs/xf@numFmtId` (absent → 0) resolves: map
entry → `numberFormat` (any id, including < 164); else a locale-dependent
built-in → `builtinFormatId = id` and `numberFormat` = the en-US code; else
another built-in → its code; 0, unknown and CJK ids → absent. A resolved code
equal to `General` (case-insensitive) is stored absent. `applyNumberFormat`
is ignored, as SPEC-117 ignores every `apply*`.

### ODS number styles
`odsNumberStyleToCode` walks one data style (`number:*-style` in
`content.xml` `office:automatic-styles` or `styles.xml`) and returns its code:

| ODS | Code |
| --- | --- |
| `number:number` without `number:decimal-places` | `General` |
| `number:number` `decimal-places` d, `min-decimal-places` m (or `loext:`), `min-integer-digits` i, `grouping="true"` | integer `0`×i (`#` when i = 0), padded with `#` to `#,##0` form when grouping; `.` + `0`×m + `#`×(d − m) when d > 0 |
| `number:display-factor` 1000^k | k trailing commas |
| `number:embedded-text number:position="p"` | the quoted text placed p digits left of the point |
| `number:scientific-number` d, i, `min-exponent-digits` e, `exponent-interval` n | `#`×(n − i) + `0`×i, decimals as above, `E+` + `0`×e |
| `number:fraction` i, `min-numerator-digits` a, `min-denominator-digits` b, `denominator-value` v | `#`/`0` integer part (none when the attribute is absent) + space, `?`×a, `/`, `?`×b or v |
| `number:percentage-style` | the number part + `%` |
| `number:currency-style` / `number:currency-symbol` (`number:language`, `number:country`) | `[$sym-LCID]` when the pair maps to a known LCID, else `[$sym]`, at its position |
| `number:boolean-style` | `"TRUE";"TRUE";"FALSE"` |
| `number:text-style` / `number:text-content` | `@` |
| `number:text` | bare when it is only `-` `/` `:` `.` `,` `(` `)` and spaces in a date/time section, or only `-` `(` `)` `$` `+` and spaces in a number section; else `"…"` with `\"` escapes |
| `number:fill-character` | `*x` |
| `number:day` / `month` (`textual`) / `year` / `day-of-week` / `quarter` / `week-of-year` (`number:style="long"` → long form) | `d`/`dd`, `m`/`mm`/`mmm`/`mmmm`, `yy`/`yyyy`, `ddd`/`dddd`, `Q`/`QQ`, `WW` |
| `number:hours` / `minutes` / `seconds` (`decimal-places`) / `am-pm` | `h`/`hh`, `m`/`mm`, `s`/`ss` + `.0…`, `AM/PM` |
| `number:time-style number:truncate-on-overflow="false"` | the first time token bracketed (`[h]`, `[mm]`, `[ss]`) |
| `number:era`, an unknown child | General (the source style is reused unchanged by `sourceStyleName`) |
| `style:text-properties@fo:color` | the named colour of that exact hex, else `[ColorN]` for a default-palette match, else none |
| `number:language`/`number:country` on a style with textual names | `[$-LCID]` prefix when the pair maps |
| `style:map style:condition="value()OPn" style:apply-style-name="X"` | section `[OPn]` + X's code, the main style last |

LibreOffice's sign pattern maps to the conventional unconditioned form: the
exact maps LibreOffice writes for two sections (`value()>=0`) and three
(`value()>0`, `value()<0`) become plain sections (`0.00;[Red]-0.00`), because
its negative style carries its own `-` literal and the sign rule drops the
automatic one only in an unconditioned second section. Ignored attributes
(`number:automatic-order`, `format-source`, `transliteration-*`, `title`,
`volatile`) survive in the source style. Known LCIDs: 409 en-US, 809 en-GB,
407 de-DE, 40C fr-FR, C0A es-ES, 405 cs-CZ.

`odsNumberStyleXml(code, name)` is the inverse: per section, the style kind
is `date-style` (any date token) / `time-style` (time tokens only;
`truncate-on-overflow="false"` for elapsed) / `percentage-style` (`%`) /
`currency-style` (a `[$sym…]` with a symbol) / `text-style` (`@`) /
`number-style`; sections 1..n−1 become `${name}P0`, `${name}P1`… and the main
style carries one `style:map` each (`value()>=0` for an unconditioned two-
section code; `value()>0`, `value()<0` for three; the text section, when
present, is the main style and the numeric sections all map from it, as
LibreOffice writes it). `Q`/`QQ`/`WW` have no xlsx form: SPEC-122's xlsx
writer drops them, the one stated loss.

### Implied formats
Applied by the readers when the source marks a kind of value but gives it
no format: ODS `office:value-type="date"` with no data style → `yyyy-mm-dd`
(+ ` hh:mm:ss` when the date-value has a time); `"time"` → `[hh]:mm:ss`;
`"percentage"` → `0.00%`; xlsx `t="d"` with a General style → `yyyy-mm-dd`
(+ ` hh:mm:ss`). The reader appends one derived style per (base style,
implied code) pair to `model.styles` (`numberFormat` set, `sourceXf`/
`sourceStyleName` absent) and points the cell at it; a clean sheet still
downloads byte-identical.

### The preference
`Preferences.spreadsheetLocale: SpreadsheetLocaleTag` (default
`DEFAULT_SPREADSHEET_LOCALE`), validator `oneOf(SPREADSHEET_LOCALE_TAGS)` —
`preferences.ts` (app layer) imports the type and list from
`spreadsheet/numberFormat/locale.ts`. `PreferencesDialog.vue` gains a
"Spreadsheets" section after "Editing" with "Number and date format"
(`preferences-spreadsheet-locale`); option labels are the language name and
a live sample — "English (United States) — 1,234.56 · 10/8/2026",
"Deutsch (Deutschland) — 1.234,56 · 08.10.2026", … — formatted by
`formatValue`. `useSpreadsheetLocale.ts` returns the reactive
`SpreadsheetLocale` and `setLocale(tag)` (→ `savePreference`); the status-bar
control `sheet-locale` shows the tag and opens a menu of `menuitemradio`
items `sheet-locale-<tag>`. A change re-renders through
`onPreferencesChange` (other tabs included).

### Reader version
Bumps `SPREADSHEET_READER_VERSION` (to the next integer) and appends a
`ReaderUpgrade` whose `upgradeDirtySheet` throws — the same choice and
reason as SPEC-117: no sheet can be dirty before SPEC-121, every read path
upgrades, so the re-read covers every real record and the throwing step only
keeps the registry contiguous.

### Fixtures
- `spec118-formats.xlsx` (Excel-shaped): `numFmts` with custom ids ≥ 164
  (`[$-F800]dddd\,\ mmmm\ dd\,\ yyyy`, `[$€-407]#,##0.00`, `0.00;[Red]\(0.00\)`,
  `# ??/??`, `[>100][Red]0;[Blue]0`, `;;;`) and explicit entries for 43 and
  44 as Excel writes them; cells on ids 9, 10, 11, 14, 22, 46 and 49;
  `A2:A4` = 46303.5729166667 under 14, 22 and the F800 code; `B2` = −1234.5
  under `#,##0.00_);[Red]\(#,##0.00\)` and `B5` = 1234.5 under `#,##0.00`;
  `C2` = 0.125 under 10; `D2` = 1.5 under 46; serials 0, 60 and 61 under
  `d-mmm-yyyy`; a narrow column F (26 px) and G (30 px) with
  `123456789012` and a date.
- `spec118-generated.xlsx` (openpyxl-shaped): inline strings, `t="d"` cells
  with a General style, no theme, no `numFmts`.
- `spec118-formats.ods` (LibreOffice-shaped): `N…` data styles with
  `style:map`s (`N2P0`), a red negative, currency-style with
  `number:language="de" number:country="DE"`, a long textual month date with
  `number:language`, a `truncate-on-overflow="false"` time style, fraction,
  scientific, percentage, boolean, text styles, a quarter + year style, a
  week-of-year style, an era style, a date cell without a data style, and
  1899-12-25 (an ODS-only date).

### Tests
- `parse.spec.ts`: every grammar row, the minute rule, each error kind with
  its offset, every unsupported form, the cache limit.
- `render.spec.ts`: a table of (value, code, locale, width) → text/colour/
  fill/overflow covering every example in this Reference, the sign rule,
  the six locales' separators, de-DE and fr-FR booleans and errors.
- `odsNumberStyles.spec.ts`: each table row both ways, and the corpus round
  trip — the locale-independent built-ins plus `0.00;[Red]-0.00`,
  `0;-0;"zero";"text: "@`, `#,##0,"K"`, `[h]:mm:ss`, `dddd d mmmm yyyy`,
  `# ?/8`, `0.00E+00`, `[$€-407]#,##0.00`.
- `locale.spec.ts`: the table, the name tables (en-US, de-DE, the cs-CZ
  genitive). `builtin.spec.ts`: every id, the locale-dependent ones in all
  six locales. Round trip: part-level `equalBytes` for the three fixtures.

### Specs this changes
- `specs/116-…` — `generalText` delegates to `formatGeneral`; `buildPreview`
  takes a locale and formats values; the reader version moves on.
- `specs/117-…` — `CellStyle.numberFormat`/`builtinFormatId` are populated;
  the grid, the formula bar (`editText`) and the status-bar Sum/Average
  render formatted; cells may point at implied-format styles.
- `specs/066-…` — `Preferences` gains `spreadsheetLocale` and the dialog a
  "Spreadsheets" section.
- For the writers and commands that follow (restated so they need not guess):
  SPEC-122's xlsx style writer writes `numFmtId = builtinFormatId` with no
  `numFmt` entry when that field is set, and may reuse a built-in id for a
  `numberFormat` equal to a **locale-independent** built-in only (never 5–8,
  14, 22 or 41–44, whose meaning follows the reader's locale); SPEC-123's
  `patchStyle` clears `builtinFormatId` whenever it sets `numberFormat`.
