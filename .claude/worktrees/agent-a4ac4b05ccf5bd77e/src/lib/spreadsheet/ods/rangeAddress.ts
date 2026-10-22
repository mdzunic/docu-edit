// ODS cell-range addresses (SPEC-121 "ODS range addresses") — THE parser and
// printer for the address lists ODF keeps outside formulas
// (`table:cell-range-address`, `table:base-cell-address`,
// `table:target-range-address`, `table:end-cell-address`,
// `chart:values-cell-range-address`, `draw:notify-on-update-of-ranges`, …).
// SPEC-121's sheet-reference pass (formula/sheetRefs.ts) is the first caller;
// SPEC-124's retained-reference rewrite and SPEC-128–136 reuse it unchanged.
//
// - One entry is `[$]SHEET.CELL`, `[$]SHEET.CELL:[[$]SHEET].CELL` or the
//   sheet-less `.CELL[:.CELL]`. SHEET is bare — a run of anything but `.`,
//   `:`, `'` and whitespace, not starting with `$` — or single-quoted with
//   `''` for a quote. CELL is `[$]LETTERS[$]DIGITS`; a whole column
//   `[$]LETTERS` or whole row `[$]DIGITS` only as a range whose two ends both
//   have that shape (the formula dialect's `[.A:.C]`). Letters in any case;
//   every part inside A1..XFD1048576.
// - A list is entries separated by runs of spaces outside quotes (a leading
//   or trailing run is ignored). splitOdsRangeAddressList keeps the runs, so
//   a rewrite can leave every entry it does not touch byte-for-byte.
// - To canonical (odsRangeAddressToCanonical): formatReference's text, so
//   tokens.ts reads each entry back as one `ref` token. A sheet's `$`
//   (absolute sheet) is dropped; the cell parts keep theirs. An end with no
//   sheet, or with the start's (compared case-insensitively — sheet names
//   are unique that way), is on the start's sheet; another sheet makes a 3-D
//   reference (`Sheet1:Sheet3!A1:B2`), whose single cell is written once
//   (`$Jan.B5:$Dec.B5` → `Jan:Dec!B5`). Two or more entries are a union in
//   parentheses, `(a,b)`; a single entry has none.
// - null for anything else — notably LibreOffice's deleted-sheet form
//   `$#REF!.A1`, an unterminated quote, an empty sheet name, and a name
//   canonical text cannot carry (a `:` reads as a 3-D span, a leading `[1]`
//   as another workbook; neither application allows either in a name).
// - From canonical (canonicalToOdsRangeAddress): a reference, or a union of
//   references optionally in one pair of parentheses, each written
//   `$Sheet.start:.end` — the sheet on the start only, absolute as
//   LibreOffice writes it; `.start:.end` without one; `$Jan.A1:$Mar.B2` for
//   3-D (a 3-D single cell as a range); `$Sheet.A1` for a single cell; the
//   parts' `$` as canonical has them; entries joined by one space. Names are
//   quoted by print.ts's predicate, and also when they hold a `.` (the ODS
//   sheet separator, which canonical text allows bare), `'` doubled.
//   Anything else (functions, operators, names, errors) is null.
// - Known non-identities of ods → canonical → ods: a relative sheet gains
//   its `$`, a repeated end sheet is dropped (`Sheet1.A1:Sheet1.B3` →
//   `$Sheet1.A1:.B3`), letters are upper-cased, a separating run becomes one
//   space. canonical → ods → canonical is the identity on formatReference's
//   text and unions of it, but for a 3-D range of one cell (`Jan:Dec!B5:B5`
//   comes back `Jan:Dec!B5`).
import { colToLetters, lettersToCol } from "../address";
import { formatReference, sheetNameNeedsQuotes } from "../formula/print";
import type { RefPart, ReferenceNode } from "../formula/references";
import { scanQuoted, tokenize } from "../formula/tokens";
import { MAX_ROWS } from "../model";

/** What an endpoint's CELL part spells. */
type PartShape = "cell" | "column" | "row";

interface PartScan {
  /** Index past the part. */
  end: number;
  part: RefPart;
  shape: PartShape;
}

interface EndpointScan extends PartScan {
  /** Unquoted and un-escaped; absent for `.A1`. */
  sheet?: string;
}

function isAsciiLetter(code: number): boolean {
  return (code >= 0x41 && code <= 0x5a) || (code >= 0x61 && code <= 0x7a);
}

function isDigit(code: number): boolean {
  return code >= 0x30 && code <= 0x39;
}

/** `.`, `:`, `'` or whitespace — the characters a bare sheet name stops at. */
function endsBareName(code: number): boolean {
  return code === 0x2e || code === 0x3a || code === 0x27 || code === 0x20 || code === 0x09 || code === 0x0a || code === 0x0d;
}

/** The sheet name at `i` (past its `$`): quoted, or a bare run; null when
 * there is none, it never closes, or it is LibreOffice's deleted `#REF!`. */
function scanSheetName(text: string, i: number): { end: number; name: string } | null {
  if (text.charCodeAt(i) === 0x27) {
    const quoted = scanQuoted(text, i);
    return quoted && quoted.value !== "" ? { end: quoted.end, name: quoted.value } : null;
  }
  let j = i;
  while (j < text.length && !endsBareName(text.charCodeAt(j))) {
    j += 1;
  }
  const name = text.slice(i, j);
  if (name === "" || name.startsWith("$") || name.toUpperCase() === "#REF!") {
    return null;
  }
  return { end: j, name };
}

/** `[$]LETTERS[$]DIGITS`, `[$]LETTERS` or `[$]DIGITS` at `i`, inside the
 * grid; a row-only part's `$` is its row's. */
function scanPart(text: string, i: number): PartScan | null {
  let j = i;
  const leadingDollar = text.charCodeAt(j) === 0x24;
  if (leadingDollar) {
    j += 1;
  }
  const lettersStart = j;
  while (j < text.length && isAsciiLetter(text.charCodeAt(j))) {
    j += 1;
  }
  const letters = text.slice(lettersStart, j);
  let rowAbs = letters === "" && leadingDollar;
  if (letters !== "" && text.charCodeAt(j) === 0x24) {
    rowAbs = true;
    j += 1;
  }
  const digitsStart = j;
  while (j < text.length && isDigit(text.charCodeAt(j))) {
    j += 1;
  }
  const digits = text.slice(digitsStart, j);
  if (digits === "" && (letters === "" || rowAbs)) {
    return null;
  }
  const col = letters === "" ? undefined : lettersToCol(letters);
  const row = digits === "" ? undefined : Number(digits) - 1;
  if (col === null || (row !== undefined && (row < 0 || row >= MAX_ROWS))) {
    return null;
  }
  const part: RefPart = { rowAbs, colAbs: letters !== "" && leadingDollar };
  if (col !== undefined) {
    part.col = col;
  }
  if (row !== undefined) {
    part.row = row;
  }
  return { end: j, part, shape: col === undefined ? "row" : row === undefined ? "column" : "cell" };
}

/** One endpoint at `i`: `.CELL`, or `[$]SHEET.CELL`. */
function scanEndpoint(text: string, i: number): EndpointScan | null {
  let j = i;
  let sheet: string | undefined;
  if (text.charCodeAt(j) !== 0x2e) {
    if (text.charCodeAt(j) === 0x24) {
      j += 1;
    }
    const name = scanSheetName(text, j);
    if (!name || text.charCodeAt(name.end) !== 0x2e) {
      return null;
    }
    sheet = name.name;
    j = name.end;
  }
  const part = scanPart(text, j + 1);
  if (!part) {
    return null;
  }
  return sheet === undefined ? part : { ...part, sheet };
}

function samePart(a: RefPart, b: RefPart): boolean {
  return a.row === b.row && a.col === b.col && a.rowAbs === b.rowAbs && a.colAbs === b.colAbs;
}

/** Whether canonical text can carry a sheet name (see the module header). */
function canonicalCarries(name: string | undefined): boolean {
  return name === undefined || (name !== "" && !name.includes(":") && !/^\[\d+\]/.test(name));
}

/** One cell-range address (a single entry, not a list) → its reference,
 * sheet names unquoted; null when the grammar cannot read it or canonical
 * text cannot carry it (see the module header). */
export function parseOdsRangeAddress(entry: string): ReferenceNode | null {
  const first = scanEndpoint(entry, 0);
  if (!first) {
    return null;
  }
  const ref: ReferenceNode = { start: first.part };
  if (first.sheet !== undefined) {
    ref.sheet = first.sheet;
  }
  if (first.end === entry.length) {
    return first.shape === "cell" && canonicalCarries(ref.sheet) ? ref : null;
  }
  if (entry.charCodeAt(first.end) !== 0x3a) {
    return null;
  }
  const second = scanEndpoint(entry, first.end + 1);
  if (!second || second.end !== entry.length || second.shape !== first.shape) {
    return null;
  }
  if (second.sheet !== undefined) {
    if (first.sheet === undefined) {
      return null;
    }
    if (second.sheet.toUpperCase() !== first.sheet.toUpperCase()) {
      ref.sheetEnd = second.sheet;
    }
  }
  // A 3-D single cell can only be written as a range in ODS.
  if (!(ref.sheetEnd !== undefined && first.shape === "cell" && samePart(first.part, second.part))) {
    ref.end = second.part;
  }
  return canonicalCarries(ref.sheet) && canonicalCarries(ref.sheetEnd) ? ref : null;
}

/** A sheet name as an address writes it: `$` (absolute), then bare or
 * quoted (see the module header). */
function odsSheetName(name: string): string {
  const quoted = sheetNameNeedsQuotes(name) || name.includes(".");
  return `$${quoted ? `'${name.replaceAll("'", "''")}'` : name}`;
}

function partText(part: RefPart): string {
  let text = "";
  if (part.col !== undefined) {
    text += `${part.colAbs ? "$" : ""}${colToLetters(part.col)}`;
  }
  if (part.row !== undefined) {
    text += `${part.rowAbs ? "$" : ""}${part.row + 1}`;
  }
  return text;
}

/** One reference as a cell-range address: `$Sheet.A1:.B2`, `.A1`,
 * `$Jan.A1:$Mar.A1` (see the module header). */
export function formatOdsRangeAddress(ref: ReferenceNode): string {
  const whole = ref.start.row === undefined || ref.start.col === undefined;
  const end = ref.end ?? (whole || ref.sheetEnd !== undefined ? ref.start : undefined);
  let text = `${ref.sheet === undefined ? "" : odsSheetName(ref.sheet)}.${partText(ref.start)}`;
  if (end) {
    text += `:${ref.sheetEnd === undefined ? "" : odsSheetName(ref.sheetEnd)}.${partText(end)}`;
  }
  return text;
}

/** A list split at its top-level space runs (outside quotes): entries at
 * even indices, the runs as written at odd ones — joining the pieces gives
 * the list back; a leading or trailing run leaves an empty entry at its end. */
export function splitOdsRangeAddressList(list: string): string[] {
  const pieces: string[] = [];
  let quoted = false;
  let start = 0;
  let i = 0;
  while (i < list.length) {
    const code = list.charCodeAt(i);
    if (code === 0x27) {
      // `''` inside quotes toggles twice: still quoted.
      quoted = !quoted;
    } else if (code === 0x20 && !quoted) {
      let j = i;
      while (list.charCodeAt(j) === 0x20) {
        j += 1;
      }
      pieces.push(list.slice(start, i), list.slice(i, j));
      start = j;
      i = j;
      continue;
    }
    i += 1;
  }
  pieces.push(list.slice(start));
  return pieces;
}

/** "Sheet1.$A$1:.$B$3", "$'Q1 Data'.A1:'Q1 Data'.B3", "Sheet1.A1:Sheet1.B3 Sheet1.D1:Sheet1.D3"
 *  → "Sheet1!$A$1:$B$3", "'Q1 Data'!A1:B3", "(Sheet1!A1:B3,Sheet1!D1:D3)"; null if unparseable. */
export function odsRangeAddressToCanonical(text: string): string | null {
  const refs: string[] = [];
  const pieces = splitOdsRangeAddressList(text);
  for (let i = 0; i < pieces.length; i += 2) {
    const entry = pieces[i]!;
    if (entry === "") {
      continue;
    }
    const ref = parseOdsRangeAddress(entry);
    if (!ref) {
      return null;
    }
    refs.push(formatReference(ref));
  }
  if (refs.length === 0) {
    return null;
  }
  return refs.length === 1 ? refs[0]! : `(${refs.join(",")})`;
}

/** The inverse; writes "$Sheet1.$A$1:.$B$3" (sheet on the start only, `$` as canonical has it),
 *  a union as a space-separated list. */
export function canonicalToOdsRangeAddress(canonical: string): string | null {
  let tokens = tokenize(canonical).filter((token) => token.kind !== "end");
  if (tokens[0]?.kind === "open" && tokens.at(-1)?.kind === "close") {
    tokens = tokens.slice(1, -1);
  }
  if (tokens.length % 2 === 0) {
    return null;
  }
  const entries: string[] = [];
  for (const [index, token] of tokens.entries()) {
    if (index % 2 === 1) {
      if (token.kind !== "union") {
        return null;
      }
    } else if (token.kind === "ref" && token.ref && canonicalCarries(token.ref.sheet) && canonicalCarries(token.ref.sheetEnd)) {
      entries.push(formatOdsRangeAddress(token.ref));
    } else {
      return null;
    }
  }
  return entries.join(" ");
}
