// Table row sorting (SPEC-076): a pure, stable multi-key sort over a table's
// rows. Framework-free by contract (AC-42, AC-78) — no Vue import, no DOM,
// unit-tested directly.
//
// Rows move whole. A row's cells, their blocks, borders, shading, spans,
// minHeightCm, preserved docxTrPr/odtStyleName and any pending tracked
// changes all travel together, because the sort permutes the `rows` array and
// never reaches inside a row (AC-48).
import { DEFAULT_LANGUAGE } from "./core/language";
import type { Table, TableRow } from "./documentModel";

/** One sort key: which column, which direction, and how its text is read. */
export interface SortKey {
  columnIndex: number;
  direction: "asc" | "desc";
  type: "text" | "number" | "date";
}

/** Typed empty seed for `SortKey[]` refs in `.vue` scripts (espree bars the
 * annotation there — same idiom as documentModel's `noNamedStyles`). */
export function noSortKeys(): SortKey[] {
  return [];
}

/** Typed pass-through for keys arriving from a `.vue` template's plain
 * `Array` prop or from its own untyped draft state (espree bars both the
 * annotation and the cast there). */
export function asSortKeys(value: unknown): SortKey[] {
  return (value ?? []) as SortKey[];
}

/** How many leading rows the table's contiguous `isHeader` run covers. */
function leadingHeaderRows(table: Table): number {
  let count = 0;
  while (count < table.rows.length && table.rows[count].isHeader) {
    count += 1;
  }
  return count;
}

/** The first row index the sort may move (AC-41): every row after the
 * leading `isHeader` run when the dialog's header toggle is on, and row 0
 * when it is off. */
export function sortRangeStart(table: Table, hasHeader: boolean): number {
  return hasHeader ? leadingHeaderRows(table) : 0;
}

/** How many rows the sort range holds — fewer than two is "there is nothing
 * to sort" (AC-50). */
export function sortRangeRowCount(table: Table, hasHeader: boolean): number {
  return Math.max(0, table.rows.length - sortRangeStart(table, hasHeader));
}

/** Whether the sort range contains a vertically merged cell (AC-49): a cell
 * with `rowSpan > 1`, or a `covered` slot whose owner spans rows. Moving
 * either would tear the merge apart, so the dialog refuses instead. A cell
 * with only `colSpan > 1` is untouched by a row move and never refuses. */
export function sortRangeHasMergedRows(table: Table, hasHeader: boolean): boolean {
  const start = sortRangeStart(table, hasHeader);
  for (let rowIndex = start; rowIndex < table.rows.length; rowIndex += 1) {
    for (const cell of table.rows[rowIndex].cells) {
      if ((cell.rowSpan ?? 1) > 1) {
        return true;
      }
    }
  }
  // A covered slot inside the range whose owner starts above it (possibly
  // above the range) is the other half of the same tear.
  for (let rowIndex = start; rowIndex < table.rows.length; rowIndex += 1) {
    const row = table.rows[rowIndex];
    for (let colIndex = 0; colIndex < row.cells.length; colIndex += 1) {
      if (!row.cells[colIndex].covered) {
        continue;
      }
      for (let above = rowIndex - 1; above >= 0; above -= 1) {
        const owner = table.rows[above].cells[colIndex];
        if (owner && !owner.covered) {
          if (above + (owner.rowSpan ?? 1) > rowIndex) {
            return true;
          }
          break;
        }
      }
    }
  }
  return false;
}

/** A cell's comparable text: its paragraphs' run text, joined by a space so
 * two paragraphs never read as one concatenated word. */
export function sortCellText(row: TableRow, columnIndex: number): string {
  const cell = row.cells[columnIndex];
  if (!cell) {
    return "";
  }
  return cell.blocks
    .map((paragraph) => paragraph.runs.map((run) => run.text).join(""))
    .join(" ")
    .trim();
}

// Currency symbols, percent signs and spaces (including the non-breaking and
// narrow no-break spaces real documents use as group separators) that a
// lenient number parse strips before looking at the digits.
const NUMBER_NOISE = /[\s   $£€¥₹%]/g;

/** A lenient number parse (AC-44): trims, strips currency/percent/space
 * noise, reads a leading sign or accounting parentheses, and works out
 * whether `,` and `.` are grouping or decimal separators — so "1,234.5",
 * "1.234,5", "(42)", "-3" and "€ 1 234,50" all parse. Returns null for text
 * that is not a number at all. */
export function parseSortNumber(text: string): number | null {
  let body = text.trim().replace(NUMBER_NOISE, "");
  if (body.length === 0) {
    return null;
  }
  let sign = 1;
  if (/^\(.*\)$/.test(body)) {
    sign = -1;
    body = body.slice(1, -1);
  }
  if (body.startsWith("-")) {
    sign = -sign;
    body = body.slice(1);
  } else if (body.startsWith("+")) {
    body = body.slice(1);
  }
  const hasComma = body.includes(",");
  const hasDot = body.includes(".");
  if (hasComma && hasDot) {
    // The rightmost of the two is the decimal separator; the other groups.
    const decimal = body.lastIndexOf(",") > body.lastIndexOf(".") ? "," : ".";
    const grouping = decimal === "," ? "." : ",";
    body = body.split(grouping).join("").replace(decimal, ".");
  } else if (hasComma) {
    body = /^\d{1,3}(,\d{3})+$/.test(body) ? body.split(",").join("") : body.replace(",", ".");
  } else if (hasDot && /^\d{1,3}(\.\d{3})+$/.test(body)) {
    body = body.split(".").join("");
  }
  if (!/^\d*\.?\d+$/.test(body) && !/^\d+\.$/.test(body)) {
    return null;
  }
  const value = Number(body);
  return Number.isFinite(value) ? sign * value : null;
}

const MONTHS = [
  "january",
  "february",
  "march",
  "april",
  "may",
  "june",
  "july",
  "august",
  "september",
  "october",
  "november",
  "december",
];

function monthFromName(name: string): number | null {
  const lower = name.toLowerCase();
  const index = MONTHS.findIndex((month) => month === lower || month.slice(0, 3) === lower.slice(0, 3));
  return index === -1 ? null : index + 1;
}

/** Whether the document's language puts the month first in a numeric date —
 * the US convention (and the Philippines'), day-first everywhere else. The
 * language tag is the only signal available, and it is exactly what
 * "the document language's common formats" (AC-45) means here. */
function monthFirst(language: string): boolean {
  const region = language.split("-")[1]?.toUpperCase();
  return region === "US" || region === "PH";
}

function utc(year: number, month: number, day: number): number | null {
  if (month < 1 || month > 12 || day < 1 || day > 31) {
    return null;
  }
  const full = year < 100 ? (year < 70 ? 2000 + year : 1900 + year) : year;
  const value = Date.UTC(full, month - 1, day);
  return Number.isFinite(value) ? value : null;
}

/** Parses ISO 8601 and the document language's common formats (AC-45),
 * returning a UTC timestamp or null:
 *
 * - `YYYY-MM-DD` (and any ISO datetime prefixed by it);
 * - `D/M/YYYY` or `M/D/YYYY` — which, decided by the language's region —
 *   with `/`, `.` or `-` separators and a 2- or 4-digit year;
 * - `D Month YYYY` and `Month D, YYYY`.
 *
 * Deliberately not `new Date(text)`: its behaviour on non-ISO input is
 * implementation-defined, so a sort built on it would order differently in
 * different browsers.
 */
export function parseSortDate(text: string, language: string = DEFAULT_LANGUAGE): number | null {
  const value = text.trim();
  if (value.length === 0) {
    return null;
  }
  const iso = /^(\d{4})-(\d{2})-(\d{2})(?:[T ].*)?$/.exec(value);
  if (iso) {
    return utc(Number(iso[1]), Number(iso[2]), Number(iso[3]));
  }
  const numeric = /^(\d{1,4})[/.-](\d{1,2})[/.-](\d{2,4})$/.exec(value);
  if (numeric) {
    const [, a, b, c] = numeric;
    if (a.length === 4) {
      return utc(Number(a), Number(b), Number(c));
    }
    return monthFirst(language)
      ? utc(Number(c), Number(a), Number(b))
      : utc(Number(c), Number(b), Number(a));
  }
  const dayFirst = /^(\d{1,2})\s+([A-Za-z]+),?\s+(\d{2,4})$/.exec(value);
  if (dayFirst) {
    const month = monthFromName(dayFirst[2]);
    return month === null ? null : utc(Number(dayFirst[3]), month, Number(dayFirst[1]));
  }
  const monthFirstText = /^([A-Za-z]+)\s+(\d{1,2}),?\s+(\d{2,4})$/.exec(value);
  if (monthFirstText) {
    const month = monthFromName(monthFirstText[1]);
    return month === null ? null : utc(Number(monthFirstText[3]), month, Number(monthFirstText[2]));
  }
  return null;
}

/** Compares two rows on one key. An unparseable number or date sorts before
 * every cell that parses (AC-44, AC-45) — and two unparseable cells tie, so
 * the next key (then the original order) decides. */
function compareOnKey(a: TableRow, b: TableRow, key: SortKey, collator: Intl.Collator, language: string): number {
  const left = sortCellText(a, key.columnIndex);
  const right = sortCellText(b, key.columnIndex);
  if (key.type === "text") {
    return collator.compare(left, right);
  }
  const parse = key.type === "number" ? parseSortNumber : (value: string) => parseSortDate(value, language);
  const leftValue = parse(left);
  const rightValue = parse(right);
  if (leftValue === null && rightValue === null) {
    return 0;
  }
  if (leftValue === null) {
    return -1;
  }
  if (rightValue === null) {
    return 1;
  }
  return leftValue < rightValue ? -1 : leftValue > rightValue ? 1 : 0;
}

/** The table with its sort range reordered by `keys`, in order.
 *
 * Stable (AC-47): ties fall through to the next key and then to the original
 * row order, so applying the same keys twice produces an `equalStructural`
 * table. `desc` reverses the comparison for its own key only (AC-46) — every
 * other key keeps its own direction.
 *
 * `language` is the document's (`DefaultStyles.language`, falling back to
 * `DEFAULT_LANGUAGE`): it drives `Intl.Collator` for text keys (AC-43) and
 * the day/month order for numeric dates. Returns a new Table; the input is
 * never mutated.
 */
export function sortTableRows(
  table: Table,
  keys: SortKey[],
  hasHeader: boolean,
  language: string = DEFAULT_LANGUAGE,
): Table {
  const start = sortRangeStart(table, hasHeader);
  const range = table.rows.slice(start);
  if (range.length < 2 || keys.length === 0) {
    return { ...table, rows: [...table.rows] };
  }
  const collator = new Intl.Collator(language, { numeric: true, sensitivity: "variant" });
  const decorated = range.map((row, index) => ({ row, index }));
  decorated.sort((a, b) => {
    for (const key of keys) {
      const order = compareOnKey(a.row, b.row, key, collator, language);
      if (order !== 0) {
        return key.direction === "desc" ? -order : order;
      }
    }
    // Every key tied: the original order breaks the tie, which is what makes
    // the sort stable regardless of the engine's own sort stability.
    return a.index - b.index;
  });
  return { ...table, rows: [...table.rows.slice(0, start), ...decorated.map((entry) => entry.row)] };
}
