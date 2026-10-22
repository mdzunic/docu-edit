// Sniffing a delimited-text file (SPEC-126 "Decoding", "Sniffing", "Locale
// inference"): the encoding from the bytes, then — over the decoded text —
// Excel's `sep=` line, the delimiter, the sample's quoting and the locale
// the numbers and dates were written in. Every facet the caller names
// explicitly (the import dialog) is taken as given instead.
//
// - Decoding. A UTF-8, UTF-16 LE or UTF-16 BE BOM decides the encoding and
//   is stripped; otherwise the whole buffer is decoded as UTF-8 with
//   `fatal: true` and, if that throws, as Windows-1252 (so a pure-ASCII file
//   is UTF-8). An explicitly chosen encoding strips its own BOM when present.
//   A U+0000 anywhere in the text is binary data with a CSV name.
// - `sep=`. A first line (up to the first CR/LF) of exactly `sep=` and one
//   candidate character decides the delimiter and is not a data row.
// - Delimiter. For each candidate in the order tab, pipe, semicolon, comma —
//   how rarely each appears inside ordinary field content — the first
//   SNIFF_RECORDS non-blank records are parsed; a candidate is eligible when
//   its modal field count is at least 2, and its consistency is the share of
//   records with that count, compared exactly as fractions. The eligible
//   candidate with the highest consistency wins, ties to the earlier one: a
//   `;` file with decimal commas also splits perfectly on `,`, but a `,` file
//   almost never has the same number of semicolons on every line. No
//   eligible candidate → the extension's own delimiter.
// - Locale. Over the sample's trimmed fields, decimal-comma evidence
//   outvoting decimal-point evidence picks a comma-decimal locale (the
//   user's if it is one, else de-DE), the reverse a point-decimal one (the
//   user's if it is one, else en-US), a tie the user's; then day-first or
//   month-first date evidence swaps en-US ↔ en-GB. The comma-decimal locales
//   are all day-first, so nothing swaps them.
import type { CsvDelimiter, CsvDialect, CsvEncoding } from "../model";
import { SPREADSHEET_LOCALES, type SpreadsheetLocaleTag } from "../numberFormat/locale";
import { CSV_NOT_TEXT_MESSAGE } from "./messages";
import { CsvOpenError, isBlankRecord, parseCsvText } from "./parse";

/** How many non-blank records the delimiter, quoting and locale sniffing
 * read. */
export const SNIFF_RECORDS = 50;

/** The delimiter candidates in tie order. */
export const DELIMITER_CANDIDATES: readonly CsvDelimiter[] = ["\t", "|", ";", ","];

const BOMS: Readonly<Partial<Record<CsvEncoding, readonly number[]>>> = {
  "utf-8": [0xef, 0xbb, 0xbf],
  "utf-16le": [0xff, 0xfe],
  "utf-16be": [0xfe, 0xff],
};

/** The BOM bytes of an encoding (none for Windows-1252). */
export function bomBytes(encoding: CsvEncoding): readonly number[] {
  return BOMS[encoding] ?? [];
}

function startsWithBytes(bytes: Uint8Array, prefix: readonly number[]): boolean {
  return prefix.length > 0 && bytes.length >= prefix.length && prefix.every((byte, index) => bytes[index] === byte);
}

export interface DecodedCsv {
  text: string;
  encoding: CsvEncoding;
  /** The bytes began with the encoding's BOM (stripped from `text`). */
  bom: boolean;
}

function decodeAs(bytes: Uint8Array, encoding: CsvEncoding, fatal = false): string {
  return new TextDecoder(encoding, { fatal, ignoreBOM: true }).decode(bytes);
}

/** Decodes a file's bytes per "Decoding" (see the module header); throws
 * CsvOpenError with CSV_NOT_TEXT_MESSAGE when the text holds a U+0000. */
export function decodeCsvBytes(bytes: Uint8Array, encoding?: CsvEncoding): DecodedCsv {
  let decoded: DecodedCsv | undefined;
  if (encoding !== undefined) {
    const bom = startsWithBytes(bytes, bomBytes(encoding));
    decoded = { text: decodeAs(bom ? bytes.subarray(bomBytes(encoding).length) : bytes, encoding), encoding, bom };
  } else {
    for (const candidate of ["utf-8", "utf-16le", "utf-16be"] as const) {
      if (startsWithBytes(bytes, bomBytes(candidate))) {
        decoded = { text: decodeAs(bytes.subarray(bomBytes(candidate).length), candidate), encoding: candidate, bom: true };
        break;
      }
    }
    if (!decoded) {
      try {
        decoded = { text: decodeAs(bytes, "utf-8", true), encoding: "utf-8", bom: false };
      } catch {
        decoded = { text: decodeAs(bytes, "windows-1252"), encoding: "windows-1252", bom: false };
      }
    }
  }
  if (decoded.text.includes("\0")) {
    throw new CsvOpenError(CSV_NOT_TEXT_MESSAGE);
  }
  return decoded;
}

const SEP_LINE = /^sep=([,;\t|])(\r\n|\n|\r|$)/;

/** Excel's `sep=X` first line: its delimiter, its length with its
 * terminator, and that terminator (absent when the file is only that
 * line). Null when the first line is anything else. */
export function sepLineOf(text: string): { delimiter: CsvDelimiter; length: number; lineEnding?: CsvDialect["lineEnding"] } | null {
  const match = SEP_LINE.exec(text);
  if (!match) {
    return null;
  }
  const sep: { delimiter: CsvDelimiter; length: number; lineEnding?: CsvDialect["lineEnding"] } = {
    delimiter: match[1] as CsvDelimiter,
    length: match[0].length,
  };
  if (match[2]) {
    sep.lineEnding = match[2] as CsvDialect["lineEnding"];
  }
  return sep;
}

/** The field counts of the first SNIFF_RECORDS non-blank records. */
function sampleCounts(text: string, delimiter: CsvDelimiter): number[] {
  const counts: number[] = [];
  parseCsvText(text, delimiter, (fields) => {
    if (isBlankRecord(fields)) {
      return undefined;
    }
    counts.push(fields.length);
    return counts.length >= SNIFF_RECORDS ? "stop" : undefined;
  });
  return counts;
}

/** The most frequent count and how often it occurs (a tie goes to the
 * larger count). */
function modal(counts: readonly number[]): { count: number; frequency: number } {
  const frequencies = new Map<number, number>();
  for (const count of counts) {
    frequencies.set(count, (frequencies.get(count) ?? 0) + 1);
  }
  let best = { count: 0, frequency: 0 };
  for (const [count, frequency] of frequencies) {
    if (frequency > best.frequency || (frequency === best.frequency && count > best.count)) {
      best = { count, frequency };
    }
  }
  return best;
}

/** The delimiter of text with no `sep=` line (see the module header). */
export function sniffDelimiter(text: string, format: "csv" | "tsv"): CsvDelimiter {
  let best: { delimiter: CsvDelimiter; matching: number; total: number } | null = null;
  for (const delimiter of DELIMITER_CANDIDATES) {
    const counts = sampleCounts(text, delimiter);
    if (counts.length === 0) {
      continue;
    }
    const { count, frequency } = modal(counts);
    if (count < 2) {
      continue;
    }
    // frequency / counts.length > best.matching / best.total, as integers.
    if (!best || frequency * best.total > best.matching * counts.length) {
      best = { delimiter, matching: frequency, total: counts.length };
    }
  }
  return best?.delimiter ?? (format === "tsv" ? "\t" : ",");
}

/** The sniff sample under the chosen delimiter: every field of the first
 * SNIFF_RECORDS non-blank records, trimmed, and whether every non-empty
 * field in it was quoted (with at least one such field). */
export function sniffSample(text: string, delimiter: CsvDelimiter): { fields: string[]; quoteAll: boolean } {
  const fields: string[] = [];
  let records = 0;
  let nonEmpty = 0;
  let allQuoted = true;
  parseCsvText(text, delimiter, (record, quoted) => {
    if (isBlankRecord(record)) {
      return undefined;
    }
    record.forEach((field, index) => {
      if (field !== "") {
        nonEmpty += 1;
        allQuoted &&= quoted[index] === true;
      }
      fields.push(field.trim());
    });
    records += 1;
    return records >= SNIFF_RECORDS ? "stop" : undefined;
  });
  return { fields, quoteAll: nonEmpty > 0 && allQuoted };
}

const COMMA_DECIMAL = [/^[+-]?\d+,\d+$/, /^[+-]?\d{1,3}(\.\d{3})+,\d+$/];
const COMMA_AMBIGUOUS = /^[+-]?\d{1,3},\d{3}$/;
const POINT_DECIMAL = [/^[+-]?\d+\.\d+$/, /^[+-]?\d{1,3}(,\d{3})+\.\d+$/];
const POINT_AMBIGUOUS = /^[+-]?\d{1,3}\.\d{3}$/;
const NUMERIC_DATE = /^(\d{1,2})[/.-](\d{1,2})[/.-]\d{2,4}$/;

/** The locale the sample's numbers and dates were written in (see the
 * module header). `fields` are already trimmed. */
export function inferLocaleTag(fields: readonly string[], userLocaleTag: SpreadsheetLocaleTag): SpreadsheetLocaleTag {
  let comma = 0;
  let point = 0;
  for (const field of fields) {
    if (!COMMA_AMBIGUOUS.test(field) && COMMA_DECIMAL.some((pattern) => pattern.test(field))) {
      comma += 1;
    }
    if (!POINT_AMBIGUOUS.test(field) && POINT_DECIMAL.some((pattern) => pattern.test(field))) {
      point += 1;
    }
  }
  const userDecimal = SPREADSHEET_LOCALES[userLocaleTag].decimal;
  let tag: SpreadsheetLocaleTag = userLocaleTag;
  if (comma > point) {
    tag = userDecimal === "," ? userLocaleTag : "de-DE";
  } else if (point > comma) {
    tag = userDecimal === "." ? userLocaleTag : "en-US";
  }
  let firstOver = false;
  let secondOver = false;
  for (const field of fields) {
    const match = NUMERIC_DATE.exec(field);
    if (match) {
      firstOver ||= Number(match[1]) > 12;
      secondOver ||= Number(match[2]) > 12;
    }
  }
  if (tag === "en-US" && firstOver && !secondOver) {
    return "en-GB";
  }
  if (tag === "en-GB" && secondOver && !firstOver) {
    return "en-US";
  }
  return tag;
}
