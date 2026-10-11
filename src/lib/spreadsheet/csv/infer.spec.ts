// SPEC-126 "Type inference": the inference table in both directions — each
// field reads as what the table says, and what it reads as writes back
// (csvFieldText, in the same locale) as exactly the trimmed field — for both
// an en-US and a de-DE spelling of every numeric row, in both date systems.
// A field becomes typed only when that is lossless; everything else stays
// text, verbatim.
import { describe, expect, it } from "vitest";
import { isCellError, type CellValue } from "../model";
import { formatCodeOf } from "../numberFormat/builtin";
import { SPREADSHEET_LOCALES, type SpreadsheetLocaleTag } from "../numberFormat/locale";
import { inferCsvField, type InferredCsvField } from "./infer";
import { csvFieldText } from "./write";

const SERIAL_2026_10_08 = 46303; // 1900 system
const SYSTEM_1904_OFFSET = 1462;

type Stored = "text" | { value: CellValue; numberFormat?: string; builtinFormatId?: number };

/** One row of the table: a field as an en-US and a de-DE file spell it. */
interface Row {
  en: string;
  de?: string;
  stored: Stored;
  why: string;
}

const TABLE: Row[] = [
  { en: "007", stored: "text", why: "leading zero" },
  { en: "-0123", stored: "text", why: "leading zero, signed" },
  { en: "00", stored: "text", why: "leading zero" },
  { en: "0123.5", de: "0123,5", stored: "text", why: "leading zero with decimals" },
  { en: "1234567890123456", stored: "text", why: "16 digits; a double keeps 15" },
  { en: "1,234,567,890,123,456", de: "1.234.567.890.123.456", stored: "text", why: "16 digits, grouped" },
  { en: "1-2", stored: "text", why: "no format writes it back unchanged" },
  { en: "SEPT2", stored: "text", why: "no format writes it back unchanged" },
  { en: "1E5", stored: "text", why: "no format writes it back unchanged" },
  { en: ".5", de: ",5", stored: "text", why: "no format writes it back unchanged" },
  { en: "+3", stored: "text", why: "no format writes it back unchanged" },
  { en: "true", stored: "text", why: "only TRUE writes back as TRUE" },
  { en: "=SUM(A1:A2)", stored: "text", why: "CSV carries data, not formulas" },
  { en: "=1+1", stored: "text", why: "CSV carries data, not formulas" },
  { en: "'0042", stored: "text", why: "CSV has no quote-prefix convention" },
  { en: "Müller", stored: "text", why: "text" },
  { en: "  ", stored: "text", why: "spaces only" },
  { en: "1.50", de: "1,50", stored: { value: 1.5, numberFormat: "0.00" }, why: "fidelity format" },
  { en: "12%", stored: { value: 0.12, numberFormat: "0%" }, why: "parseInput's percent" },
  { en: "1.5%", de: "1,5%", stored: { value: 0.015, numberFormat: "0.0%" }, why: "fidelity percent" },
  { en: "2026-10-08", stored: { value: SERIAL_2026_10_08, numberFormat: "yyyy-mm-dd" }, why: "fidelity date" },
  {
    en: "10/8/2026",
    de: "08.10.2026",
    stored: { value: SERIAL_2026_10_08, numberFormat: "m/d/yyyy", builtinFormatId: 14 },
    why: "parseInput's date format renders back identically",
  },
  { en: "2026-10-08 14:30", stored: { value: SERIAL_2026_10_08 + 14.5 / 24, numberFormat: "yyyy-mm-dd hh:mm" }, why: "fidelity date-time" },
  {
    en: "2026-10-08 14:30:15",
    stored: { value: SERIAL_2026_10_08 + (14 * 3600 + 30 * 60 + 15) / 86400, numberFormat: "yyyy-mm-dd hh:mm:ss" },
    why: "fidelity date-time",
  },
  { en: "14:30", stored: { value: 14.5 / 24, numberFormat: "h:mm" }, why: "parseInput's time format" },
  { en: "09:05", stored: { value: (9 * 60 + 5) / 1440, numberFormat: "hh:mm" }, why: "fidelity time" },
  { en: "14:30:15", stored: { value: (14 * 3600 + 30 * 60 + 15) / 86400, numberFormat: "h:mm:ss" }, why: "parseInput's time format" },
  { en: "123456789012", stored: { value: 123456789012 }, why: "General here is the 15-digit form" },
  { en: "1,234.50", de: "1.234,50", stored: { value: 1234.5, numberFormat: "#,##0.00" }, why: "grouped" },
  { en: "1,234", de: "1.234", stored: { value: 1234, numberFormat: "#,##0" }, why: "grouped" },
  { en: "-7.25", de: "-7,25", stored: { value: -7.25 }, why: "General" },
  { en: "0", stored: { value: 0 }, why: "a lone zero is no leading zero" },
  { en: "0.5", de: "0,5", stored: { value: 0.5 }, why: "0.x is no leading zero" },
  { en: "$1,234.00", de: "1.234,00 €", stored: { value: 1234, numberFormat: "currency" }, why: "the locale's currency" },
  { en: "TRUE", stored: { value: true }, why: "boolean" },
  { en: "FALSE", stored: { value: false }, why: "boolean" },
  { en: "#N/A", stored: { value: { error: "#N/A" } }, why: "error" },
  { en: "#DIV/0!", stored: { value: { error: "#DIV/0!" } }, why: "error" },
];

/** The expected inference in `tag` and date system, the currency row
 * resolved to the locale's own currency code. */
function expectedOf(stored: Stored, tag: SpreadsheetLocaleTag, system: "1900" | "1904"): InferredCsvField | "text" {
  if (stored === "text") {
    return "text";
  }
  const { value, numberFormat, builtinFormatId } = stored;
  const isDate = numberFormat !== undefined && numberFormat !== "currency" && /[yd]|h:/.test(numberFormat);
  const shifted = typeof value === "number" && isDate && value >= 1 && system === "1904" ? value - SYSTEM_1904_OFFSET : value;
  const out: InferredCsvField = { value: shifted };
  if (numberFormat === "currency") {
    out.numberFormat = tag === "en-US" ? '"$"#,##0.00' : '#,##0.00" €"';
  } else if (numberFormat !== undefined) {
    out.numberFormat = numberFormat;
  }
  if (builtinFormatId !== undefined) {
    out.builtinFormatId = builtinFormatId;
  }
  return out;
}

function sameValue(actual: CellValue, expected: CellValue): void {
  if (typeof expected === "number" && typeof actual === "number") {
    expect(actual).toBeCloseTo(expected, 9);
  } else if (isCellError(expected)) {
    expect(actual).toEqual(expected);
  } else {
    expect(actual).toBe(expected);
  }
}

describe("inferCsvField — the table, both directions, both locales, both date systems", () => {
  for (const system of ["1900", "1904"] as const) {
    for (const tag of ["en-US", "de-DE"] as const) {
      const locale = SPREADSHEET_LOCALES[tag];
      for (const row of TABLE) {
        const field = tag === "en-US" ? row.en : (row.de ?? row.en);
        it(`${tag} ${system}: ${JSON.stringify(field)} — ${row.why}`, () => {
          const inferred = inferCsvField(field, locale, system);
          const expected = expectedOf(row.stored, tag, system);
          if (expected === "text") {
            // Read: text, verbatim (untrimmed). Write: itself.
            expect(inferred).toEqual({ value: field });
            expect(csvFieldText(inferred.value, undefined, locale, system)).toBe(field);
            return;
          }
          sameValue(inferred.value, expected.value);
          expect(inferred.numberFormat).toBe(expected.numberFormat);
          expect(inferred.builtinFormatId).toBe(expected.builtinFormatId);
          // Write: the typed value through its format is the field again.
          const code = formatCodeOf({ numberFormat: inferred.numberFormat, builtinFormatId: inferred.builtinFormatId }, locale);
          expect(csvFieldText(inferred.value, code, locale, system)).toBe(field.trim());
        });
      }
    }
  }
});

describe("inferCsvField — the rules", () => {
  const enUS = SPREADSHEET_LOCALES["en-US"];
  const deDE = SPREADSHEET_LOCALES["de-DE"];

  it("stores nothing for an empty field", () => {
    expect(inferCsvField("", enUS, "1900")).toEqual({ value: null });
  });

  it("compares the trimmed field, and keeps text untrimmed", () => {
    expect(inferCsvField(" 1.50 ", enUS, "1900")).toEqual({ value: 1.5, numberFormat: "0.00" });
    expect(inferCsvField(" Ada ", enUS, "1900")).toEqual({ value: " Ada " });
    expect(inferCsvField(" =1+1", enUS, "1900")).toEqual({ value: " =1+1" });
    expect(inferCsvField(" 007 ", enUS, "1900")).toEqual({ value: " 007 " });
  });

  it("reads a spelling only in the dialect's locale", () => {
    // 1,5 is a decimal in German and no number in English.
    expect(inferCsvField("1,5", deDE, "1900")).toEqual({ value: 1.5 });
    expect(inferCsvField("1,5", enUS, "1900")).toEqual({ value: "1,5" });
    // 8.10.2026 has no German format that writes it back (08.10.2026 does).
    expect(inferCsvField("8.10.2026", deDE, "1900")).toEqual({ value: "8.10.2026" });
  });

  it("keeps 15 digits typed and 16 as text", () => {
    expect(inferCsvField("123456789012345", enUS, "1900")).toEqual({ value: 123456789012345 });
    expect(inferCsvField("1234567890123456", enUS, "1900")).toEqual({ value: "1234567890123456" });
    expect(inferCsvField("-1234567890123456", enUS, "1900")).toEqual({ value: "-1234567890123456" });
  });
});
