// SPEC-126 "Decoding", "Sniffing" and "Locale inference": the decoding table
// and NUL refusal, Excel's `sep=` line, the delimiter table and its tie order
// (the decimal-commas-in-a-`;`-file case included), line endings, quoteAll,
// width, and the locale table — all six user locales × comma / point / tie
// evidence × date evidence.
import { describe, expect, it } from "vitest";
import type { CsvDelimiter } from "../model";
import { SPREADSHEET_LOCALE_TAGS, type SpreadsheetLocaleTag } from "../numberFormat/locale";
import { buildSpec126MacPipeCsvBytes, buildSpec126Semicolon1252CsvBytes, buildSpec126SepCsvBytes, buildSpec126UnicodeTsvBytes } from "../fixtures/build";
import { CSV_NOT_TEXT_MESSAGE } from "./messages";
import { CsvOpenError, readCsvBytes } from "./parse";
import { decodeCsvBytes, DELIMITER_CANDIDATES, inferLocaleTag, SNIFF_RECORDS, sepLineOf, sniffDelimiter, sniffSample } from "./sniff";

const encoder = new TextEncoder();

function bytes(...values: number[]): Uint8Array {
  return Uint8Array.from(values);
}

function dialectOf(text: string | Uint8Array, format: "csv" | "tsv" = "csv", userLocaleTag: SpreadsheetLocaleTag = "en-US") {
  const source = typeof text === "string" ? encoder.encode(text) : text;
  return readCsvBytes(source, format, "t", { userLocaleTag }).csv!;
}

describe("decoding", () => {
  it("lets a UTF-8, UTF-16 LE or UTF-16 BE BOM decide the encoding, and strips it", () => {
    expect(decodeCsvBytes(bytes(0xef, 0xbb, 0xbf, 0x61, 0xc3, 0xa9))).toEqual({ text: "aé", encoding: "utf-8", bom: true });
    expect(decodeCsvBytes(bytes(0xff, 0xfe, 0x61, 0x00, 0xe9, 0x00))).toEqual({ text: "aé", encoding: "utf-16le", bom: true });
    expect(decodeCsvBytes(bytes(0xfe, 0xff, 0x00, 0x61, 0x00, 0xe9))).toEqual({ text: "aé", encoding: "utf-16be", bom: true });
  });

  it("decodes a BOM-less buffer as fatal UTF-8 when it is valid — a pure-ASCII file is UTF-8", () => {
    expect(decodeCsvBytes(encoder.encode("a,b\n"))).toEqual({ text: "a,b\n", encoding: "utf-8", bom: false });
    expect(decodeCsvBytes(encoder.encode("Müller"))).toEqual({ text: "Müller", encoding: "utf-8", bom: false });
  });

  it("falls back to Windows-1252 when fatal UTF-8 throws", () => {
    // "Müller €" in Windows-1252: ü is FC, € is 80 — neither valid UTF-8.
    expect(decodeCsvBytes(bytes(0x4d, 0xfc, 0x6c, 0x6c, 0x65, 0x72, 0x20, 0x80))).toEqual({
      text: "Müller €",
      encoding: "windows-1252",
      bom: false,
    });
  });

  it("decodes with an explicitly chosen encoding, stripping its BOM only when present", () => {
    expect(decodeCsvBytes(bytes(0xef, 0xbb, 0xbf, 0x61), "utf-8")).toEqual({ text: "a", encoding: "utf-8", bom: true });
    expect(decodeCsvBytes(bytes(0x61, 0xe9), "windows-1252")).toEqual({ text: "aé", encoding: "windows-1252", bom: false });
    // A UTF-8 BOM read as Windows-1252 is three characters of data.
    expect(decodeCsvBytes(bytes(0xef, 0xbb, 0xbf, 0x61), "windows-1252").text).toBe("ï»¿a");
    expect(decodeCsvBytes(bytes(0x61, 0x00), "utf-16le")).toEqual({ text: "a", encoding: "utf-16le", bom: false });
  });

  it("refuses text holding a U+0000 with CSV_NOT_TEXT_MESSAGE", () => {
    expect(CSV_NOT_TEXT_MESSAGE).toBe("This file doesn't contain text we can read as CSV");
    // UTF-16 without its BOM decodes as UTF-8 full of NULs: binary to us.
    const read = () => readCsvBytes(bytes(0x61, 0x00, 0x2c, 0x00, 0x62, 0x00), "csv", "t", { userLocaleTag: "en-US" });
    expect(read).toThrow(CSV_NOT_TEXT_MESSAGE);
    expect(read).toThrow(CsvOpenError);
    expect(() => decodeCsvBytes(bytes(0x50, 0x4b, 0x00, 0x00))).toThrow(CSV_NOT_TEXT_MESSAGE);
  });

  it("records each fixture's encoding and BOM", () => {
    expect(dialectOf(buildSpec126Semicolon1252CsvBytes())).toMatchObject({ encoding: "windows-1252", bom: false });
    expect(dialectOf(buildSpec126UnicodeTsvBytes(), "tsv")).toMatchObject({ encoding: "utf-16le", bom: true });
    expect(dialectOf(buildSpec126MacPipeCsvBytes())).toMatchObject({ encoding: "windows-1252", bom: false });
  });
});

describe("the sep= line", () => {
  it("decides the delimiter from a first line of exactly sep=X, X one of , ; tab |", () => {
    for (const delimiter of [",", ";", "\t", "|"] as const) {
      expect(sepLineOf(`sep=${delimiter}\r\na`)).toEqual({ delimiter, length: 7, lineEnding: "\r\n" });
    }
    expect(sepLineOf("sep=;\na")).toEqual({ delimiter: ";", length: 6, lineEnding: "\n" });
    expect(sepLineOf("sep=;")).toEqual({ delimiter: ";", length: 5 });
    for (const text of ["sep=:\na", "Sep=;\na", "sep=;x\na", " sep=;\na", "sep=\na", "a\nsep=;"]) {
      expect(sepLineOf(text), text).toBeNull();
    }
  });

  it("is not a data row, is recorded as sepLine and its terminator as the line ending", () => {
    const model = readCsvBytes(buildSpec126SepCsvBytes(), "csv", "t", { userLocaleTag: "en-US" });
    expect(model.csv).toMatchObject({ delimiter: ";", sepLine: true, lineEnding: "\r\n", width: 3 });
    expect(model.sheets[0]!.rows[0]!.cells.map((cell) => cell.value)).toEqual(["Item", "Qty", "Price"]);
    expect(model.sheets[0]!.rows).toHaveLength(3);
    // sep= wins over what the records would sniff as.
    expect(dialectOf("sep=|\na,b\nc,d\n")).toMatchObject({ delimiter: "|", sepLine: true });
    expect(dialectOf("a,b\n").sepLine).toBeUndefined();
  });
});

describe("the delimiter", () => {
  it("tries tab, pipe, semicolon, comma in that order", () => {
    expect(DELIMITER_CANDIDATES).toEqual(["\t", "|", ";", ","]);
    expect(SNIFF_RECORDS).toBe(50);
  });

  const table: { text: string; csv: CsvDelimiter; tsv?: CsvDelimiter; why: string }[] = [
    { text: "a,b,c\n1,2,3\n", csv: ",", why: "commas only" },
    { text: "a;b;c\n1;2;3\n", csv: ";", why: "semicolons only" },
    { text: "a\tb\n1\t2\n", csv: "\t", why: "tabs only" },
    { text: "a|b\n1|2\n", csv: "|", why: "pipes only" },
    { text: "Name;Amount\nA;1,5\nB;2,25\nC;3,75\n", csv: ";", why: "decimal commas inside a ; file split perfectly on , too — the tie goes to ;" },
    { text: "a\tb|c;d,e\nf\tg|h;i,j\n", csv: "\t", why: "a four-way tie goes to tab" },
    { text: "a|b;c,d\ne|f;g,h\n", csv: "|", why: "pipe before semicolon and comma" },
    { text: "a;b,c\nd;e,f\n", csv: ";", why: "semicolon before comma" },
    { text: "a,b;c\nd,e\nf,g\n", csv: ",", why: "a ; inside one field is not consistent: its modal count is 1" },
    { text: "a;b,c\nd;e,f\ng,h\ni,j\n", csv: ",", why: "the highest consistency wins over the tie order (, 4/4 over ; 2/4)" },
    { text: "hello\nworld\n", csv: ",", tsv: "\t", why: "no candidate yields two fields: the extension's own delimiter" },
    { text: "", csv: ",", tsv: "\t", why: "an empty file" },
  ];
  for (const row of table) {
    it(`${JSON.stringify(row.text)} → ${JSON.stringify(row.csv)} (${row.why})`, () => {
      expect(sniffDelimiter(row.text, "csv")).toBe(row.csv);
      expect(sniffDelimiter(row.text, "tsv")).toBe(row.tsv ?? row.csv);
    });
  }

  it("compares consistency as exact fractions", () => {
    // ; : 2 of 3 records share the modal count; , : 2 of 3 as well — equal
    // fractions tie, and the tie goes to ;.
    expect(sniffDelimiter("a;b,c\nd;e\nf,g\n", "csv")).toBe(";");
    // 50/50 on , beats 48/50 on ; .
    const lines = [...Array.from({ length: 48 }, () => "a;b,c"), "a;b;c,d", "x,y"];
    expect(sniffDelimiter(lines.join("\n"), "csv")).toBe(",");
  });

  it("reads only the first SNIFF_RECORDS non-blank records", () => {
    const head = Array.from({ length: SNIFF_RECORDS }, () => "a;b").join("\n\n");
    const tail = Array.from({ length: 200 }, () => "a,b,c,d").join("\n");
    expect(sniffDelimiter(`${head}\n${tail}`, "csv")).toBe(";");
  });

  it("reads the German fixture as semicolon-separated", () => {
    expect(dialectOf(buildSpec126Semicolon1252CsvBytes()).delimiter).toBe(";");
    expect(dialectOf(buildSpec126MacPipeCsvBytes()).delimiter).toBe("|");
    expect(dialectOf(buildSpec126UnicodeTsvBytes(), "tsv").delimiter).toBe("\t");
  });
});

describe("line endings, quoteAll and width", () => {
  it("records the first terminator outside quotes and whether the last record was terminated", () => {
    expect(dialectOf("a\nb\r\n")).toMatchObject({ lineEnding: "\n", finalLineEnding: true });
    expect(dialectOf("a\r\nb")).toMatchObject({ lineEnding: "\r\n", finalLineEnding: false });
    expect(dialectOf(buildSpec126MacPipeCsvBytes())).toMatchObject({ lineEnding: "\r", finalLineEnding: true });
    expect(dialectOf("a,b")).toMatchObject({ lineEnding: "\r\n", finalLineEnding: false });
  });

  it("sets quoteAll when every non-empty field in the sample was quoted", () => {
    expect(sniffSample('"a","b"\n"c",""\n', ",").quoteAll).toBe(true);
    expect(sniffSample('"a",,"b"\n', ",").quoteAll).toBe(true);
    expect(sniffSample('"a",b\n', ",").quoteAll).toBe(false);
    expect(sniffSample(",\n\n", ",").quoteAll).toBe(false);
    expect(dialectOf('"id","name"\n"1","Ada"').quoteAll).toBe(true);
  });

  it("collects the sample's fields trimmed, blank lines skipped", () => {
    expect(sniffSample(" 1,5 ; x\n\n2;y\n", ";").fields).toEqual(["1,5", "x", "2", "y"]);
  });

  it("sets width only when every non-blank record in the whole file agrees", () => {
    expect(dialectOf("a,b\n\nc,d\n").width).toBe(2);
    expect(dialectOf("a,b\nc,d,e\n").width).toBeUndefined();
    const late = `${Array.from({ length: 60 }, () => "a,b").join("\n")}\nc\n`;
    expect(dialectOf(late).width).toBeUndefined();
  });
});

/** The date evidence's effect: en-US ↔ en-GB only. */
function swapped(tag: SpreadsheetLocaleTag, dates: "none" | "day" | "month"): SpreadsheetLocaleTag {
  if (dates === "day" && tag === "en-US") {
    return "en-GB";
  }
  if (dates === "month" && tag === "en-GB") {
    return "en-US";
  }
  return tag;
}

describe("locale inference", () => {
  const COMMA = ["1,5", "2,25", "-1.234,56"];
  const POINT = ["1.5", "2.25", "+1,234.56"];
  const TIE = ["1,5", "1.5"];
  const DAY_FIRST = ["31/12/2025", "1/2/2025"];
  const MONTH_FIRST = ["12/31/2025", "2/1/2025"];

  const expected: Record<SpreadsheetLocaleTag, Record<"comma" | "point" | "tie", SpreadsheetLocaleTag>> = {
    "en-US": { comma: "de-DE", point: "en-US", tie: "en-US" },
    "en-GB": { comma: "de-DE", point: "en-GB", tie: "en-GB" },
    "de-DE": { comma: "de-DE", point: "en-US", tie: "de-DE" },
    "fr-FR": { comma: "fr-FR", point: "en-US", tie: "fr-FR" },
    "es-ES": { comma: "es-ES", point: "en-US", tie: "es-ES" },
    "cs-CZ": { comma: "cs-CZ", point: "en-US", tie: "cs-CZ" },
  };

  for (const user of SPREADSHEET_LOCALE_TAGS) {
    for (const [kind, fields] of [
      ["comma", COMMA],
      ["point", POINT],
      ["tie", TIE],
    ] as const) {
      for (const [dates, dateFields] of [
        ["none", []],
        ["day", DAY_FIRST],
        ["month", MONTH_FIRST],
      ] as const) {
        const want = swapped(expected[user][kind], dates);
        it(`a ${user} user, ${kind} evidence, ${dates === "none" ? "no" : `${dates}-first`} dates → ${want}`, () => {
          expect(inferLocaleTag([...fields, ...dateFields, "text"], user)).toBe(want);
        });
      }
    }
  }

  it("counts no evidence as a tie: the user's locale", () => {
    for (const user of SPREADSHEET_LOCALE_TAGS) {
      expect(inferLocaleTag(["a", "b"], user)).toBe(user);
    }
  });

  it("does not count the ambiguous 1,234 or 1.234 as evidence either way", () => {
    expect(inferLocaleTag(["1,234", "-12,345", "1,5"], "en-US")).toBe("de-DE");
    expect(inferLocaleTag(["1.234", "+12.345", "1.5"], "de-DE")).toBe("en-US");
    expect(inferLocaleTag(["1,234", "1.234"], "fr-FR")).toBe("fr-FR");
  });

  it("swaps nothing when both components exceed 12 somewhere", () => {
    expect(inferLocaleTag(["31/12/2025", "12/31/2025"], "en-US")).toBe("en-US");
    expect(inferLocaleTag(["31/12/2025", "12/31/2025"], "en-GB")).toBe("en-GB");
  });

  it("reads a ; file with decimal commas as German for an English user, its dialect recording it", () => {
    expect(dialectOf(buildSpec126Semicolon1252CsvBytes(), "csv", "en-US").localeTag).toBe("de-DE");
    expect(dialectOf(buildSpec126Semicolon1252CsvBytes(), "csv", "en-GB").localeTag).toBe("de-DE");
    expect(dialectOf(buildSpec126Semicolon1252CsvBytes(), "csv", "fr-FR").localeTag).toBe("fr-FR");
  });

  it("takes an explicitly chosen locale without sniffing", () => {
    const model = readCsvBytes(buildSpec126Semicolon1252CsvBytes(), "csv", "t", { userLocaleTag: "en-US", localeTag: "en-US" });
    expect(model.csv?.localeTag).toBe("en-US");
  });
});
