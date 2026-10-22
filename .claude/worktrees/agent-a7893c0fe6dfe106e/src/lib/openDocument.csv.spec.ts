// SPEC-126 through the open funnel: a file that is not media and not a ZIP
// package is CSV if its name ends `.csv` and TSV if it ends `.tsv` — the
// second extension exception after Markdown's. Media sniffing, the unzip
// attempt and the flat-ODF check still run first (a ZIP named `.csv` opens by
// its own declarations), an OLE compound file named `.csv`/`.tsv` is refused
// as an old or protected workbook, the text-only entry point refuses a CSV,
// and the open surface grows by exactly the two extensions.
import { afterEach, describe, expect, it } from "vitest";
import {
  FlatOdfError,
  hasSupportedExtension,
  openAnyDocumentFile,
  openDocumentFile,
  SPREADSHEET_NOT_TEXT_MESSAGE,
  SpreadsheetOpenError,
  titleFromFileName,
  UNSUPPORTED_FILE_TYPE_MESSAGE,
  UnsupportedFileError,
} from "./openDocument";
import { OPENABLE_EXTENSIONS, kindOfFormat } from "./documentKinds";
import { buildPngFixtureBytes } from "./media/fixtures/build";
import { PREFERENCES_STORAGE_KEY } from "./preferences";
import { CSV_NOT_TEXT_MESSAGE } from "./spreadsheet/csv/messages";
import { CSV_SOURCE_PART } from "./spreadsheet/csv/parse";
import { LEGACY_OR_ENCRYPTED_WORKBOOK_MESSAGE } from "./spreadsheet/errors";
import { buildMiniBookXlsxBytes, buildSpec126Semicolon1252CsvBytes, buildSpec126UnicodeTsvBytes } from "./spreadsheet/fixtures/build";

const encoder = new TextEncoder();
const COMPOUND_FILE = Uint8Array.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1, 0, 0, 0, 0, 0, 0, 0, 0]);

function asFile(bytes: Uint8Array, name: string): File {
  return new File([bytes as BlobPart], name);
}

afterEach(() => localStorage.clear());

describe("detection", () => {
  it("opens a non-ZIP, non-media file named .csv as a csv workbook and .tsv as a tsv one", async () => {
    const csv = await openAnyDocumentFile(asFile(buildSpec126Semicolon1252CsvBytes(), "Kunden.csv"));
    expect(csv.kind).toBe("spreadsheet");
    if (csv.kind !== "spreadsheet") {
      return;
    }
    expect(csv.model.format).toBe("csv");
    expect(csv.model.title).toBe("Kunden");
    expect(csv.model.sheets[0]!.name).toBe("Kunden");
    expect(csv.model.opaque.map((part) => part.name)).toEqual([CSV_SOURCE_PART]);
    expect(csv.model.csv).toMatchObject({ delimiter: ";", encoding: "windows-1252", localeTag: "de-DE" });

    const tsv = await openAnyDocumentFile(asFile(buildSpec126UnicodeTsvBytes(), "Greetings.TSV"));
    expect(tsv.kind === "spreadsheet" && tsv.model.format).toBe("tsv");
    expect(kindOfFormat("csv")).toBe("spreadsheet");
    expect(kindOfFormat("tsv")).toBe("spreadsheet");
  });

  it("decides by name only for what is neither media nor a ZIP: the same bytes named .txt are no CSV", async () => {
    const bytes = encoder.encode("a,b\n1,2\n");
    expect((await openAnyDocumentFile(asFile(bytes, "data.csv"))).kind).toBe("spreadsheet");
    await expect(openAnyDocumentFile(asFile(bytes, "data.txt"))).rejects.toThrow();
  });

  it("sniffs media first: a PNG named .csv is a PNG", async () => {
    const opened = await openAnyDocumentFile(asFile(await buildPngFixtureBytes(), "chart.csv"));
    expect(opened.kind).toBe("media");
  });

  it("keeps a ZIP named .csv a ZIP: an .xlsx renamed .csv opens as the .xlsx it is", async () => {
    const opened = await openAnyDocumentFile(asFile(await buildMiniBookXlsxBytes(), "mini-book.csv"));
    expect(opened.kind).toBe("spreadsheet");
    if (opened.kind === "spreadsheet") {
      expect(opened.model.format).toBe("xlsx");
      expect(opened.model.csv).toBeUndefined();
      expect(opened.model.sheets.map((sheet) => sheet.name)).toEqual(["Summary", "Data", "Hidden"]);
    }
  });

  it("runs the flat-ODF check first", async () => {
    const flat = encoder.encode(
      `<?xml version="1.0" encoding="UTF-8"?>\n<office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" office:mimetype="application/vnd.oasis.opendocument.text" office:version="1.3"/>`,
    );
    await expect(openAnyDocumentFile(asFile(flat, "flat.csv"))).rejects.toBeInstanceOf(FlatOdfError);
  });

  it("refuses an OLE compound file named .csv or .tsv with LEGACY_OR_ENCRYPTED_WORKBOOK_MESSAGE before decoding it", async () => {
    for (const name of ["old.csv", "old.tsv", "OLD.CSV"]) {
      const open = openAnyDocumentFile(asFile(COMPOUND_FILE, name));
      await expect(open).rejects.toThrow(LEGACY_OR_ENCRYPTED_WORKBOOK_MESSAGE);
      await expect(openAnyDocumentFile(asFile(COMPOUND_FILE, name))).rejects.toBeInstanceOf(SpreadsheetOpenError);
    }
  });

  it("refuses binary data with a CSV name as CSV_NOT_TEXT_MESSAGE", async () => {
    await expect(openAnyDocumentFile(asFile(Uint8Array.from([0x61, 0x00, 0x62]), "binary.csv"))).rejects.toThrow(CSV_NOT_TEXT_MESSAGE);
  });

  it("reads numbers and dates in the user's spreadsheet locale unless the file's evidence says otherwise", async () => {
    localStorage.setItem(PREFERENCES_STORAGE_KEY, JSON.stringify({ spreadsheetLocale: "fr-FR" }));
    const tie = await openAnyDocumentFile(asFile(encoder.encode("a;b\nx;y\n"), "plain.csv"));
    expect(tie.kind === "spreadsheet" && tie.model.csv?.localeTag).toBe("fr-FR");
    const commas = await openAnyDocumentFile(asFile(buildSpec126Semicolon1252CsvBytes(), "commas.csv"));
    expect(commas.kind === "spreadsheet" && commas.model.csv?.localeTag).toBe("fr-FR");
    const points = await openAnyDocumentFile(asFile(encoder.encode("a,1.5\nb,2.25\n"), "points.csv"));
    expect(points.kind === "spreadsheet" && points.model.csv?.localeTag).toBe("en-US");
  });
});

describe("the text-only entry point", () => {
  it("rejects a CSV or TSV with SPREADSHEET_NOT_TEXT_MESSAGE rather than misparsing it", async () => {
    for (const [bytes, name] of [
      [buildSpec126Semicolon1252CsvBytes(), "data.csv"],
      [buildSpec126UnicodeTsvBytes(), "data.tsv"],
    ] as const) {
      await expect(openDocumentFile(asFile(bytes, name))).rejects.toThrow(SPREADSHEET_NOT_TEXT_MESSAGE);
      await expect(openDocumentFile(asFile(bytes, name))).rejects.toBeInstanceOf(UnsupportedFileError);
    }
  });
});

describe("the open surface", () => {
  it("accepts .csv and .tsv at the drop check, right after the workbook package extensions", () => {
    for (const name of ["a.csv", "a.tsv", "A.CSV", "b.Tsv"]) {
      expect(hasSupportedExtension(name)).toBe(true);
    }
    for (const name of ["a.txt", "a.tab", "a.psv", "a.prn", "a.csv.txt"]) {
      expect(hasSupportedExtension(name)).toBe(false);
    }
    const ots = OPENABLE_EXTENSIONS.indexOf("ots");
    expect(OPENABLE_EXTENSIONS.slice(ots, ots + 3)).toEqual(["ots", "csv", "tsv"]);
    expect(UNSUPPORTED_FILE_TYPE_MESSAGE).toBe("We can only open .docx, .odt, .md, .pptx, .odp, .key, .xlsx, .ods, .csv, .tsv, .pdf and image files");
  });

  it("strips .csv and .tsv from the title", () => {
    expect(titleFromFileName("Bank statement.csv")).toBe("Bank statement");
    expect(titleFromFileName("export.TSV")).toBe("export");
    expect(titleFromFileName("notes.csv.txt")).toBe("notes.csv.txt");
  });
});
