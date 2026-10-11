// SPEC-116 through the open funnel: workbooks are detected inside the ZIP
// funnel from the package's own declarations — never the extension — as
// "xlsx" (with its package variant) or "ods"; every workbook the product
// cannot read is refused with the message that names the way out (.xlsb, an
// old or password-protected workbook, a Numbers document — told apart from
// Keynote and Pages — and flat ODF); the text-only entry point refuses a
// workbook rather than misparsing it; and the open surface's strings grow.
import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";
import {
  FILE_TOO_LARGE_MESSAGE,
  hasSupportedExtension,
  LegacyDocError,
  MAX_FILE_SIZE_BYTES,
  openAnyDocumentFile,
  openDocumentFile,
  SPREADSHEET_NOT_TEXT_MESSAGE,
  SpreadsheetOpenError,
  titleFromFileName,
  UNSUPPORTED_FILE_TYPE_MESSAGE,
  UNSUPPORTED_ODF_MESSAGE,
  UnsupportedFileError,
} from "./openDocument";
import { FLAT_ODF_MESSAGE } from "./odf/reader";
import { LEGACY_DOC_MESSAGE } from "./docx/reader";
import { NOT_KEYNOTE_MESSAGE } from "./presentation/keynote/detect";
import { buildKeynoteFromArchives, message, varintField } from "./presentation/keynote/fixtures";
import {
  LEGACY_OR_ENCRYPTED_WORKBOOK_MESSAGE,
  NUMBERS_UNSUPPORTED_MESSAGE,
  WORKBOOK_TOO_LARGE_MESSAGE,
  XLSB_UNSUPPORTED_MESSAGE,
} from "./spreadsheet/errors";
import {
  buildMiniBookOdsBytes,
  buildMiniBookXlsxBytes,
  buildSpecMacrosXlsmBytes,
  buildSpecTemplateXltxBytes,
} from "./spreadsheet/fixtures/build";
import { unzip } from "./core/zip";
import { zip } from "./core/zipWriter";
import { spreadsheetFormatOf, detectPackage, SPREADSHEETML_MAIN_CONTENT_TYPES, ODS_MIMETYPES, XLSB_MAIN_CONTENT_TYPE } from "./core/packageDetect";

const e2eFixturesDir = join(dirname(fileURLToPath(import.meta.url)), "..", "..", "e2e", "fixtures");
const encoder = new TextEncoder();

function asFile(bytes: Uint8Array, name: string): File {
  return new File([bytes as BlobPart], name);
}

/** The bytes of a package with one entry swapped for `content`. */
async function repackaged(bytes: Uint8Array, name: string, content: string): Promise<Uint8Array> {
  const parts = await unzip(bytes);
  parts.set(name, encoder.encode(content));
  return zip([...parts].map(([entry, data]) => ({ name: entry, data, store: entry === "mimetype" })));
}

/** An OPC package whose main part declares `contentType`. */
async function opcPackage(contentType: string): Promise<Uint8Array> {
  return zip([
    {
      name: "[Content_Types].xml",
      data: encoder.encode(
        `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
          `<Override PartName="/xl/workbook.bin" ContentType="${contentType}"/></Types>`,
      ),
    },
    {
      name: "_rels/.rels",
      data: encoder.encode(
        `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
          `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.bin"/></Relationships>`,
      ),
    },
    { name: "xl/workbook.bin", data: Uint8Array.from([0x83, 0x01, 0x00]) },
  ]);
}

/** The Compound File Binary signature followed by filler. */
const COMPOUND_FILE = Uint8Array.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1, ...Array.from({ length: 504 }, () => 0)]);

describe("detection from the package's own declarations (never the extension)", () => {
  it("opens every SpreadsheetML variant as \"xlsx\" with its package variant", async () => {
    const cases = [
      { bytes: await buildMiniBookXlsxBytes(), variant: undefined },
      { bytes: await buildSpecMacrosXlsmBytes(), variant: "macroEnabled" },
      { bytes: await buildSpecTemplateXltxBytes(), variant: "template" },
      {
        bytes: await repackaged(
          await buildSpecMacrosXlsmBytes(),
          "[Content_Types].xml",
          new TextDecoder()
            .decode((await unzip(await buildSpecMacrosXlsmBytes())).get("[Content_Types].xml"))
            .replace(SPREADSHEETML_MAIN_CONTENT_TYPES.macroEnabled, SPREADSHEETML_MAIN_CONTENT_TYPES.macroEnabledTemplate),
        ),
        variant: "macroEnabledTemplate",
      },
    ];
    for (const { bytes, variant } of cases) {
      // A misleading name proves the extension decides nothing.
      const opened = await openAnyDocumentFile(asFile(bytes, "report.docx"));
      expect(opened.kind).toBe("spreadsheet");
      if (opened.kind === "spreadsheet") {
        expect(opened.model.format).toBe("xlsx");
        expect(opened.model.packageVariant).toBe(variant);
        expect(opened.model.title).toBe("report");
      }
    }
  });

  it("opens an ODF spreadsheet or template mimetype as \"ods\"", async () => {
    const ods = await openAnyDocumentFile(asFile(await buildMiniBookOdsBytes(), "notes.odt"));
    expect(ods.kind).toBe("spreadsheet");
    if (ods.kind === "spreadsheet") {
      expect(ods.model.format).toBe("ods");
      expect(ods.model.packageVariant).toBeUndefined();
      expect(ods.model.sheets.map((sheet) => sheet.name)).toEqual(["Archive", "Inventory", "Totals"]);
    }
    const ots = await repackaged(await buildMiniBookOdsBytes(), "mimetype", ODS_MIMETYPES.template);
    const template = await openAnyDocumentFile(asFile(ots, "form.ots"));
    expect(template.kind === "spreadsheet" && template.model.packageVariant).toBe("template");
  });

  it("spreadsheetFormatOf classifies by declarations", async () => {
    expect(spreadsheetFormatOf(detectPackage(await unzip(await buildMiniBookXlsxBytes())))).toEqual({ format: "xlsx" });
    expect(spreadsheetFormatOf(detectPackage(await unzip(await opcPackage(XLSB_MAIN_CONTENT_TYPE))))).toBe("xlsb");
    expect(spreadsheetFormatOf({ container: "odf", mimetype: ODS_MIMETYPES.spreadsheet })).toEqual({ format: "ods" });
    expect(spreadsheetFormatOf({ container: "odf", mimetype: "application/vnd.oasis.opendocument.text" })).toBeNull();
    expect(spreadsheetFormatOf({ container: "zip" })).toBeNull();
    expect(
      spreadsheetFormatOf({
        container: "opc",
        mainPartPath: "ppt/presentation.xml",
        mainContentType: "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml",
      }),
    ).toBeNull();
  });
});

describe("refusals that name the way out", () => {
  it("refuses an .xlsb package with XLSB_UNSUPPORTED_MESSAGE", async () => {
    const open = openAnyDocumentFile(asFile(await opcPackage(XLSB_MAIN_CONTENT_TYPE), "book.xlsb"));
    await expect(open).rejects.toThrow(XLSB_UNSUPPORTED_MESSAGE);
    await expect(openAnyDocumentFile(asFile(await opcPackage(XLSB_MAIN_CONTENT_TYPE), "renamed.xlsx"))).rejects.toBeInstanceOf(
      SpreadsheetOpenError,
    );
  });

  it("refuses an OLE compound file named .xls, .xlsx or .xlsm as an old or protected workbook", async () => {
    for (const name of ["old.xls", "Locked.XLSX", "macros.xlsm"]) {
      const open = openAnyDocumentFile(asFile(COMPOUND_FILE, name));
      await expect(open).rejects.toThrow(LEGACY_OR_ENCRYPTED_WORKBOOK_MESSAGE);
    }
  });

  it("leaves the .docx path's own .doc message for every other compound file", async () => {
    for (const name of ["letter.doc", "letter.docx", "deck.ppt", "noextension"]) {
      const open = openAnyDocumentFile(asFile(COMPOUND_FILE, name));
      await expect(open).rejects.toThrow(LEGACY_DOC_MESSAGE);
      await expect(openAnyDocumentFile(asFile(COMPOUND_FILE, name))).rejects.toBeInstanceOf(LegacyDocError);
    }
  });

  it("refuses the committed Numbers document — by content, whatever its name — before the Keynote dispatch", async () => {
    const numbers = new Uint8Array(readFileSync(join(e2eFixturesDir, "spec116-numbers.numbers")));
    for (const name of ["spec116-numbers.numbers", "renamed.zip"]) {
      const open = openAnyDocumentFile(asFile(numbers, name));
      await expect(open).rejects.toThrow(NUMBERS_UNSUPPORTED_MESSAGE);
    }
  });

  it("still imports a Keynote deck (the same engine entry plus slides)", async () => {
    const deck = new Uint8Array(readFileSync(join(e2eFixturesDir, "mini-deck.key")));
    const entries = [...(await unzip(deck)).keys()];
    expect(entries).toContain("Index/CalculationEngine.iwa");
    expect(entries.some((entry) => entry.startsWith("Index/Slide"))).toBe(true);
    const opened = await openAnyDocumentFile(asFile(deck, "mini-deck.key"));
    expect(opened.kind).toBe("presentation");
  });

  it("keeps NOT_KEYNOTE_MESSAGE for a Pages-shaped package (suffixed engine, no slides)", async () => {
    const pages = await buildKeynoteFromArchives(
      [{ id: 1, type: 10000, payload: message(varintField(1, 1)) }],
      [],
      [{ name: "Index/CalculationEngine-1234567.iwa", data: Uint8Array.from([0, 1, 2]) }],
    );
    await expect(openAnyDocumentFile(asFile(pages, "letter.pages"))).rejects.toThrow(NOT_KEYNOTE_MESSAGE);
  });

  it("gives a .numbers-named ZIP with neither marker the Numbers message (the extension steers only to a better message)", async () => {
    const plain = await zip([{ name: "readme.txt", data: encoder.encode("hello") }]);
    await expect(openAnyDocumentFile(asFile(plain, "budget.numbers"))).rejects.toThrow(NUMBERS_UNSUPPORTED_MESSAGE);
    const bareIwork = await buildKeynoteFromArchives([{ id: 1, type: 1, payload: message(varintField(1, 1)) }]);
    await expect(openAnyDocumentFile(asFile(bareIwork, "budget.numbers"))).rejects.toThrow(NUMBERS_UNSUPPORTED_MESSAGE);
  });

  it("keeps FLAT_ODF_MESSAGE for a flat .fods", async () => {
    const fods = encoder.encode(
      `<?xml version="1.0" encoding="UTF-8"?><office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" ` +
        `office:version="1.3" office:mimetype="application/vnd.oasis.opendocument.spreadsheet"><office:body><office:spreadsheet/></office:body></office:document>`,
    );
    await expect(openAnyDocumentFile(asFile(fods, "flat.fods"))).rejects.toThrow(FLAT_ODF_MESSAGE);
  });

  it("names spreadsheets among what an unsupported OpenDocument file is not", async () => {
    expect(UNSUPPORTED_ODF_MESSAGE).toBe("This OpenDocument file isn't a text document, presentation or spreadsheet we can open");
    const drawing = await repackaged(await buildMiniBookOdsBytes(), "mimetype", "application/vnd.oasis.opendocument.graphics");
    await expect(openAnyDocumentFile(asFile(drawing, "sketch.odg"))).rejects.toThrow(UNSUPPORTED_ODF_MESSAGE);
  });

  it("refuses a workbook past the cell ceiling with WORKBOOK_TOO_LARGE_MESSAGE", async () => {
    const huge = await repackaged(
      await buildMiniBookOdsBytes(),
      "content.xml",
      `<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0">` +
        `<office:body><office:spreadsheet><table:table table:name="Huge"><table:table-row table:number-rows-repeated="200">` +
        `<table:table-cell office:value-type="float" office:value="1" table:number-columns-repeated="16384"/></table:table-row>` +
        `</table:table></office:spreadsheet></office:body></office:document-content>`,
    );
    await expect(openAnyDocumentFile(asFile(huge, "huge.ods"))).rejects.toThrow(WORKBOOK_TOO_LARGE_MESSAGE);
  });

  it("keeps the one 50 MB limit for every kind", async () => {
    expect(MAX_FILE_SIZE_BYTES).toBe(50 * 1024 * 1024);
    const oversized = new Uint8Array(MAX_FILE_SIZE_BYTES + 1);
    await expect(openAnyDocumentFile(asFile(oversized, "big.xlsx"))).rejects.toThrow(FILE_TOO_LARGE_MESSAGE);
  });
});

describe("the text-only entry point", () => {
  it("refuses an .xlsx and an .ods with SPREADSHEET_NOT_TEXT_MESSAGE instead of misparsing them", async () => {
    expect(SPREADSHEET_NOT_TEXT_MESSAGE).toBe("This is a spreadsheet file — open it from the library instead");
    for (const [bytes, name] of [
      [await buildMiniBookXlsxBytes(), "mini-book.xlsx"],
      [await buildMiniBookOdsBytes(), "mini-book.ods"],
    ] as const) {
      const open = openDocumentFile(asFile(bytes, name));
      await expect(open).rejects.toThrow(SPREADSHEET_NOT_TEXT_MESSAGE);
      await expect(openDocumentFile(asFile(bytes, name))).rejects.toBeInstanceOf(UnsupportedFileError);
    }
  });
});

describe("the open surface", () => {
  it("accepts the six workbook extensions at the drop check, and only those", () => {
    for (const name of ["a.xlsx", "a.xlsm", "a.xltx", "a.xltm", "a.ods", "a.ots", "A.XLSX"]) {
      expect(hasSupportedExtension(name)).toBe(true);
    }
    // SPEC-126 admits .csv/.tsv by extension (openDocument.csv.spec.ts).
    for (const name of ["a.xls", "a.xlsb", "a.numbers", "a.fods"]) {
      expect(hasSupportedExtension(name)).toBe(false);
    }
    expect(UNSUPPORTED_FILE_TYPE_MESSAGE).toBe(
      "We can only open .docx, .odt, .md, .pptx, .odp, .key, .xlsx, .ods, .csv, .tsv, .pdf and image files",
    );
  });

  it("strips the six workbook extensions from the title", () => {
    for (const extension of ["xlsx", "xlsm", "xltx", "xltm", "ods", "ots"]) {
      expect(titleFromFileName(`Budget 2026.${extension}`)).toBe("Budget 2026");
      expect(titleFromFileName(`Budget.${extension.toUpperCase()}`)).toBe("Budget");
    }
    expect(titleFromFileName("old.xls")).toBe("old.xls");
  });
});
