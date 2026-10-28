// @vitest-environment node
// SPEC-140 — readWorkbook: generated (or any) .xlsx/.ods bytes back into a
// SpreadsheetModel in plain Node — no File, no disk, no browser. The format
// comes from the package's own declarations (unzip, detectPackage,
// spreadsheetFormatOf), never a name; anything that is not such a workbook
// is a DocumentGenerationError.
import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";
import { SPREADSHEETML_MAIN_CONTENT_TYPES, XLSB_MAIN_CONTENT_TYPE } from "../core/packageDetect";
import { zip } from "../core/zipWriter";
import { DEFAULT_SPREADSHEET_TITLE } from "../spreadsheet/model";
import { buildMiniBookOdsBytes, buildSpecMacrosXlsmBytes } from "../spreadsheet/fixtures/build";
import { DocumentBuilder } from "./builder";
import { DocumentGenerationError } from "./errors";
import { PresentationBuilder } from "./presentation";
import { NOT_A_WORKBOOK_MESSAGE, readWorkbook } from "./readWorkbook";
import { WorkbookBuilder } from "./workbook";

const fixtures = join(dirname(fileURLToPath(import.meta.url)), "..", "..", "..", "e2e", "fixtures");
const encoder = new TextEncoder();

function sales(): WorkbookBuilder {
  return new WorkbookBuilder({ title: "Q3 sales" })
    .addSheet("Sales")
    .setRow("A1", ["Region", "July", "August"])
    .setRow("A2", ["East", 1200, 1350])
    .setFormula("D2", "SUM(B2:C2)")
    .addSheet("Notes", { hidden: true });
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

describe("readWorkbook (AC-34)", () => {
  it("reads generated .xlsx and .ods bytes in plain Node, each in its own format", async () => {
    expect(typeof window).toBe("undefined");
    for (const format of ["xlsx", "ods"] as const) {
      const model = await readWorkbook(await sales().toBytes(format));
      expect(model.format).toBe(format);
      expect(model.title).toBe(DEFAULT_SPREADSHEET_TITLE);
      expect(model.sheets.map((sheet) => [sheet.name, sheet.visibility])).toEqual([
        ["Sales", "visible"],
        ["Notes", "hidden"],
      ]);
      expect(model.sheets[0]!.rows[1]!.cells.map((cell) => [cell.value, cell.formula])).toEqual([
        ["East", undefined],
        [1200, undefined],
        [1350, undefined],
        [2550, "SUM(B2:C2)"],
      ]);
    }
  });

  it("takes a Uint8Array, a Buffer or an ArrayBuffer, and a title", async () => {
    const bytes = await sales().toBytes();
    const buffer = await sales().toBuffer();
    const arrayBuffer = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
    for (const input of [bytes, buffer, arrayBuffer]) {
      expect((await readWorkbook(input, "Q3 sales")).title).toBe("Q3 sales");
    }
  });

  it("decides the format from the package's declarations: a template or macro variant, an existing file", async () => {
    const xlsm = await readWorkbook(await buildSpecMacrosXlsmBytes());
    expect([xlsm.format, xlsm.packageVariant]).toEqual(["xlsx", "macroEnabled"]);
    expect((await readWorkbook(await buildMiniBookOdsBytes())).format).toBe("ods");
    const fixture = await readWorkbook(readFileSync(join(fixtures, "spec121-links.xlsx")));
    expect(fixture.sheets.map((sheet) => sheet.name)).toContain("Summary");
    expect(fixture.opaque.length).toBeGreaterThan(0);
  });

  it('refuses anything that is not an .xlsx/.ods workbook with "bytes: not an .xlsx or .ods workbook"', async () => {
    const notWorkbooks: [string, Uint8Array][] = [
      ["no bytes", new Uint8Array()],
      ["not a ZIP", encoder.encode("Region,July\nEast,1200\n")],
      ["a .docx", await new DocumentBuilder().addParagraph("x").toBytes("docx")],
      ["an .odt", await new DocumentBuilder().addParagraph("x").toBytes("odt")],
      ["a .pptx", await new PresentationBuilder().addTitleSlide("x").toBytes("pptx")],
      ["an .odp", await new PresentationBuilder().addTitleSlide("x").toBytes("odp")],
      ["an empty ZIP", await zip([{ name: "readme.txt", data: encoder.encode("hi") }])],
      ["an .xlsb", await opcPackage(XLSB_MAIN_CONTENT_TYPE)],
    ];
    for (const [label, bytes] of notWorkbooks) {
      const read = readWorkbook(bytes);
      await expect(read, label).rejects.toThrow(DocumentGenerationError);
      await expect(readWorkbook(bytes), label).rejects.toThrow(NOT_A_WORKBOOK_MESSAGE);
    }
    expect(NOT_A_WORKBOOK_MESSAGE).toBe("bytes: not an .xlsx or .ods workbook");
    expect(SPREADSHEETML_MAIN_CONTENT_TYPES.workbook).toContain("spreadsheetml");
  });
});
