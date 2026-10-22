// SPEC-136's round trip and fixtures: a no-edit download is byte-identical
// part for part; editing every setup field → save → reopen reproduces the
// model in both formats, and writing the reopened model twice is byte-stable;
// the sanctioned deltas (margins to 1/1000 inch in xlsx, an ods header with
// an unmapped field regenerated once edited, scale-to-pages → fit 1 × N); a
// copy across formats converts the setup field for field; printing touches
// nothing (buildPdf twice on the same pages is byte-identical).
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { equalBytes } from "../core/equality";
import { buildPdf } from "../core/pdfBuild";
import { unzip } from "../core/zip";
import { zip } from "../core/zipWriter";
import { spreadsheetDownload } from "./download";
import { buildSpec136PrintOdsBytes, buildSpec136PrintXlsxBytes, SPEC136_PRINT_ODS, SPEC136_PRINT_XLSX } from "./fixtures/build";
import { packageIntegrity } from "./fixtures/packageIntegrity";
import type { SheetPageSetup, SpreadsheetFormat, SpreadsheetModel } from "./model";
import { readSpreadsheetBytes } from "./open";
import { getReservedName } from "./names";
import { setPageSetup, setPrintArea, setPrintTitles } from "./pageSetup";
import { stripRetention } from "./sheets";
import { writeOds } from "./ods/writer";
import { writeXlsx } from "./xlsx/writer";

const E2E_FIXTURES = join(__dirname, "../../../e2e/fixtures");
const decoder = new TextDecoder();

const write = (model: SpreadsheetModel): Promise<Uint8Array> => (model.format === "xlsx" ? writeXlsx(model) : writeOds(model));
const read = (bytes: Uint8Array, format: SpreadsheetFormat): Promise<SpreadsheetModel> => readSpreadsheetBytes(bytes, format, "spec136-print");

async function differingParts(a: Uint8Array, b: Uint8Array): Promise<string[]> {
  const [left, right] = [await unzip(a), await unzip(b)];
  const names = new Set([...left.keys(), ...right.keys()]);
  return [...names].filter((name) => !left.has(name) || !right.has(name) || !equalBytes(left.get(name)!, right.get(name)!));
}

/** A setup with every field set, each value one both formats hold exactly. */
function everyField(format: SpreadsheetFormat): SheetPageSetup {
  return {
    orientation: "landscape",
    paper: { widthCm: 29.7, heightCm: 42, xlsxPaperSize: 8 },
    scale: 85,
    margins: { topCm: 2.54, rightCm: 1.524, bottomCm: 2.286, leftCm: 1.27, headerCm: 1.016, footerCm: 0.762 },
    centerHorizontally: true,
    centerVertically: true,
    printGridLines: true,
    printHeadings: true,
    pageOrder: "overThenDown",
    firstPageNumber: 3,
    headerFooter: {
      odd: { header: { left: "&A", center: "Page &P of &N", right: "&BBold&B" }, footer: { center: "&D &T" } },
      even: { header: { center: "Even &P" } },
      first: { footer: { right: "&F" } },
      differentOddEven: true,
      differentFirst: true,
      ...(format === "xlsx" ? { scaleWithDoc: false as const, alignWithMargins: false as const } : {}),
    },
    rowBreaks: [5, 12],
    colBreaks: [2],
  };
}

for (const [format, build] of [
  ["xlsx", buildSpec136PrintXlsxBytes],
  ["ods", buildSpec136PrintOdsBytes],
] as const) {
  describe(`SPEC-136 round trip — spec136-print.${format}`, () => {
    it("is deterministic, the committed bytes, whole", async () => {
      const bytes = await build();
      expect(equalBytes(bytes, await build())).toBe(true);
      expect(equalBytes(new Uint8Array(readFileSync(join(E2E_FIXTURES, `spec136-print.${format}`))), bytes)).toBe(true);
      expect(await packageIntegrity(bytes)).toEqual([]);
    });

    it("a no-edit download is byte-identical part for part", async () => {
      const bytes = await build();
      expect(await differingParts(await write(await read(bytes, format)), bytes)).toEqual([]);
    });

    it("every setup field edited → save → reopen reproduces the model; the reopened model writes twice byte-stable", async () => {
      const model = await read(await build(), format);
      const budget = model.sheets[0]!;
      // Percent mode replaces the fixture's fit (an .ods page layout holds one scale form).
      setPageSetup(model, budget, { ...everyField(format), fitTo: undefined, odsScaleToPages: undefined });
      setPrintArea(model, budget.id, [{ top: 0, left: 0, bottom: 39, right: 5 }]);
      setPrintTitles(model, budget.id, { rows: [0, 1] });
      const reopened = await read(await write(model), format);
      expect(reopened.sheets[0]!.pageSetup).toEqual(model.sheets[0]!.pageSetup);
      expect(getReservedName(reopened, reopened.sheets[0]!.id, "Print_Area")!.formula).toBe("Budget!$A$1:$F$40");
      expect(getReservedName(reopened, reopened.sheets[0]!.id, "Print_Titles")!.formula).toBe("Budget!$1:$2");
      expect(reopened.sheets[1]!.pageSetup).toEqual(model.sheets[1]!.pageSetup);
      // Printing reads, never writes: a second save of the reopened model is the same bytes.
      const once = await write(reopened);
      expect(equalBytes(once, await write(reopened))).toBe(true);
    });
  });
}

describe("SPEC-136 sanctioned deltas", () => {
  it("xlsx: margins round to 1/1000 inch", async () => {
    const model = await read(await buildSpec136PrintXlsxBytes(), "xlsx");
    setPageSetup(model, model.sheets[0]!, { margins: { topCm: 1.8, rightCm: 1.27, bottomCm: 1.905, leftCm: 1.27, headerCm: 0.762, footerCm: 0.762 } });
    const reopened = await read(await writeXlsx(model), "xlsx");
    expect(reopened.sheets[0]!.pageSetup!.margins!.topCm).toBe(1.80086);
  });

  it("ods: a header region holding an unmapped field is regenerated with that field's text once its setup is edited", async () => {
    const parts = await unzip(await buildSpec136PrintOdsBytes());
    const styles = decoder.decode(parts.get("styles.xml")).replace("<text:sheet-name>???</text:sheet-name>", "<text:title>Plan</text:title>");
    const bytes = await zip([...parts].map(([name, data]) => (name === "styles.xml" ? { name, data: new TextEncoder().encode(styles) } : { name, data, ...(name === "mimetype" ? { store: true } : {}) })));
    const model = await read(bytes, "ods");
    expect(model.sheets[0]!.pageSetup!.headerFooter!.odd.header!.left).toBe("Plan");
    const hf = model.sheets[0]!.pageSetup!.headerFooter!;
    setPageSetup(model, model.sheets[0]!, { headerFooter: { ...hf, odd: { ...hf.odd, header: { ...hf.odd.header, center: "Edited" } } } });
    const written = decoder.decode((await unzip(await writeOds(model))).get("styles.xml"));
    const minted = /<style:master-page style:name="de-mp\d+"[\s\S]*?<\/style:master-page>/.exec(written)![0];
    expect(minted).toContain("<style:region-left><text:p>Plan</text:p></style:region-left>");
    expect(minted).not.toContain("text:title");
  });

  it("ods style:scale-to-pages converts to xlsx as fit 1 × N", async () => {
    const parts = await unzip(await buildSpec136PrintOdsBytes());
    const styles = decoder.decode(parts.get("styles.xml")).replace(`style:scale-to-X="1"`, `style:scale-to-pages="2"`);
    const bytes = await zip([...parts].map(([name, data]) => (name === "styles.xml" ? { name, data: new TextEncoder().encode(styles) } : { name, data, ...(name === "mimetype" ? { store: true } : {}) })));
    const model = await read(bytes, "ods");
    expect(model.sheets[0]!.pageSetup!.odsScaleToPages).toBe(2);
    const copy = await read((await spreadsheetDownload(model, "xlsx")).bytes, "xlsx");
    expect(copy.sheets[0]!.pageSetup!.fitTo).toEqual({ width: 1, height: 2 });
    expect(copy.sheets[0]!.pageSetup!.odsScaleToPages).toBeUndefined();
  });
});

describe("SPEC-136 across formats and without mutation", () => {
  it("Download a copy as .ods converts the setup field for field (stripRetention keeps pageSetup, drops pageSetupDirty) — print area and titles included", async () => {
    const model = await read(await buildSpec136PrintXlsxBytes(), "xlsx");
    setPageSetup(model, model.sheets[0]!, { centerVertically: true });
    expect(model.sheets[0]!.pageSetupDirty).toBe(true);
    const stripped = stripRetention(model);
    expect(stripped.sheets[0]!.pageSetupDirty).toBeUndefined();
    expect(stripped.sheets[0]!.pageSetup).toEqual(model.sheets[0]!.pageSetup);
    const copy = await read((await spreadsheetDownload(model, "ods")).bytes, "ods");
    const { headerFooter: _hf, ...layout } = model.sheets[0]!.pageSetup!;
    const { headerFooter: copiedHf, ...copied } = copy.sheets[0]!.pageSetup!;
    expect(copied).toEqual(layout);
    expect(copiedHf!.odd.header).toEqual({ left: "Budget", center: "&P of &N", right: "&D" });
    expect(getReservedName(copy, copy.sheets[0]!.id, "Print_Area")!.formula).toBe(SPEC136_PRINT_XLSX.printArea);
    expect(getReservedName(copy, copy.sheets[0]!.id, "Print_Titles")!.formula).toBe(SPEC136_PRINT_XLSX.printTitles);
  });

  it("the fixtures hold what the specs say: the xlsx printer settings and names, the ods shared master page and header rows", async () => {
    const xlsx = await unzip(await buildSpec136PrintXlsxBytes());
    expect(decoder.decode(xlsx.get("xl/worksheets/sheet1.xml"))).toContain(`r:id="rId1"`);
    expect(xlsx.has(SPEC136_PRINT_XLSX.printerSettings)).toBe(true);
    expect(decoder.decode(xlsx.get("xl/workbook.xml"))).toContain(`<definedName name="_xlnm.Print_Titles" localSheetId="0">Budget!$1:$1</definedName>`);
    const ods = await unzip(await buildSpec136PrintOdsBytes());
    const content = decoder.decode(ods.get("content.xml"));
    expect(content.match(/table:style-name="ta1"/g)).toHaveLength(2);
    expect(content).toContain(`table:print-ranges="${SPEC136_PRINT_ODS.printRanges}"`);
    expect(content).toContain("<table:table-header-rows>");
  });

  it("buildPdf twice on the same pages is byte-identical", async () => {
    const pages = [{ widthPt: 841.89, heightPt: 595.28, items: [], links: [{ xPt: 10, yPt: 10, widthPt: 20, heightPt: 10, href: "https://example.com/" }], headings: [{ level: 1, title: "Budget", yPt: 0 }], images: [] }];
    const first = await buildPdf({ title: "spec136-print", pages, fonts: new Map() });
    expect(equalBytes(first, await buildPdf({ title: "spec136-print", pages, fonts: new Map() }))).toBe(true);
  });
});
