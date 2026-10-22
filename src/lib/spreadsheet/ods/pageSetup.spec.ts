// SPEC-136 "ods wire mapping": a table's style → master page → page layout,
// the header/footer heights folded into the margins, `loext:` tolerance,
// print ranges and header rows/columns as the reserved names, row breaks,
// and minting a master page rather than editing a shared one.
import { describe, expect, it } from "vitest";
import { unzip } from "../../core/zip";
import { buildSpec136PrintOdsBytes } from "../fixtures/build";
import type { Sheet } from "../model";
import { odsMasterPageSetup, odsTableStructure, OdsPageSetupWriter, readOdsPageStyles, readOdsTablePageSetup, withOdsPrintRanges } from "./pageSetup";

const STYLE_NS = `xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0" xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"`;

function stylesXml(layout: string, master: string): string {
  return (
    `<?xml version="1.0" encoding="UTF-8"?><office:document-styles ${STYLE_NS}><office:automatic-styles>` +
    `<style:style style:name="T1" style:family="text"><style:text-properties fo:font-style="italic"/></style:style>${layout}</office:automatic-styles>` +
    `<office:master-styles>${master}</office:master-styles></office:document-styles>`
  );
}

async function fixtureParts(): Promise<{ styles: string; content: string }> {
  const parts = await unzip(await buildSpec136PrintOdsBytes());
  const decoder = new TextDecoder();
  return { styles: decoder.decode(parts.get("styles.xml")), content: decoder.decode(parts.get("content.xml")) };
}

describe("reading", () => {
  it("the fixture's master page: A4 landscape, fit 1 wide, grid and headers, centred, the header regions as codes, the margins folded", async () => {
    const { styles } = await fixtureParts();
    expect(odsMasterPageSetup(readOdsPageStyles(styles), "Default")).toEqual({
      orientation: "landscape",
      paper: { widthCm: 21, heightCm: 29.7, xlsxPaperSize: 9 },
      fitTo: { width: 1, height: 0 },
      margins: { topCm: 1.905, rightCm: 1.27, bottomCm: 1.905, leftCm: 1.27, headerCm: 0.762, footerCm: 0.762 },
      centerHorizontally: true,
      printGridLines: true,
      printHeadings: true,
      headerFooter: { odd: { header: { left: "&A", center: "&P of &N", right: "&BBudget&B" }, footer: { center: "Page &P" } } },
    });
  });

  it("a hidden header leaves topCm the page's margin and headerCm the default; loext:scale-to-X/-Y and scale-to-pages are read; spans and unmapped fields become codes and text", () => {
    const styles = stylesXml(
      `<style:page-layout style:name="pm1"><style:page-layout-properties fo:margin-top="1cm" fo:margin-bottom="2cm" fo:margin-left="1cm" fo:margin-right="1cm" loext:scale-to-X="2" loext:scale-to-Y="3" style:print-page-order="ltr" style:first-page-number="4"/>` +
        `<style:footer-style><style:header-footer-properties fo:min-height="0.5cm" fo:margin-top="0.25cm"/></style:footer-style></style:page-layout>`,
      `<style:master-page style:name="M" style:page-layout-name="pm1"><style:header style:display="false"><text:p>gone</text:p></style:header>` +
        `<style:footer><style:region-right><text:p><text:span text:style-name="T1">By</text:span> <text:author-name>Ann &amp; Bo</text:author-name> <text:title>Plan</text:title></text:p></style:region-right></style:footer></style:master-page>`,
    );
    expect(odsMasterPageSetup(readOdsPageStyles(styles), "M")).toEqual({
      fitTo: { width: 2, height: 3 },
      margins: { topCm: 1, rightCm: 1, bottomCm: 2.75, leftCm: 1, headerCm: 0.762, footerCm: 2 },
      pageOrder: "overThenDown",
      firstPageNumber: 4,
      headerFooter: { odd: { footer: { right: "&IBy&I Ann && Bo Plan" } } },
    });
  });

  it("print ranges and header rows/columns are the reserved names; fo:break-before rows and columns the breaks", () => {
    const sheet = {
      id: "s1",
      name: "Data",
      sourceXml:
        `<table:table table:name="Data" table:style-name="ta1" table:print-ranges="Data.A1:Data.D20">` +
        `<table:table-header-columns><table:table-column table:style-name="co1"/></table:table-header-columns><table:table-column table:style-name="co2" table:number-columns-repeated="3"/>` +
        `<table:table-header-rows><table:table-row table:style-name="ro1" table:number-rows-repeated="2"/></table:table-header-rows><table:table-row table:style-name="ro1" table:number-rows-repeated="5"/><table:table-row table:style-name="ro2"/></table:table>`,
    } as Pick<Sheet, "id" | "name" | "sourceXml">;
    const prefix = `<style:style style:name="ro2" style:family="table-row"><style:table-row-properties fo:break-before="page"/></style:style><style:style style:name="co2" style:family="table-column"><style:table-column-properties fo:break-before="page"/></style:style>`;
    const read = readOdsTablePageSetup(sheet, prefix, readOdsPageStyles(undefined));
    expect(read.setup).toEqual({ rowBreaks: [7], colBreaks: [1] });
    expect(read.names).toEqual([
      { name: "_xlnm.Print_Area", scopeSheetId: "s1", formula: "Data!$A$1:$D$20" },
      { name: "_xlnm.Print_Titles", scopeSheetId: "s1", formula: "Data!$1:$2,Data!$A:$A" },
    ]);
    expect(odsTableStructure("<table:table/>", { rows: new Set(), cols: new Set() })).toEqual({ titles: {}, rowBreaks: [], colBreaks: [] });
  });
});

describe("writing", () => {
  it("a changed setup mints de-pm/de-mp copies of the table's master page; styles.xml's other bytes stay — the shared master page untouched", async () => {
    const { styles, content } = await fixtureParts();
    const prefix = content.slice(0, content.indexOf("<office:body>"));
    let n = 0;
    const writer = new OdsPageSetupWriter(prefix, styles, (kind) => `de-${kind}${(n += 1)}`);
    const read = odsMasterPageSetup(readOdsPageStyles(styles), "Default")!;
    const sheet = { pageSetup: { ...read, headerFooter: { odd: { ...read.headerFooter!.odd, header: { center: "Budget &P" } } } } } as Sheet;
    expect(writer.masterPageFor({ pageSetup: read } as Sheet, "ta1")).toBeUndefined();
    const master = writer.masterPageFor(sheet, "ta1");
    expect(master).toBe("de-mp2");
    const next = writer.stylesXmlWith(styles);
    const minted = /<style:page-layout style:name="de-pm1">[\s\S]*?<\/style:page-layout>/.exec(next)![0];
    const mintedMaster = /<style:master-page style:name="de-mp2"[\s\S]*?<\/style:master-page>/.exec(next)![0];
    expect(next.replace(minted, "").replace(mintedMaster, "")).toBe(styles);
    expect(mintedMaster).toContain(`<style:header><style:region-center><text:p>Budget <text:page-number>1</text:page-number></text:p></style:region-center></style:header>`);
    expect(mintedMaster).toContain(`<style:footer><text:p>Page <text:page-number>1</text:page-number></text:p></style:footer>`);
    expect(odsMasterPageSetup(readOdsPageStyles(next), "de-mp2")).toEqual(sheet.pageSetup);
  });

  it("withOdsPrintRanges writes the model's print area through canonicalToOdsRangeAddress, keeping the attribute while it names the same ranges", () => {
    const table = `<table:table table:name="Data" table:print-ranges="Data.A1:Data.D20"></table:table>`;
    const names = (formula: string | null) => ({ names: formula === null ? [] : [{ name: "_xlnm.Print_Area", scopeSheetId: "s1", formula }] });
    expect(withOdsPrintRanges(names("Data!$A$1:$D$20"), { id: "s1" }, table)).toBe(table);
    expect(withOdsPrintRanges(names("Data!$A$1:$D$22"), { id: "s1" }, table)).toBe(`<table:table table:name="Data" table:print-ranges="$Data.$A$1:.$D$22"></table:table>`);
    expect(withOdsPrintRanges(names(null), { id: "s1" }, table)).toBe(`<table:table table:name="Data"></table:table>`);
  });
});
