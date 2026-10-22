// SPEC-136 "xlsx wire mapping": reading a worksheet's page elements, and the
// attribute-level splice — unmodelled attributes keep their bytes and place,
// missing elements go in at their CT_Worksheet place, an unchanged setup
// returns the very string.
import { describe, expect, it } from "vitest";
import { formatInches, readXlsxPageSetup, spliceXlsxPageSetup } from "./pageSetup";

const MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const worksheet = (tail: string, pr = ""): string =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<worksheet xmlns="${MAIN}" xmlns:r="${REL}">${pr}<dimension ref="A1"/><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>${tail}</worksheet>`;

const PAGE_SETUP = `<pageSetup paperSize="9" orientation="portrait" horizontalDpi="600" copies="2" blackAndWhite="1" draft="1" cellComments="atEnd" errors="blank" r:id="rId1"/>`;

describe("readXlsxPageSetup", () => {
  it("reads printOptions, pageMargins (inches), pageSetup, sheetPr/pageSetUpPr, headerFooter and the manual breaks", () => {
    const xml = worksheet(
      `<hyperlinks/><printOptions horizontalCentered="1" verticalCentered="true" gridLines="1" headings="1" gridLinesSet="0"/>` +
        `<pageMargins left="0.5" right="0.5" top="1" bottom="0.75" header="0.3" footer="0.3"/>` +
        `<pageSetup paperSize="8" scale="85" fitToWidth="2" fitToHeight="0" orientation="landscape" pageOrder="overThenDown" firstPageNumber="5" useFirstPageNumber="1" r:id="rId1"/>` +
        `<headerFooter differentOddEven="1" differentFirst="1" scaleWithDoc="0" alignWithMargins="0"><oddHeader>&amp;CPage &amp;P</oddHeader><evenFooter>&amp;LEven</evenFooter><firstHeader>&amp;RFirst</firstHeader></headerFooter>` +
        `<rowBreaks count="2" manualBreakCount="1"><brk id="10" max="16383" man="1"/><brk id="20" max="16383"/></rowBreaks><colBreaks count="1" manualBreakCount="1"><brk id="3" max="1048575" man="1"/></colBreaks>`,
      `<sheetPr><tabColor rgb="FFFF0000"/><pageSetUpPr fitToPage="1"/></sheetPr>`,
    );
    expect(readXlsxPageSetup(xml)).toEqual({
      orientation: "landscape",
      paper: { widthCm: 29.7, heightCm: 42, xlsxPaperSize: 8 },
      scale: 85,
      fitTo: { width: 2, height: 0 },
      margins: { topCm: 2.54, rightCm: 1.27, bottomCm: 1.905, leftCm: 1.27, headerCm: 0.762, footerCm: 0.762 },
      centerHorizontally: true,
      centerVertically: true,
      printGridLines: true,
      printHeadings: true,
      pageOrder: "overThenDown",
      firstPageNumber: 5,
      headerFooter: {
        odd: { header: { center: "Page &P" } },
        even: { footer: { left: "Even" } },
        first: { header: { right: "First" } },
        differentOddEven: true,
        differentFirst: true,
        scaleWithDoc: false,
        alignWithMargins: false,
      },
      rowBreaks: [10],
      colBreaks: [3],
    });
  });

  it("a part that says nothing (or only Excel's default margins) has no setup; a pageSetup without paperSize is Letter; an unknown code is kept; paperWidth/paperHeight are a custom paper", () => {
    expect(readXlsxPageSetup(worksheet(""))).toBeUndefined();
    expect(readXlsxPageSetup(worksheet(`<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>`))).toBeUndefined();
    expect(readXlsxPageSetup(worksheet(`<pageSetup orientation="portrait" r:id="rId1"/>`))).toEqual({ paper: { widthCm: 21.59, heightCm: 27.94, xlsxPaperSize: 1 } });
    expect(readXlsxPageSetup(worksheet(`<pageSetup paperSize="42"/>`))!.paper).toEqual({ widthCm: 21.59, heightCm: 27.94, xlsxPaperSize: 42 });
    expect(readXlsxPageSetup(worksheet(`<pageSetup paperWidth="210mm" paperHeight="11.69in"/>`))!.paper).toEqual({ widthCm: 21, heightCm: expect.closeTo(29.6926, 3) });
  });
});

describe("spliceXlsxPageSetup", () => {
  it("returns the very string when the part already reads as the setup", () => {
    const xml = worksheet(`<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>${PAGE_SETUP}`);
    expect(spliceXlsxPageSetup(xml, readXlsxPageSetup(xml))).toBe(xml);
  });

  it("splices an attribute in place: r:id, horizontalDpi, copies, blackAndWhite, draft, cellComments and errors keep their bytes and position", () => {
    const xml = worksheet(PAGE_SETUP);
    const next = spliceXlsxPageSetup(xml, { ...readXlsxPageSetup(xml), orientation: "landscape" });
    expect(next).toBe(xml.replace(`orientation="portrait"`, `orientation="landscape"`));
  });

  it("inserts missing elements in CT_Worksheet order and pageSetUpPr as the last of sheetPr's children; margins go to 1/1000 inch", () => {
    const xml = worksheet(`<hyperlinks><hyperlink ref="A1" r:id="rId2"/></hyperlinks><drawing r:id="rId3"/>`, `<sheetPr><tabColor rgb="FFFF0000"/></sheetPr>`);
    const next = spliceXlsxPageSetup(xml, {
      fitTo: { width: 1, height: 0 },
      printGridLines: true,
      margins: { topCm: 1.8, rightCm: 1.778, bottomCm: 1.905, leftCm: 1.778, headerCm: 0.762, footerCm: 0.762 },
      headerFooter: { odd: { header: { center: "&P & co" } } },
      rowBreaks: [4],
    });
    expect(next).toContain(`<sheetPr><tabColor rgb="FFFF0000"/><pageSetUpPr fitToPage="1"/></sheetPr>`);
    expect(next).toContain(
      `</hyperlinks><printOptions gridLines="1"/><pageMargins left="0.7" right="0.7" top="0.709" bottom="0.75" header="0.3" footer="0.3"/><pageSetup fitToHeight="0"/>` +
        `<headerFooter><oddHeader>&amp;C&amp;P &amp; co</oddHeader></headerFooter><rowBreaks count="1" manualBreakCount="1"><brk id="4" max="16383" man="1"/></rowBreaks><drawing r:id="rId3"/>`,
    );
    expect(readXlsxPageSetup(next)!.margins!.topCm).toBeCloseTo(1.80086, 5);
    expect(formatInches(1.778)).toBe("0.7");
  });

  it("a break list rewritten keeps its start tag, the breaks that stay and every break that is not manual verbatim", () => {
    const breaks = `<rowBreaks count="3" manualBreakCount="2" x14:extra="kept"><brk id="10" max="16383" man="1" pt="1"/><brk id="20" max="16383"/><brk id="30" max="16383" man="1"/></rowBreaks>`;
    const xml = worksheet(breaks);
    const setup = readXlsxPageSetup(xml)!;
    expect(setup.rowBreaks).toEqual([10, 30]);
    // 30 removed, 25 added: 10 and the automatic 20 keep their bytes.
    expect(spliceXlsxPageSetup(xml, { ...setup, rowBreaks: [10, 25] })).toBe(
      worksheet(
        `<rowBreaks count="3" manualBreakCount="2" x14:extra="kept"><brk id="10" max="16383" man="1" pt="1"/><brk id="20" max="16383"/><brk id="25" max="16383" man="1"/></rowBreaks>`,
      ),
    );
    // A manual break where an automatic one was replaces it; with no manual break left the automatic ones stay.
    expect(spliceXlsxPageSetup(xml, { ...setup, rowBreaks: [10, 20] })).toContain(`<brk id="10" max="16383" man="1" pt="1"/><brk id="20" max="16383" man="1"/></rowBreaks>`);
    expect(spliceXlsxPageSetup(xml, { ...setup, rowBreaks: undefined })).toContain(`<rowBreaks count="1" manualBreakCount="0" x14:extra="kept"><brk id="20" max="16383"/></rowBreaks>`);
    expect(spliceXlsxPageSetup(worksheet(`<rowBreaks count="1" manualBreakCount="1"><brk id="4" max="16383" man="1"/></rowBreaks>`), {})).toBe(worksheet(""));
  });

  it("an ods scale-to-pages setup writes as fit 1 × N", () => {
    const next = spliceXlsxPageSetup(worksheet(""), { odsScaleToPages: 3 });
    expect(readXlsxPageSetup(next)).toEqual({ paper: { widthCm: 21.59, heightCm: 27.94, xlsxPaperSize: 1 }, fitTo: { width: 1, height: 3 } });
  });
});
