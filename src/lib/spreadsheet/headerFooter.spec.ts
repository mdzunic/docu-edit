// SPEC-136 "Header and footer codes": the section split and its inverse,
// every code's tokenizing and rendering, unknown codes kept literal, and the
// ODS field mapping.
import { describe, expect, it } from "vitest";
import { formatHeaderFooter, ODS_FIELD_CODES, parseHeaderFooter, renderHeaderFooterSection, tokenizeHeaderFooter, type HeaderFooterContext } from "./headerFooter";
import { SPEC136_PRINT_ODS, SPEC136_PRINT_XLSX } from "./fixtures/build";
import { SPREADSHEET_LOCALES } from "./numberFormat/locale";

const ctx: HeaderFooterContext = {
  page: 2,
  pages: 7,
  date: new Date(2026, 9, 22, 14, 5, 0),
  sheetName: "Budget",
  fileName: "budget.xlsx",
  locale: SPREADSHEET_LOCALES["en-US"],
  dateSystem: "1900",
};

const text = (code: string): string =>
  renderHeaderFooterSection(code, ctx)
    .map((run) => run.text)
    .join("");

describe("parseHeaderFooter / formatHeaderFooter", () => {
  it("splits on &L/&C/&R (text before the first code is the centre), and formats back every fixture string unchanged", () => {
    expect(parseHeaderFooter(SPEC136_PRINT_XLSX.header)).toEqual({ left: '&"Arial,Bold"&12Budget', center: "&P of &N", right: "&D" });
    expect(parseHeaderFooter("Page &P")).toEqual({ center: "Page &P" });
    expect(parseHeaderFooter("&&L&LLeft")).toEqual({ center: "&&L", left: "Left" });
    for (const fixture of [SPEC136_PRINT_XLSX.header, SPEC136_PRINT_XLSX.oddFooter, SPEC136_PRINT_XLSX.firstFooter, `&L${SPEC136_PRINT_ODS.header.left}&C${SPEC136_PRINT_ODS.header.center}&R${SPEC136_PRINT_ODS.header.right}`]) {
      expect(formatHeaderFooter(parseHeaderFooter(fixture))).toBe(fixture);
    }
  });
});

describe("tokenizeHeaderFooter / renderHeaderFooterSection", () => {
  it("renders the fields: &P (&P+n, &P-n), &N, &D and &T in the locale, &F, &A, &&", () => {
    expect(text("Page &P of &N")).toBe("Page 2 of 7");
    expect(text("&P+10 &P-1")).toBe("12 1");
    expect(text("&D")).toBe("10/22/2026");
    expect(text("&T")).toMatch(/^2:05 PM$/);
    expect(text("&F / &A")).toBe("budget.xlsx / Budget");
    expect(text("Q&&A")).toBe("Q&A");
  });

  it("styles runs: &\"font,style\", &nn, &Krrggbb, &B, &I, &U, &E, &S toggles", () => {
    expect(renderHeaderFooterSection('&"Arial,Bold"&12Budget', ctx)).toEqual([{ text: "Budget", bold: true, fontFamily: "Arial", fontSizePt: 12 }]);
    expect(renderHeaderFooterSection("&KFF0000red&B bold&B &Iit&I &Uu&U &Ee&E &Ss", ctx)).toEqual([
      { text: "red", color: "#ff0000" },
      { text: " bold", color: "#ff0000", bold: true },
      { text: " ", color: "#ff0000" },
      { text: "it", color: "#ff0000", italic: true },
      { text: " ", color: "#ff0000" },
      { text: "u", color: "#ff0000", underline: "single" },
      { text: " ", color: "#ff0000" },
      { text: "e", color: "#ff0000", underline: "double" },
      { text: " ", color: "#ff0000" },
      { text: "s", color: "#ff0000", strike: true },
    ]);
  });

  it("&Z renders empty, &G nothing; &X/&Y/&O/&H and a theme &K are plain text in the default colour; anything unknown is literal", () => {
    expect(renderHeaderFooterSection("&Z", ctx)).toEqual([{ text: "" }]);
    expect(renderHeaderFooterSection("&G", ctx)).toEqual([]);
    expect(renderHeaderFooterSection("&Xa&Yb&Oc&Hd&K01+000e", ctx)).toEqual([{ text: "abcde" }]);
    expect(text("&Q&")).toBe("&Q&");
    expect(tokenizeHeaderFooter("&Q&P")).toEqual([{ kind: "text", text: "&Q" }, { kind: "page", offset: 0 }]);
  });
});

describe("the ODS field mapping", () => {
  it("text:page-number, page-count, date, time, sheet-name and file-name are &P, &N, &D, &T, &A and &F", () => {
    expect(ODS_FIELD_CODES).toEqual({ "page-number": "&P", "page-count": "&N", date: "&D", time: "&T", "sheet-name": "&A", "file-name": "&F" });
  });
});
