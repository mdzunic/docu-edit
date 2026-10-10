// SPEC-117 "Reading ODS styles" and "ODS borders": the property mapping
// table row by row, the border read table row by row (width tiers, units,
// none/hidden) and its write direction, per-property flattening along the
// parent chain, font-face resolution, the style table's first-reference
// order and sourceStyleName, and the LibreOffice-shaped fixture's styles.
// SPEC-122 "Styles" (ods): the generated de-ce/de-T/de-N families — the
// property table's inverse row by row (a written style reads back equal,
// the documented approximations aside), naming past a collision, number
// styles through odsNumberStyleXml with their P0 sub-styles, a built-in
// written in the user's locale, equal property sets sharing one name, the
// parent style, and a rich run keeping its span's own text style.
import { describe, expect, it } from "vitest";
import { DEFAULT_THEME, resolveColor, resolveFont } from "../../core/theme";
import { parseA1 } from "../address";
import { buildSpec117StylesOdsBytes, SPEC117_STYLES_ODS } from "../fixtures/build";
import type { CellFont, CellStyle, SheetBorderStyle } from "../model";
import { formatCodeOf } from "../numberFormat/builtin";
import { SPREADSHEET_LOCALES } from "../numberFormat/locale";
import { odsNumberStyleXml } from "../numberFormat/odsNumberStyles";
import { cellStyle, effectiveStyleId } from "../styles";
import { parseOds } from "./reader";
import {
  emptyOdsStyleSheet,
  ODS_BORDER_WRITE,
  odsBorderEdge,
  odsBorderString,
  odsLengthToPt,
  odsLengthToPx,
  odsTextRotation,
  OdsStyleBook,
  OdsStyleRegistry,
  readOdsContentStyles,
  readOdsStyleSheet,
  type OdsStyleEnv,
} from "./styles";

const NS =
  `xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" ` +
  `xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" ` +
  `xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" ` +
  `xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:tableooo="http://openoffice.org/2009/table" ` +
  `xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"`;

/** LibreOffice's table-cell default style and its two common styles. */
const LO_COMMON =
  `<style:default-style style:family="table-cell"><style:paragraph-properties style:tab-stop-distance="0.5in"/>` +
  `<style:text-properties style:font-name="Liberation Sans" fo:font-size="10pt" fo:language="en" fo:country="US"/></style:default-style>` +
  `<style:style style:name="Default" style:family="table-cell"/>` +
  `<style:style style:name="Heading" style:family="table-cell" style:parent-style-name="Default">` +
  `<style:text-properties fo:font-size="24pt" fo:font-weight="bold"/></style:style>`;

const LO_FONT_FACES =
  `<office:font-face-decls><style:font-face style:name="Liberation Sans" svg:font-family="&apos;Liberation Sans&apos;" ` +
  `style:font-family-generic="swiss"/></office:font-face-decls>`;

function styleBook(automatic: string, common = "", fontFaces = "", contentFontFaces = ""): OdsStyleBook {
  const sheet = readOdsStyleSheet(`<office:document-styles ${NS}>${fontFaces}<office:styles>${common}</office:styles></office:document-styles>`);
  const content = readOdsContentStyles([contentFontFaces, `<office:automatic-styles>${automatic}</office:automatic-styles>`], NS);
  return new OdsStyleBook(sheet, content);
}

/** The CellStyle an automatic style "ce1" with these property elements
 * maps to — no default style underneath, sourceStyleName dropped. */
function mapped(properties: string): Omit<CellStyle, "sourceStyleName"> {
  const book = styleBook(`<style:style style:name="ce1" style:family="table-cell">${properties}</style:style>`);
  const { sourceStyleName: _name, ...style } = book.styles[book.cellStyleId("ce1")]!;
  return style;
}

const cell = (attrs: string): string => `<style:table-cell-properties ${attrs}/>`;
const paragraph = (attrs: string): string => `<style:paragraph-properties ${attrs}/>`;
const text = (attrs: string): string => `<style:text-properties ${attrs}/>`;

describe("the property table, row by row", () => {
  it("fo:background-color → a solid fill, lowercased; transparent → no fill", () => {
    expect(mapped(cell(`fo:background-color="#FFFF00"`))).toEqual({ fill: { kind: "solid", color: "#ffff00" } });
    expect(mapped(cell(`fo:background-color="transparent"`))).toEqual({});
  });

  it("fo:border and fo:border-{top,bottom,left,right} → borders, by the ODS borders table", () => {
    const thin = { style: "thin", color: "#000000" };
    expect(mapped(cell(`fo:border="0.74pt solid #000000"`))).toEqual({ borders: { top: thin, right: thin, bottom: thin, left: thin } });
    expect(
      mapped(cell(`fo:border-bottom="1.75pt dashed #0000FF" fo:border-left="none" fo:border-right="none" fo:border-top="2.5pt solid #000000"`)),
    ).toEqual({ borders: { top: { style: "thick", color: "#000000" }, bottom: { style: "mediumDashed", color: "#0000ff" } } });
    // A side the same element names wins over its fo:border.
    expect(mapped(cell(`fo:border="0.74pt solid #000000" fo:border-top="none"`))).toEqual({ borders: { right: thin, bottom: thin, left: thin } });
    expect(mapped(cell(`fo:border="none"`))).toEqual({});
  });

  it("style:diagonal-bl-tr / style:diagonal-tl-br → the diagonal with up / down", () => {
    expect(mapped(cell(`style:diagonal-bl-tr="0.74pt solid #ff0000" style:diagonal-tl-br="none"`))).toEqual({
      borders: { diagonal: { style: "thin", color: "#ff0000", up: true } },
    });
    expect(mapped(cell(`style:diagonal-tl-br="1.75pt solid #000000"`))).toEqual({
      borders: { diagonal: { style: "medium", color: "#000000", down: true } },
    });
    expect(mapped(cell(`style:diagonal-bl-tr="0.74pt dotted #000000" style:diagonal-tl-br="0.74pt dotted #000000"`))).toEqual({
      borders: { diagonal: { style: "dotted", color: "#000000", up: true, down: true } },
    });
  });

  it('fo:wrap-option="wrap" → alignment.wrap', () => {
    expect(mapped(cell(`fo:wrap-option="wrap"`))).toEqual({ alignment: { wrap: true } });
    expect(mapped(cell(`fo:wrap-option="no-wrap"`))).toEqual({});
  });

  it("style:vertical-align top / middle / bottom / automatic → top / center / bottom / absent", () => {
    expect(mapped(cell(`style:vertical-align="top"`))).toEqual({ alignment: { vertical: "top" } });
    expect(mapped(cell(`style:vertical-align="middle"`))).toEqual({ alignment: { vertical: "center" } });
    expect(mapped(cell(`style:vertical-align="bottom"`))).toEqual({ alignment: { vertical: "bottom" } });
    expect(mapped(cell(`style:vertical-align="automatic"`))).toEqual({});
  });

  it("style:rotation-angle a: a ≤ 90 → a; a ≥ 270 → 90 + (360 − a); 90 < a ≤ 180 → 90; else 180; 0 → none", () => {
    const cases: Array<[string, number | undefined]> = [
      ["0", undefined],
      ["45", 45],
      ["90", 90],
      ["91", 90],
      ["135", 90],
      ["180", 90],
      ["181", 180],
      ["269", 180],
      ["270", 180],
      ["315", 135],
      ["359", 91],
      ["360", undefined],
      ["-45", 135],
      ["90deg", 90],
      ["100grad", 90],
    ];
    for (const [angle, expected] of cases) {
      expect(odsTextRotation(angle), angle).toBe(expected);
    }
    expect(mapped(cell(`style:rotation-angle="90" style:rotation-align="none"`))).toEqual({ alignment: { textRotation: 90 } });
  });

  it('style:direction="ttb" → 255 (vertical stacked), whatever the rotation', () => {
    expect(mapped(cell(`style:direction="ttb"`))).toEqual({ alignment: { textRotation: 255 } });
    expect(mapped(cell(`style:direction="ttb" style:rotation-angle="45"`))).toEqual({ alignment: { textRotation: 255 } });
    expect(mapped(cell(`style:direction="ltr"`))).toEqual({});
  });

  it('style:shrink-to-fit="true" → alignment.shrinkToFit', () => {
    expect(mapped(cell(`style:shrink-to-fit="true"`))).toEqual({ alignment: { shrinkToFit: true } });
    expect(mapped(cell(`style:shrink-to-fit="false"`))).toEqual({});
  });

  it('style:repeat-content="true" → horizontal "fill" (over any fo:text-align)', () => {
    expect(mapped(cell(`style:text-align-source="fix" style:repeat-content="true"`) + paragraph(`fo:margin-left="0in"`))).toEqual({
      alignment: { horizontal: "fill" },
    });
    expect(mapped(cell(`style:text-align-source="fix" style:repeat-content="true"`) + paragraph(`fo:text-align="center"`))).toEqual({
      alignment: { horizontal: "fill" },
    });
  });

  it("style:cell-protect none / protected / formula-hidden / hidden-and-protected / protected formula-hidden", () => {
    expect(mapped(cell(`style:cell-protect="none"`))).toEqual({ protection: { locked: false } });
    expect(mapped(cell(`style:cell-protect="protected"`))).toEqual({});
    expect(mapped(cell(`style:cell-protect="formula-hidden"`))).toEqual({ protection: { hidden: true } });
    expect(mapped(cell(`style:cell-protect="hidden-and-protected"`))).toEqual({ protection: { hidden: true } });
    expect(mapped(cell(`style:cell-protect="protected formula-hidden"`))).toEqual({ protection: { hidden: true } });
  });

  it("fo:text-align start / end / center / justify (fix) → left / right / center / justify; value-type or none → general (absent)", () => {
    const fix = cell(`style:text-align-source="fix" style:repeat-content="false"`);
    expect(mapped(fix + paragraph(`fo:text-align="start"`))).toEqual({ alignment: { horizontal: "left" } });
    expect(mapped(fix + paragraph(`fo:text-align="end"`))).toEqual({ alignment: { horizontal: "right" } });
    expect(mapped(fix + paragraph(`fo:text-align="center"`))).toEqual({ alignment: { horizontal: "center" } });
    expect(mapped(fix + paragraph(`fo:text-align="justify"`))).toEqual({ alignment: { horizontal: "justify" } });
    expect(mapped(cell(`style:text-align-source="value-type"`) + paragraph(`fo:text-align="center"`))).toEqual({});
    expect(mapped(fix)).toEqual({});
  });

  it("fo:margin-left → indent = round(px / 9), in any ODF unit; 0 → none", () => {
    expect(mapped(paragraph(`fo:text-align="start" fo:margin-left="0.1874in"`))).toEqual({ alignment: { horizontal: "left", indent: 2 } });
    expect(mapped(paragraph(`fo:margin-left="9px"`))).toEqual({ alignment: { indent: 1 } });
    expect(mapped(paragraph(`fo:margin-left="0.5cm"`))).toEqual({ alignment: { indent: 2 } });
    expect(mapped(paragraph(`fo:margin="27pt"`))).toEqual({ alignment: { indent: 4 } });
    expect(mapped(paragraph(`fo:margin-left="4px"`))).toEqual({});
    expect(mapped(paragraph(`fo:margin-left="0in"`))).toEqual({});
  });

  it("text-properties → the font: size, bold, italic, underline, line-through, colour, position", () => {
    expect(mapped(text(`fo:font-size="12pt"`))).toEqual({ font: { sizePt: 12 } });
    expect(mapped(text(`fo:font-weight="bold" style:font-weight-asian="bold"`))).toEqual({ font: { bold: true } });
    expect(mapped(text(`fo:font-weight="700"`))).toEqual({ font: { bold: true } });
    expect(mapped(text(`fo:font-weight="normal"`))).toEqual({});
    expect(mapped(text(`fo:font-style="italic"`))).toEqual({ font: { italic: true } });
    expect(mapped(text(`style:text-underline-style="solid" style:text-underline-width="auto" style:text-underline-color="font-color"`))).toEqual({
      font: { underline: "single" },
    });
    expect(mapped(text(`style:text-underline-style="solid" style:text-underline-type="double"`))).toEqual({ font: { underline: "double" } });
    expect(mapped(text(`style:text-underline-style="none"`))).toEqual({});
    expect(mapped(text(`style:text-line-through-style="solid" style:text-line-through-type="single"`))).toEqual({ font: { strike: true } });
    expect(mapped(text(`style:text-line-through-style="none"`))).toEqual({});
    expect(mapped(text(`fo:color="#C9211E"`))).toEqual({ font: { color: "#c9211e" } });
    expect(mapped(text(`style:text-position="super 58%"`))).toEqual({ font: { vertAlign: "superscript" } });
    expect(mapped(text(`style:text-position="33% 58%"`))).toEqual({ font: { vertAlign: "superscript" } });
    expect(mapped(text(`style:text-position="sub 58%"`))).toEqual({ font: { vertAlign: "subscript" } });
    expect(mapped(text(`style:text-position="-33% 58%"`))).toEqual({ font: { vertAlign: "subscript" } });
    expect(mapped(text(`style:text-position="0% 100%"`))).toEqual({});
  });

  it("never stores an undefined-valued key", () => {
    const style = mapped(cell(`fo:background-color="transparent" fo:border="none" style:vertical-align="automatic"`) + text(`fo:font-weight="normal"`));
    expect(JSON.parse(JSON.stringify(style))).toEqual(style);
    expect(Object.keys(style)).toEqual([]);
  });
});

describe("ODS borders: the read table", () => {
  const widths: Array<[string, 0 | 1 | 2]> = [
    ["0.74pt", 0],
    ["1pt", 0],
    ["1.01pt", 1],
    ["1.75pt", 1],
    ["2pt", 1],
    ["2.01pt", 2],
    ["2.5pt", 2],
  ];
  const table: Array<[string, [SheetBorderStyle, SheetBorderStyle, SheetBorderStyle]]> = [
    ["solid", ["thin", "medium", "thick"]],
    ["dashed", ["dashed", "mediumDashed", "mediumDashed"]],
    ["dotted", ["dotted", "dotted", "dotted"]],
    ["fine-dashed", ["hair", "hair", "hair"]],
    ["dash-dot", ["dashDot", "mediumDashDot", "mediumDashDot"]],
    ["dash-dot-dot", ["dashDotDot", "mediumDashDotDot", "mediumDashDotDot"]],
    ["double", ["double", "double", "double"]],
    ["double-thin", ["double", "double", "double"]],
  ];

  for (const [line, tiers] of table) {
    it(`${line}: w ≤ 1.0 pt → ${tiers[0]}, 1.0 < w ≤ 2.0 → ${tiers[1]}, w > 2.0 → ${tiers[2]}`, () => {
      for (const [width, tier] of widths) {
        expect(odsBorderEdge(`${width} ${line} #000000`), `${width} ${line}`).toEqual({ style: tiers[tier], color: "#000000" });
      }
    });
  }

  it("none and hidden draw no edge", () => {
    expect(odsBorderEdge("none")).toBeUndefined();
    expect(odsBorderEdge("0.74pt none #000000")).toBeUndefined();
    expect(odsBorderEdge("0.74pt hidden #000000")).toBeUndefined();
    expect(odsBorderEdge("0pt solid #000000")).toBeUndefined();
  });

  it("reads the width in any ODF length unit", () => {
    expect(odsBorderEdge("0.0104in solid #000000")?.style).toBe("thin"); // 0.7488 pt
    expect(odsBorderEdge("0.035cm solid #000000")?.style).toBe("thin"); // 0.992 pt
    expect(odsBorderEdge("0.06cm solid #000000")?.style).toBe("medium"); // 1.70 pt
    expect(odsBorderEdge("0.8mm solid #000000")?.style).toBe("thick"); // 2.27 pt
    expect(odsBorderEdge("1px solid #000000")?.style).toBe("thin"); // 0.75 pt
    expect(odsBorderEdge("3px solid #000000")?.style).toBe("thick"); // 2.25 pt
    expect(odsBorderEdge("0.1pc solid #000000")?.style).toBe("medium"); // 1.2 pt
    expect(odsBorderEdge("0.0278in solid #000000")?.style).toBe("thick"); // 2.0016 pt
  });

  it("keeps a written colour (lowercased), in any token order; none written → automatic", () => {
    expect(odsBorderEdge("solid #FF0000 1.75pt")).toEqual({ style: "medium", color: "#ff0000" });
    expect(odsBorderEdge("0.74pt solid")).toEqual({ style: "thin" });
  });
});

describe("ODS borders: the write table (pinned for SPEC-122)", () => {
  const written: Record<SheetBorderStyle, string> = {
    thin: "0.74pt solid",
    medium: "1.75pt solid",
    thick: "2.5pt solid",
    dashed: "0.74pt dashed",
    mediumDashed: "1.75pt dashed",
    dotted: "0.74pt dotted",
    hair: "0.74pt fine-dashed",
    dashDot: "0.74pt dash-dot",
    mediumDashDot: "1.75pt dash-dot",
    slantDashDot: "1.75pt dash-dot",
    dashDotDot: "0.74pt dash-dot-dot",
    mediumDashDotDot: "1.75pt dash-dot-dot",
    double: "2.6pt double",
  };

  it("writes each style's width and line, then the colour — #000000 when automatic", () => {
    expect(ODS_BORDER_WRITE).toEqual(written);
    for (const [style, value] of Object.entries(written) as [SheetBorderStyle, string][]) {
      expect(odsBorderString({ style })).toBe(`${value} #000000`);
      expect(odsBorderString({ style, color: "#c9211e" })).toBe(`${value} #c9211e`);
    }
  });

  it("resolves a theme colour (an edge that came from an xlsx)", () => {
    expect(odsBorderString({ style: "thin", color: { theme: "accent1" } })).toBe(`0.74pt solid ${DEFAULT_THEME.colors.accent1}`);
  });

  it("reads every written style back as itself — but slantDashDot, which reads back as mediumDashDot", () => {
    for (const style of Object.keys(written) as SheetBorderStyle[]) {
      const back = odsBorderEdge(odsBorderString({ style, color: "#123456" }));
      expect(back, style).toEqual({ style: style === "slantDashDot" ? "mediumDashDot" : style, color: "#123456" });
    }
  });
});

describe("lengths", () => {
  it("converts every ODF unit exactly, at 96 dpi", () => {
    expect(odsLengthToPt("0.178in")).toBe(12.816);
    expect(odsLengthToPt("0.3in")).toBe(21.6);
    expect(odsLengthToPt("12pt")).toBe(12);
    expect(odsLengthToPt("1pc")).toBe(12);
    expect(odsLengthToPt("10px")).toBe(7.5);
    expect(odsLengthToPt("2.54cm")).toBe(72);
    expect(odsLengthToPt("25.4mm")).toBe(72);
    expect(odsLengthToPx("0.889in")).toBe(85.344);
    expect(odsLengthToPx("1.2in")).toBe(115.2);
    expect(odsLengthToPx("12pt")).toBe(16);
    expect(odsLengthToPx("2.54cm")).toBe(96);
    expect(odsLengthToPx(".5in")).toBe(48);
    expect(odsLengthToPx("0")).toBe(0);
  });

  it("is undefined for anything that is not an absolute length", () => {
    for (const value of [undefined, "", "abc", "12", "12em", "50%", "in"]) {
      expect(odsLengthToPt(value), String(value)).toBeUndefined();
    }
  });
});

describe("flattening along the parent chain", () => {
  it("walks the default style → common ancestors → the named style, later definitions winning per property", () => {
    const book = styleBook(
      `<style:style style:name="ce1" style:family="table-cell" style:parent-style-name="Heading">` +
        `<style:table-cell-properties fo:background-color="#ffff00"/><style:text-properties fo:font-weight="normal" fo:color="#ff0000"/></style:style>`,
      LO_COMMON +
        `<style:style style:name="Accent" style:family="table-cell" style:parent-style-name="Heading">` +
        `<style:text-properties fo:font-style="italic"/></style:style>`,
      LO_FONT_FACES,
    );
    // Heading's 24pt beats the default style's 10pt; ce1's weight beats Heading's bold.
    expect(book.styles[book.cellStyleId("ce1")]).toEqual({
      font: { family: "Liberation Sans", sizePt: 24, color: "#ff0000" },
      fill: { kind: "solid", color: "#ffff00" },
      sourceStyleName: "ce1",
    });
    // A common style named directly is flattened the same way.
    expect(book.styles[book.cellStyleId("Accent")]).toEqual({
      font: { family: "Liberation Sans", sizePt: 24, bold: true, italic: true },
      sourceStyleName: "Accent",
    });
  });

  it("an fo:border shorthand counts as its four sides when merging", () => {
    const book = styleBook(
      `<style:style style:name="ce1" style:family="table-cell" style:parent-style-name="Boxed">${cell(`fo:border-left="none"`)}</style:style>` +
        `<style:style style:name="ce2" style:family="table-cell" style:parent-style-name="TopOnly">${cell(`fo:border="0.74pt dotted #000000"`)}</style:style>`,
      `<style:style style:name="Boxed" style:family="table-cell">${cell(`fo:border="2.5pt solid #000000"`)}</style:style>` +
        `<style:style style:name="TopOnly" style:family="table-cell">${cell(`fo:border-top="2.5pt solid #000000"`)}</style:style>`,
    );
    const thick = { style: "thick", color: "#000000" };
    expect(book.styles[book.cellStyleId("ce1")]!.borders).toEqual({ top: thick, right: thick, bottom: thick });
    const dotted = { style: "dotted", color: "#000000" };
    expect(book.styles[book.cellStyleId("ce2")]!.borders).toEqual({ top: dotted, right: dotted, bottom: dotted, left: dotted });
  });

  it("a percentage font size scales the inherited size", () => {
    const book = styleBook(`<style:style style:name="ce1" style:family="table-cell" style:parent-style-name="Heading">${text(`fo:font-size="50%"`)}</style:style>`, LO_COMMON);
    expect(book.styles[book.cellStyleId("ce1")]!.font?.sizePt).toBe(12);
  });

  it("survives a parent cycle and an undefined parent", () => {
    const book = styleBook(
      `<style:style style:name="ce1" style:family="table-cell" style:parent-style-name="ce2">${text(`fo:font-weight="bold"`)}</style:style>` +
        `<style:style style:name="ce2" style:family="table-cell" style:parent-style-name="ce1">${text(`fo:font-style="italic"`)}</style:style>` +
        `<style:style style:name="ce3" style:family="table-cell" style:parent-style-name="Nowhere">${text(`fo:font-style="italic"`)}</style:style>`,
    );
    expect(book.styles[book.cellStyleId("ce1")]!.font).toEqual({ bold: true, italic: true });
    expect(book.styles[book.cellStyleId("ce3")]!.font).toEqual({ italic: true });
  });
});

describe("font faces", () => {
  it("resolve style:font-name through office:font-face-decls, quotes stripped; content.xml's declarations win", () => {
    const automatic =
      `<style:style style:name="ce1" style:family="table-cell">${text(`style:font-name="Liberation Sans"`)}</style:style>` +
      `<style:style style:name="ce2" style:family="table-cell">${text(`style:font-name="Body1"`)}</style:style>` +
      `<style:style style:name="ce3" style:family="table-cell">${text(`style:font-name="Undeclared"`)}</style:style>` +
      `<style:style style:name="ce4" style:family="table-cell">${text(`fo:font-family="'Courier New', monospace"`)}</style:style>`;
    const contentFaces =
      `<office:font-face-decls><style:font-face style:name="Body1" svg:font-family="&apos;DejaVu Sans&apos;, sans-serif"/>` +
      `<style:font-face style:name="Liberation Sans" svg:font-family="Liberation Sans Narrow"/></office:font-face-decls>`;
    const book = styleBook(automatic, "", LO_FONT_FACES, contentFaces);
    const family = (name: string): unknown => book.styles[book.cellStyleId(name)]!.font?.family;
    expect(family("ce1")).toBe("Liberation Sans Narrow");
    expect(family("ce2")).toBe("DejaVu Sans");
    expect(family("ce3")).toBe("Undeclared");
    expect(family("ce4")).toBe("Courier New");
    expect(styleBook(automatic, "", LO_FONT_FACES).styles[0]).toEqual({});
    const fromStyles = styleBook(automatic, "", LO_FONT_FACES);
    expect(fromStyles.styles[fromStyles.cellStyleId("ce1")]!.font?.family).toBe("Liberation Sans");
  });
});

describe("the style table", () => {
  it("styles[0] is the flattened Default style; no styles.xml → a default-only table", () => {
    const book = styleBook("", LO_COMMON, LO_FONT_FACES);
    expect(book.styles).toEqual([{ font: { family: "Liberation Sans", sizePt: 10 }, sourceStyleName: "Default" }]);
    expect(book.cellStyleId("Default")).toBe(0);
    expect(new OdsStyleBook(emptyOdsStyleSheet()).styles).toEqual([{}]);
  });

  it("gives each referenced name one entry, in first-reference order, with sourceStyleName — never deduplicated", () => {
    const book = styleBook(
      `<style:style style:name="ce1" style:family="table-cell" style:parent-style-name="Default">${text(`fo:font-weight="bold"`)}</style:style>` +
        `<style:style style:name="ce2" style:family="table-cell" style:parent-style-name="Default">${text(`fo:font-weight="bold"`)}</style:style>`,
      LO_COMMON,
    );
    expect(book.cellStyleId("ce2")).toBe(1);
    expect(book.cellStyleId("ce1")).toBe(2);
    expect(book.cellStyleId("ce2")).toBe(1);
    expect(book.cellStyleId("Heading")).toBe(3);
    expect(book.cellStyleId("ce9")).toBe(4); // a name nothing defines: the default style's properties, under its name
    expect(book.styles.map((style) => style.sourceStyleName)).toEqual(["Default", "ce2", "ce1", "Heading", "ce9"]);
    expect(book.styles[1]).toEqual({ font: { family: "Liberation Sans", sizePt: 10, bold: true }, sourceStyleName: "ce2" });
    expect(book.styles[4]).toEqual({ font: { family: "Liberation Sans", sizePt: 10 }, sourceStyleName: "ce9" });
    expect(book.styles.every((style) => style.namedStyle === undefined && style.numberFormat === undefined)).toBe(true);
  });

  it("a text:span's font is its text style's own properties, never the cell defaults", () => {
    const book = styleBook(
      `<style:style style:name="T1" style:family="text"><style:text-properties fo:color="#c9211e" fo:font-weight="bold" style:font-weight-asian="bold"/></style:style>` +
        `<style:style style:name="T2" style:family="text"/>`,
      LO_COMMON,
    );
    expect(book.textFont("T1")).toEqual({ bold: true, color: "#c9211e" });
    expect(book.textFont("T2")).toBeUndefined();
    expect(book.textFont("T9")).toBeUndefined();
  });

  it("reads a table style's display and tab colour", () => {
    const book = styleBook(
      `<style:style style:name="ta1" style:family="table"><style:table-properties table:display="true" tableooo:tab-color="#729FCF"/></style:style>` +
        `<style:style style:name="ta2" style:family="table"><style:table-properties table:display="false"/></style:style>`,
    );
    expect(book.tableStyle("ta1")).toEqual({ hidden: false, tabColor: "#729fcf" });
    expect(book.tableStyle("ta2")).toEqual({ hidden: true });
    expect(book.tableStyle(undefined)).toEqual({ hidden: false });
  });
});

/** The fixture, and the style each A1 position of "Styled" shows. */
async function read() {
  const parsed = await parseOds(await buildSpec117StylesOdsBytes());
  const styled = parsed.sheets[0]!;
  const at = (a1: string): CellStyle => {
    const ref = parseA1(a1)!;
    return cellStyle(parsed, effectiveStyleId(styled, ref.row, ref.col));
  };
  return { parsed, at };
}

describe("number formats: style:data-style-name (SPEC-118)", () => {
  /** LibreOffice's two-section red-negative style, as Calc writes it. */
  const N2 =
    `<number:number-style style:name="N2P0" style:volatile="true"><number:number number:decimal-places="2" number:min-decimal-places="2" number:min-integer-digits="1"/></number:number-style>` +
    `<number:number-style style:name="N2"><style:text-properties fo:color="#ff0000"/><number:text>-</number:text>` +
    `<number:number number:decimal-places="2" number:min-decimal-places="2" number:min-integer-digits="1"/>` +
    `<style:map style:condition="value()&gt;=0" style:apply-style-name="N2P0"/></number:number-style>`;

  it("maps an automatic cell style's data style to numberFormat (LibreOffice's sign pattern as plain sections)", () => {
    const book = styleBook(`${N2}<style:style style:name="ce1" style:family="table-cell" style:parent-style-name="Default" style:data-style-name="N2"/>`, LO_COMMON);
    expect(book.styles[book.cellStyleId("ce1")]!.numberFormat).toBe("0.00;[Red]-0.00");
  });

  it("finds a data style in styles.xml's common styles, and inherits the name down the parent chain", () => {
    const common =
      `${LO_COMMON}<number:percentage-style style:name="N10"><number:number number:decimal-places="2" number:min-decimal-places="2" number:min-integer-digits="1"/>` +
      `<number:text>%</number:text></number:percentage-style>` +
      `<style:style style:name="Percent" style:family="table-cell" style:parent-style-name="Default" style:data-style-name="N10"/>`;
    const book = styleBook(`<style:style style:name="ce1" style:family="table-cell" style:parent-style-name="Percent"><style:text-properties fo:font-weight="bold"/></style:style>`, common);
    expect(book.styles[book.cellStyleId("ce1")]).toMatchObject({ numberFormat: "0.00%", font: { bold: true } });
    expect(book.styles[book.cellStyleId("Percent")]!.numberFormat).toBe("0.00%");
    // The nearest definition wins.
    const own = styleBook(
      `${N2}<style:style style:name="ce2" style:family="table-cell" style:parent-style-name="Percent" style:data-style-name="N2"/>`,
      common,
    );
    expect(own.styles[own.cellStyleId("ce2")]!.numberFormat).toBe("0.00;[Red]-0.00");
  });

  it("leaves General absent: no data style, a General number style, an era style, a name no part defines", () => {
    const automatic =
      `<number:number-style style:name="N0"><number:number number:min-integer-digits="1"/></number:number-style>` +
      `<number:date-style style:name="N80"><number:era number:style="long"/><number:text> </number:text><number:year/></number:date-style>` +
      `<style:style style:name="plain" style:family="table-cell"/>` +
      `<style:style style:name="general" style:family="table-cell" style:data-style-name="N0"/>` +
      `<style:style style:name="era" style:family="table-cell" style:data-style-name="N80"/>` +
      `<style:style style:name="missing" style:family="table-cell" style:data-style-name="N404"/>`;
    const book = styleBook(automatic, LO_COMMON);
    for (const name of ["plain", "general", "era", "missing"]) {
      expect(book.styles[book.cellStyleId(name)]!.numberFormat, name).toBeUndefined();
    }
  });

  it("derives one style per (base, implied code) for a value whose style names no data style — and none for one that does", () => {
    const automatic =
      `${N2}<style:style style:name="ce1" style:family="table-cell"><style:text-properties fo:font-weight="bold"/></style:style>` +
      `<style:style style:name="ce2" style:family="table-cell" style:data-style-name="N2"/>` +
      `<number:date-style style:name="N80"><number:era/></number:date-style>` +
      `<style:style style:name="era" style:family="table-cell" style:data-style-name="N80"/>`;
    const book = styleBook(automatic, LO_COMMON);
    const ce1 = book.cellStyleId("ce1");
    const derived = book.impliedStyleId(ce1, "yyyy-mm-dd");
    expect(derived).toBe(book.styles.length - 1);
    expect(book.styles[derived]).toEqual({ font: { family: "Liberation Sans", sizePt: 10, bold: true }, numberFormat: "yyyy-mm-dd" });
    expect(book.impliedStyleId(ce1, "yyyy-mm-dd")).toBe(derived);
    expect(book.impliedStyleId(ce1, "[hh]:mm:ss")).toBe(derived + 1);
    expect(book.impliedStyleId(0, "0.00%")).toBe(derived + 2);
    expect(book.styles[derived + 2]!.sourceStyleName).toBeUndefined();
    const ce2 = book.cellStyleId("ce2");
    expect(book.impliedStyleId(ce2, "yyyy-mm-dd")).toBe(ce2);
    const era = book.cellStyleId("era");
    expect(book.impliedStyleId(era, "yyyy-mm-dd")).toBe(era); // it names a data style, if one that reads as General
  });
});

describe("spec117-styles.ods (LibreOffice-shaped)", () => {

  it("B2 is bold red text on yellow, in the default font", async () => {
    const { at } = await read();
    const { header } = SPEC117_STYLES_ODS.styled;
    expect(at(header.a1)).toEqual({
      font: { family: "Liberation Sans", sizePt: 10, bold: true, color: header.color },
      fill: { kind: "solid", color: header.fill },
      sourceStyleName: header.styleName,
    });
  });

  it("styles[0] is LibreOffice's Default: Liberation Sans 10pt", async () => {
    const { parsed, at } = await read();
    expect(parsed.styles[0]).toEqual({ font: { family: "Liberation Sans", sizePt: 10 }, sourceStyleName: "Default" });
    expect(at("B3")).toBe(parsed.styles[0]);
  });

  it("the merged title flattens the common Heading style and centres", async () => {
    const { at } = await read();
    const { merge, titleFont } = SPEC117_STYLES_ODS.styled;
    expect(at(merge.anchor)).toEqual({ font: titleFont, alignment: { horizontal: "center" }, sourceStyleName: merge.styleName });
  });

  it("reads the six border cells, the diagonal and every alignment cell", async () => {
    const { at } = await read();
    const { borders, diagonal, alignment } = SPEC117_STYLES_ODS.styled;
    for (const [a1, style] of Object.entries(borders)) {
      const edge = { style, color: "#000000" };
      expect(at(a1).borders, a1).toEqual({ top: edge, right: edge, bottom: edge, left: edge });
    }
    expect(at(diagonal.a1).borders).toEqual({ bottom: diagonal.bottom, diagonal: diagonal.diagonal });
    for (const [a1, expected] of Object.entries(alignment)) {
      expect(at(a1).alignment, a1).toEqual(expected);
    }
  });

  it("gives every referenced name one entry, in first-reference order", async () => {
    const { parsed } = await read();
    expect(parsed.styles.map((style) => style.sourceStyleName)).toEqual([
      "Default",
      "ce9", // column A's default cell style comes first: columns precede rows
      "ce10",
      "ce1",
      "ce12",
      "ce16",
      "ce2",
      "ce8",
      "ce3",
      "ce11",
      "ce4",
      "ce13",
      "ce5",
      "ce14",
      "ce6",
      "ce15",
      "ce7",
    ]);
  });
});

// --- SPEC-122: the generated styles -------------------------------------------------------------

const ENV: OdsStyleEnv = { theme: DEFAULT_THEME, locale: SPREADSHEET_LOCALES["en-US"] };

/** A style generated by the registry as a de-ce style, then read back as
 * the reader reads content.xml's automatic styles (sourceStyleName
 * dropped). */
function writtenAndRead(style: CellStyle): Omit<CellStyle, "sourceStyleName"> {
  const registry = new OdsStyleRegistry();
  const name = registry.cellStyle(style, ENV);
  const book = styleBook(registry.xml());
  const { sourceStyleName: _name, ...back } = book.styles[book.cellStyleId(name)]!;
  return back;
}

const BORDER_STYLES = (Object.keys(ODS_BORDER_WRITE) as SheetBorderStyle[]).filter((style) => style !== "slantDashDot");

/** Every row of the property table, as styles that read back as themselves. */
const INVERSE: [string, CellStyle][] = [
  ["font family, quoted for a space", { font: { family: "Liberation Sans" } }],
  ["font family", { font: { family: "Arial" } }],
  ["font size", { font: { sizePt: 10.5 } }],
  ["bold", { font: { bold: true } }],
  ["italic", { font: { italic: true } }],
  ["single underline", { font: { underline: "single" } }],
  ["double underline", { font: { underline: "double" } }],
  ["strike", { font: { strike: true } }],
  ["colour", { font: { color: "#c9211e" } }],
  ["superscript", { font: { vertAlign: "superscript" } }],
  ["subscript", { font: { vertAlign: "subscript" } }],
  ["a whole font", { font: { family: "Liberation Serif", sizePt: 12, bold: true, italic: true, underline: "single", strike: true, color: "#0000ff", vertAlign: "subscript" } }],
  ["solid fill", { fill: { kind: "solid", color: "#ffff00" } }],
  ...BORDER_STYLES.map((style): [string, CellStyle] => {
    const edge = { style, color: "#123456" };
    return [`${style} borders`, { borders: { top: edge, right: edge, bottom: edge, left: edge } }];
  }),
  ["one border", { borders: { bottom: { style: "double", color: "#000000" } } }],
  ["diagonal down", { borders: { diagonal: { style: "thin", color: "#000000", down: true } } }],
  ["diagonal up", { borders: { diagonal: { style: "medium", color: "#ff0000", up: true } } }],
  ["both diagonals", { borders: { diagonal: { style: "dashed", color: "#00ff00", up: true, down: true } } }],
  ...(["left", "center", "right", "justify", "fill"] as const).map((horizontal): [string, CellStyle] => [`horizontal ${horizontal}`, { alignment: { horizontal } }]),
  ...(["top", "center", "bottom"] as const).map((vertical): [string, CellStyle] => [`vertical ${vertical}`, { alignment: { vertical } }]),
  ["wrap", { alignment: { wrap: true } }],
  ["indent 1", { alignment: { horizontal: "left", indent: 1 } }],
  ["indent 3", { alignment: { indent: 3 } }],
  ...[1, 45, 90, 91, 135, 180, 255].map((textRotation): [string, CellStyle] => [`rotation ${textRotation}`, { alignment: { textRotation } }]),
  ["shrink", { alignment: { shrinkToFit: true } }],
  ["unlocked", { protection: { locked: false } }],
  ["hidden", { protection: { hidden: true } }],
  ["number format", { numberFormat: "#,##0.00" }],
  ["percent format", { numberFormat: "0.00%" }],
  ["date format", { numberFormat: "yyyy-mm-dd" }],
  ["sections", { numberFormat: "#,##0.00;[Red]-#,##0.00" }],
];

describe("SPEC-122 generated cell styles: the property table's inverse", () => {
  for (const [label, style] of INVERSE) {
    it(`${label} reads back equal`, () => {
      expect(writtenAndRead(style)).toEqual(style);
    });
  }

  it("writes the table's exact properties, the property elements in the schema's order", () => {
    const registry = new OdsStyleRegistry();
    registry.cellStyle(
      {
        font: { family: "Liberation Sans", sizePt: 10, bold: true, color: "#c9211e" },
        fill: { kind: "solid", color: "#ffff00" },
        borders: { top: { style: "thin" }, left: { style: "thick", color: "#ff0000" }, diagonal: { style: "hair", color: "#00ff00", up: true } },
        alignment: { horizontal: "center", vertical: "center", wrap: true, indent: 1, textRotation: 135, shrinkToFit: true },
        protection: { locked: false },
      },
      ENV,
    );
    registry.cellStyle({ alignment: { horizontal: "fill", textRotation: 255 }, protection: { hidden: true }, font: { underline: "double", strike: true, vertAlign: "superscript" } }, ENV);
    registry.cellStyle({ alignment: { horizontal: "general" } }, ENV);
    expect(registry.xml()).toBe(
      `<style:style style:name="de-ce1" style:family="table-cell" style:parent-style-name="Default">` +
        `<style:table-cell-properties style:text-align-source="fix" style:repeat-content="false" fo:background-color="#ffff00" fo:border-top="0.74pt solid #000000" ` +
        `fo:border-left="2.5pt solid #ff0000" style:diagonal-bl-tr="0.74pt fine-dashed #00ff00" fo:wrap-option="wrap" style:vertical-align="middle" ` +
        `style:rotation-angle="315" style:shrink-to-fit="true" style:cell-protect="none"/>` +
        `<style:paragraph-properties fo:text-align="center" fo:margin-left="6.75pt"/>` +
        `<style:text-properties fo:font-family="'Liberation Sans'" fo:font-size="10pt" fo:font-weight="bold" fo:color="#c9211e"/></style:style>` +
        `<style:style style:name="de-ce2" style:family="table-cell" style:parent-style-name="Default">` +
        `<style:table-cell-properties style:repeat-content="true" style:direction="ttb" style:cell-protect="hidden-and-protected"/>` +
        `<style:text-properties style:text-underline-style="solid" style:text-underline-type="double" style:text-line-through-style="solid" style:text-position="super 58%"/></style:style>` +
        `<style:style style:name="de-ce3" style:family="table-cell" style:parent-style-name="Default"><style:table-cell-properties style:text-align-source="value-type"/></style:style>`,
    );
  });

  it("approximates what ODS cannot say: a pattern or gradient fill as its first colour, an automatic colour as #000000, theme colours and fonts resolved", () => {
    expect(writtenAndRead({ fill: { kind: "pattern", pattern: "gray125", fg: "#ff0000", bg: "#00ff00" } })).toEqual({ fill: { kind: "solid", color: "#ff0000" } });
    expect(writtenAndRead({ fill: { kind: "pattern", pattern: "gray125", bg: "#00ff00" } })).toEqual({ fill: { kind: "solid", color: "#00ff00" } });
    expect(
      writtenAndRead({ fill: { kind: "gradient", type: "linear", degree: 90, stops: [{ position: 0, color: "#112233" }, { position: 1, color: "#445566" }] } }),
    ).toEqual({ fill: { kind: "solid", color: "#112233" } });
    expect(writtenAndRead({ borders: { top: { style: "thin" } } })).toEqual({ borders: { top: { style: "thin", color: "#000000" } } });
    expect(writtenAndRead({ font: { family: { theme: "minor" }, color: { theme: "accent1" } } })).toEqual({
      font: { family: resolveFont({ theme: "minor" }, DEFAULT_THEME), color: resolveColor({ theme: "accent1" }, DEFAULT_THEME).toLowerCase() },
    });
    expect(writtenAndRead({ alignment: { horizontal: "general" } })).toEqual({});
    expect(writtenAndRead({ alignment: { horizontal: "centerContinuous", vertical: "justify" } })).toEqual({ alignment: { horizontal: "center", vertical: "top" } });
    expect(writtenAndRead({ alignment: { horizontal: "distributed", vertical: "distributed" } })).toEqual({ alignment: { horizontal: "justify", vertical: "top" } });
    expect(writtenAndRead({ font: { underline: "doubleAccounting" } })).toEqual({ font: { underline: "double" } });
  });

  it('writes a hidden unlocked cell as formula-hidden, which SPEC-117 reads as hidden (and locked) — the stated loss', () => {
    const registry = new OdsStyleRegistry();
    registry.cellStyle({ protection: { locked: false, hidden: true } }, ENV);
    expect(registry.xml()).toContain(`<style:table-cell-properties style:cell-protect="formula-hidden"/>`);
    expect(writtenAndRead({ protection: { locked: false, hidden: true } })).toEqual({ protection: { hidden: true } });
  });
});

describe("SPEC-122 generated styles: names and sharing", () => {
  it("names de-ce, de-T and de-N styles past every de- name already in the package — a file holding de-ce1 gets de-ce2", () => {
    const registry = new OdsStyleRegistry([
      `<style:style style:name="de-ce1" style:family="table-cell"/><style:style style:name="de-T2" style:family="text"/>`,
      `<number:number-style style:name="de-N4P0"/>`,
    ]);
    expect(registry.cellStyle({ font: { bold: true }, numberFormat: "0.0" }, ENV)).toBe("de-ce2");
    expect(registry.textStyle({ italic: true }, DEFAULT_THEME)).toBe("de-T3");
    expect(registry.xml()).toBe(
      `${odsNumberStyleXml("0.0", "de-N5")}<style:style style:name="de-ce2" style:family="table-cell" style:parent-style-name="Default" style:data-style-name="de-N5">` +
        `<style:text-properties fo:font-weight="bold"/></style:style>` +
        `<style:style style:name="de-T3" style:family="text"><style:text-properties fo:font-style="italic"/></style:style>`,
    );
  });

  it("writes number styles as odsNumberStyleXml(code, de-N<n>) — its P0… sub-styles included — before the cell style naming them; General names none", () => {
    const code = "#,##0.00;[Red]-#,##0.00";
    const registry = new OdsStyleRegistry();
    expect(registry.cellStyle({ numberFormat: code }, ENV)).toBe("de-ce1");
    expect(registry.cellStyle({ numberFormat: "General", font: { bold: true } }, ENV)).toBe("de-ce2");
    const numberStyle = odsNumberStyleXml(code, "de-N1");
    expect(numberStyle).toContain(`style:name="de-N1P0"`);
    expect(registry.xml()).toBe(
      `${numberStyle}<style:style style:name="de-ce1" style:family="table-cell" style:parent-style-name="Default" style:data-style-name="de-N1"/>` +
        `<style:style style:name="de-ce2" style:family="table-cell" style:parent-style-name="Default"><style:text-properties fo:font-weight="bold"/></style:style>`,
    );
  });

  it("writes a builtinFormatId style with that built-in's code in the user's locale — ODS has no locale-dependent ids", () => {
    const style: CellStyle = { builtinFormatId: 14, numberFormat: "m/d/yyyy" };
    const codes = new Set<string>();
    for (const locale of Object.values(SPREADSHEET_LOCALES)) {
      const registry = new OdsStyleRegistry();
      registry.cellStyle(style, { theme: DEFAULT_THEME, locale });
      const code = formatCodeOf(style, locale)!;
      codes.add(code);
      expect(registry.xml(), locale.tag).toBe(
        `${odsNumberStyleXml(code, "de-N1")}<style:style style:name="de-ce1" style:family="table-cell" style:parent-style-name="Default" style:data-style-name="de-N1"/>`,
      );
    }
    expect(codes.size).toBeGreaterThan(1);
  });

  it("shares one name between equal property sets, whatever a style was read from; equal codes share one number style", () => {
    const registry = new OdsStyleRegistry();
    const bold = registry.cellStyle({ font: { bold: true }, numberFormat: "0.0" }, ENV);
    expect(registry.cellStyle({ numberFormat: "0.0", font: { bold: true }, sourceXf: 4, sourceStyleName: "ce9" }, ENV)).toBe(bold);
    const italic = registry.cellStyle({ font: { italic: true }, numberFormat: "0.0" }, ENV);
    expect(italic).not.toBe(bold);
    expect(registry.cellStyle({ font: { bold: true }, numberFormat: "0.00" }, ENV)).not.toBe(bold);
    expect(registry.textStyle({ bold: true, color: "#ff0000" }, DEFAULT_THEME)).toBe(registry.textStyle({ color: "#FF0000", bold: true }, DEFAULT_THEME));
    expect(registry.textStyle({}, DEFAULT_THEME)).toBeUndefined();
    expect(registry.textStyle(undefined, DEFAULT_THEME)).toBeUndefined();
    const xml = registry.xml();
    expect([...xml.matchAll(/<number:[\w-]+-style style:name="(de-N\d+)"/g)].map((match) => match[1])).toEqual(["de-N1", "de-N2"]);
    expect([...xml.matchAll(/<style:style style:name="(de-[a-zA-Z]+\d+)"/g)].map((match) => match[1])).toEqual(["de-ce1", "de-ce2", "de-ce3", "de-T1"]);
  });

  it("parents a generated style on its namedStyle when the package defines that common style, else on Default", () => {
    const env: OdsStyleEnv = { ...ENV, commonCellStyle: (name) => name === "Heading" };
    const registry = new OdsStyleRegistry();
    registry.cellStyle({ namedStyle: "Heading", font: { bold: true } }, env);
    registry.cellStyle({ namedStyle: "Good", font: { italic: true } }, env);
    registry.cellStyle({ namedStyle: "Normal", font: { strike: true } }, env);
    expect([...registry.xml().matchAll(/style:parent-style-name="([^"]*)"/g)].map((match) => match[1])).toEqual(["Heading", "Default", "Default"]);
  });

  it("writes a rich run in the file's own text style when one gives its font — the first such, never one holding more than a run font", () => {
    const book = styleBook(
      `<style:style style:name="T1" style:family="text"><style:text-properties fo:font-weight="bold" style:font-weight-asian="bold"/></style:style>` +
        `<style:style style:name="T2" style:family="text"><style:text-properties fo:font-style="italic" fo:background-color="#ffff00"/></style:style>` +
        `<style:style style:name="T3" style:family="text"><style:text-properties fo:font-weight="bold"/></style:style>` +
        `<style:style style:name="ce1" style:family="table-cell"><style:text-properties fo:font-weight="bold"/></style:style>`,
      `<style:style style:name="Strong" style:family="text"><style:text-properties fo:font-style="italic"/></style:style>`,
    );
    const runStyles: { name: string; font: CellFont }[] = [
      { name: "T1", font: { bold: true } },
      { name: "T3", font: { bold: true } },
      { name: "Strong", font: { italic: true } },
    ];
    expect(book.runTextStyles()).toEqual(runStyles);
    const registry = new OdsStyleRegistry();
    for (const { name, font } of book.runTextStyles()) {
      registry.adoptTextStyle(name, font, DEFAULT_THEME);
    }
    expect(registry.textStyle({ bold: true }, DEFAULT_THEME)).toBe("T1");
    expect(registry.textStyle({ italic: true }, DEFAULT_THEME)).toBe("Strong");
    expect(registry.textStyle({ strike: true }, DEFAULT_THEME)).toBe("de-T1");
    expect(registry.xml()).toBe(`<style:style style:name="de-T1" style:family="text"><style:text-properties style:text-line-through-style="solid"/></style:style>`);
  });
});
