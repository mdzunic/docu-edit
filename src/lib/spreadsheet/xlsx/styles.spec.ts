// SPEC-117 "Reading xlsx styles" and "Colours": the colour codec
// (xlsx/themeColor.ts) — Excel's theme-index swap, the tint ↔ lumMod/lumOff
// arithmetic both ways, the indexed palette with its override, 64, 65, rgb
// ARGB and 6-digit values, auto — with the pinned "Accent 1, Lighter 40%" =
// Excel's #8ea9db; and the style table (xlsx/styles.ts) row by row: fonts,
// fills (pattern and gradient), borders (all thirteen styles, diagonals,
// start/end), cellXfs one-to-one with copies, alignment, protection,
// quotePrefix, the override palette, the parts left unread, the default
// font — over synthetic parts and spec117-styles.xlsx.
//
// SPEC-122 "Styles" › xlsx: the XlsxStyleWriter — a style with sourceXf
// writes that index; a new one reuses equal source fonts, fills, borders and
// number formats, else an earlier append, else appends; the numFmtId rules
// (builtinFormatId as the id; a locale-independent built-in found by its
// code, never 14/22 or another LOCALE_DEPENDENT_BUILTINS id; an existing
// custom id; custom ids from 164 or past the highest); the xf (xfId of the
// named style, apply* against that parent, quotePrefix, alignment,
// protection); append-only (the source bytes unchanged outside the inserted
// runs and count values, an absent collection inserted in CT_Stylesheet
// order, nothing rewritten when nothing was appended); and the from-scratch
// base built from model.styles[0].
//
// SPEC-123 "Named cell styles" › xlsx: reading cellStyles (file order, the
// first per xfId, builtinId, hidden, sourceXfId) with their cellStyleXfs
// definitions (applyX "0"/"false" excluding, absent including; an orphan
// xfId an empty definition) and each cellXfs style's namedStyle (Normal —
// builtinId 0, else the one named "Normal" — absent); writing a style's
// xfId (a source sourceXfId reused, a definition the part lacks appended
// once as a cellStyleXfs xf with applyX="0" for its excluded aspects plus a
// cellStyle carrying a built-in's builtinId, Normal's xf and cellStyle
// first in a part without them, from scratch too), write → read, append-only
// over every collection, and a whole save of spec117-styles.xlsx.
//
// SPEC-129 "xlsx wire mapping" › dxfs: `<dxfs>` read by index with Excel's
// quirks (a patternFill with no patternType is solid in its bgColor, `<b
// val="0"/>` an explicit not-bold, `<u val="none"/>`, tinted theme colours),
// each dxf's XML kept verbatim; the writer's append-only `<dxfs>` (nothing
// appended until a rule writes a dxfId past the source's, the count
// updated, a solid fill as `bgColor`, a numFmt's id the one numFmts gives
// its code else the next free one ≥ 164, `<dxfs>` created before
// tableStyles), read back to the same dxfs.
import { describe, expect, it } from "vitest";
import { DEFAULT_THEME, hexToRgb, resolveColor, resolveFont, type DocumentTheme, type ModelColor } from "../../core/theme";
import { unzip } from "../../core/zip";
import { getCell } from "../address";
import { buildSpec117StylesXlsxBytes, SPEC117_BORDER_STYLES, SPEC117_STYLES_XLSX, SPEC117_XF } from "../fixtures/build";
import { packageIntegrity } from "../fixtures/packageIntegrity";
import type { Dxf } from "../conditionalFormats";
import type { CellFont, CellStyle, SheetColorRef } from "../model";
import { readSpreadsheetBytes } from "../open";
import { createBlankWorkbook } from "../sheets";
import { builtinCellStyle, cellStyle, INDEXED_COLORS, internStyle, STYLE_ASPECTS, styleKey, type NamedCellStyle, type StyleAspect } from "../styles";
import { parseXlsx } from "./reader";
import { readXlsxDxfs, readXlsxStyles, XlsxFontReader, XlsxStyleWriter } from "./styles";
import {
  modelColorToXlsxAttrs,
  tintTransforms,
  XLSX_THEME_SLOTS,
  xlsxColorToModel,
  xlsxThemeIndex,
  xlsxThemeSlot,
  type XlsxColorAttributes,
} from "./themeColor";
import { FROM_SCRATCH_STYLES_XML, writeXlsx } from "./writer";

const MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const STRICT = "http://purl.oclc.org/ooxml/spreadsheetml/main";

/** Asserts `actual` (#rrggbb) is within ±1 of `expected` on every channel. */
function expectNearColor(actual: string, expected: string): void {
  const a = hexToRgb(actual).map((c) => Math.round(c * 255));
  const e = hexToRgb(expected).map((c) => Math.round(c * 255));
  a.forEach((channel, i) => expect(Math.abs(channel - e[i]!), `${actual} vs ${expected}`).toBeLessThanOrEqual(1));
}

/** A styles part with the given records; the defaults are a Calibri 11
 * font, the two fills Excel always writes, an empty border and one xf. */
function stylesXml(options: {
  fonts?: string[];
  fills?: string[];
  borders?: string[];
  xfs?: string[];
  extra?: string;
  ns?: string;
}): string {
  const fonts = options.fonts ?? [`<font><sz val="11"/><name val="Calibri"/></font>`];
  const fills = options.fills ?? [`<fill><patternFill patternType="none"/></fill>`, `<fill><patternFill patternType="gray125"/></fill>`];
  const borders = options.borders ?? [`<border><left/><right/><top/><bottom/><diagonal/></border>`];
  const xfs = options.xfs ?? [`<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>`];
  return (
    `<styleSheet xmlns="${options.ns ?? MAIN}">` +
    `<fonts count="${fonts.length}">${fonts.join("")}</fonts><fills count="${fills.length}">${fills.join("")}</fills>` +
    `<borders count="${borders.length}">${borders.join("")}</borders>` +
    `<cellXfs count="${xfs.length}">${xfs.join("")}</cellXfs>${options.extra ?? ""}</styleSheet>`
  );
}

/** The CellFont one `<font>` element reads as. */
function fontOf(fontXml: string, theme?: DocumentTheme): CellStyle["font"] {
  return readXlsxStyles(stylesXml({ fonts: [fontXml] }), theme).styles[0]!.font;
}

/** The fill of an xf pointing at fills[2] = `fillXml`. */
function fillOf(fillXml: string): CellStyle["fill"] {
  const xml = stylesXml({
    fills: [`<fill><patternFill patternType="none"/></fill>`, `<fill><patternFill patternType="gray125"/></fill>`, fillXml],
    xfs: [`<xf fontId="0" fillId="0" borderId="0"/>`, `<xf fontId="0" fillId="2" borderId="0" applyFill="1"/>`],
  });
  return readXlsxStyles(xml, undefined).styles[1]!.fill;
}

/** The borders of an xf pointing at borders[1] = `borderXml`. */
function bordersOf(borderXml: string): CellStyle["borders"] {
  const xml = stylesXml({
    borders: [`<border><left/><right/><top/><bottom/><diagonal/></border>`, borderXml],
    xfs: [`<xf fontId="0" fillId="0" borderId="0"/>`, `<xf fontId="0" fillId="0" borderId="1" applyBorder="1"/>`],
  });
  return readXlsxStyles(xml, undefined).styles[1]!.borders;
}

/** The style of one xf (the second, after a plain Normal). */
function xfStyle(xfXml: string): CellStyle {
  const xml = stylesXml({ xfs: [`<xf fontId="0" fillId="0" borderId="0"/>`, xfXml] });
  return readXlsxStyles(xml, undefined).styles[1]!;
}

const color = (attrs: XlsxColorAttributes, palette?: string[]): ModelColor | undefined => xlsxColorToModel(attrs, palette);

describe("colours (xlsx/themeColor.ts)", () => {
  it("maps theme indices through Excel's swap: 0 lt1, 1 dk1, 2 lt2, 3 dk2, 4–9 accents, 10 hlink, 11 folHlink", () => {
    const expected = ["lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"];
    expect(XLSX_THEME_SLOTS).toEqual(expected);
    expected.forEach((slot, n) => {
      expect(color({ theme: String(n) })).toEqual({ theme: slot });
      expect(xlsxThemeSlot(n)).toBe(slot);
      expect(xlsxThemeIndex(slot as SheetColorRef["theme"])).toBe(n);
    });
  });

  it("ignores any other theme index (absent)", () => {
    for (const theme of ["12", "-1", "1.5", "x", ""]) {
      expect(color({ theme }), theme).toBeUndefined();
    }
  });

  it("turns a negative tint into lumMod 1 + t and a positive one into lumMod 1 − t, lumOff t, keeping the source text", () => {
    expect(tintTransforms(-0.25)).toEqual({ lumMod: 0.75 });
    expect(tintTransforms(0.4)).toEqual({ lumMod: 0.6, lumOff: 0.4 });
    expect(tintTransforms(0)).toEqual({});
    expect(color({ theme: "4", tint: "-0.25" })).toEqual({ theme: "accent1", lumMod: 0.75, xlsxTint: "-0.25" });
    const lighter = color({ theme: "4", tint: "0.39997558519241921" }) as SheetColorRef;
    expect(lighter.theme).toBe("accent1");
    expect(lighter.lumMod).toBeCloseTo(1 - 0.39997558519241921, 15);
    expect(lighter.lumOff).toBe(0.39997558519241921);
    expect(lighter.xlsxTint).toBe("0.39997558519241921");
    // A zero tint is no transform, but its text is kept for the writer.
    expect(color({ theme: "1", tint: "0" })).toEqual({ theme: "dk1", xlsxTint: "0" });
    // A malformed tint is ignored.
    expect(color({ theme: "1", tint: "abc" })).toEqual({ theme: "dk1" });
  });

  it('pins <color theme="4" tint="0.39997558519241921"/> ("Accent 1, Lighter 40%") to Excel\'s #8ea9db ±1 per channel', () => {
    const ref = color({ theme: "4", tint: "0.39997558519241921" })!;
    expectNearColor(resolveColor(ref, DEFAULT_THEME), "#8ea9db");
    // …and "Accent 1, Darker 25%" to Excel's #2f5597.
    expectNearColor(resolveColor(color({ theme: "4", tint: "-0.249977111117893" })!, DEFAULT_THEME), "#2f5597");
  });

  it("reads indexed colours through the legacy palette, or the file's override; 64 is automatic, 65 white", () => {
    for (let n = 0; n < 64; n += 1) {
      expect(color({ indexed: String(n) })).toBe(INDEXED_COLORS[n]);
    }
    const palette = [...INDEXED_COLORS];
    palette[44] = "#ffb6c1";
    expect(color({ indexed: "44" }, palette)).toBe("#ffb6c1");
    expect(color({ indexed: "43" }, palette)).toBe(INDEXED_COLORS[43]);
    // A short override falls back to the legacy entry.
    expect(color({ indexed: "10" }, ["#123456"])).toBe("#ff0000");
    expect(color({ indexed: "64" })).toBeUndefined();
    expect(color({ indexed: "65" })).toBe("#ffffff");
    expect(color({ indexed: "66" })).toBeUndefined();
    expect(color({ indexed: "x" })).toBeUndefined();
  });

  it("reads rgb ARGB as lowercase #rrggbb with the alpha ignored, and accepts a 6-digit value", () => {
    expect(color({ rgb: "FFFF0000" })).toBe("#ff0000");
    expect(color({ rgb: "00C6EFCE" })).toBe("#c6efce");
    expect(color({ rgb: "80AbCdEf" })).toBe("#abcdef");
    expect(color({ rgb: "00FF00" })).toBe("#00ff00");
    expect(color({ rgb: "red" })).toBeUndefined();
  });

  it("reads auto as automatic (absent) and maps by the first present attribute: auto → theme → indexed → rgb", () => {
    expect(color({ auto: "1", rgb: "FFFF0000" })).toBeUndefined();
    expect(color({ auto: "true", theme: "4" })).toBeUndefined();
    expect(color({ auto: "0", rgb: "FFFF0000" })).toBe("#ff0000");
    expect(color({ theme: "4", indexed: "2", rgb: "FF00FF00" })).toEqual({ theme: "accent1" });
    expect(color({ indexed: "2", rgb: "FF00FF00" })).toBe("#ff0000");
    expect(color({})).toBeUndefined();
  });

  describe("writing back (modelColorToXlsxAttrs)", () => {
    it("re-emits a read reference's source tint text verbatim", () => {
      for (const tint of ["0.39997558519241921", "-0.249977111117893", "0.59999389629810485", "-0.499984740745262", "0"]) {
        const ref = color({ theme: "4", tint })!;
        expect(modelColorToXlsxAttrs(ref, DEFAULT_THEME)).toBe(` theme="4" tint="${tint}"`);
      }
      expect(modelColorToXlsxAttrs(color({ theme: "1" })!, DEFAULT_THEME)).toBe(` theme="1"`);
    });

    it("writes lumOff as the tint, else lumMod − 1, for product-made references", () => {
      expect(modelColorToXlsxAttrs({ theme: "accent2", lumMod: 0.75 }, DEFAULT_THEME)).toBe(` theme="5" tint="-0.25"`);
      expect(modelColorToXlsxAttrs({ theme: "accent1", lumMod: 0.6, lumOff: 0.4 }, DEFAULT_THEME)).toBe(` theme="4" tint="0.4"`);
      expect(modelColorToXlsxAttrs({ theme: "lt1", lumMod: 0.5 }, DEFAULT_THEME)).toBe(` theme="0" tint="-0.5"`);
      expect(modelColorToXlsxAttrs({ theme: "dk2" }, undefined)).toBe(` theme="3"`);
    });

    it("drops a source tint text the transforms no longer stand for", () => {
      const edited: SheetColorRef = { theme: "accent1", lumMod: 0.5, xlsxTint: "-0.25" };
      expect(modelColorToXlsxAttrs(edited, DEFAULT_THEME)).toBe(` theme="4" tint="-0.5"`);
    });

    it("writes a reference no tint expresses exactly as the rgb it resolves to", () => {
      const inexpressible: SheetColorRef[] = [
        { theme: "accent1", lumMod: 0.5, lumOff: 0.1 }, // does not sum to 1
        { theme: "accent1", lumOff: 0.2 }, // lumOff alone
        { theme: "accent1", lumMod: 1.2 }, // a lone lumMod above 1 is no tint
        { theme: "accent1", shade: 0.5 },
        { theme: "accent1", tint: 0.5 },
      ];
      for (const ref of inexpressible) {
        const resolved = resolveColor(ref, DEFAULT_THEME).slice(1).toUpperCase();
        expect(modelColorToXlsxAttrs(ref, DEFAULT_THEME), JSON.stringify(ref)).toBe(` rgb="FF${resolved}"`);
      }
      // Resolved against the workbook's own theme when it has one.
      const theme: DocumentTheme = { ...DEFAULT_THEME, colors: { ...DEFAULT_THEME.colors, accent1: "#000080" } };
      expect(modelColorToXlsxAttrs({ theme: "accent1", shade: 1 }, theme)).toBe(` rgb="FF000080"`);
    });

    it("writes a hex colour as FFRRGGBB", () => {
      expect(modelColorToXlsxAttrs("#ff0000", DEFAULT_THEME)).toBe(` rgb="FFFF0000"`);
      expect(modelColorToXlsxAttrs("#8ea9db", undefined)).toBe(` rgb="FF8EA9DB"`);
    });

    it("round-trips the tint arithmetic both ways", () => {
      for (const t of [-1, -0.75, -0.249977111117893, -0.1, 0.1, 0.39997558519241921, 0.8, 1]) {
        const read = color({ theme: "6", tint: String(t) }) as SheetColorRef;
        const { xlsxTint: _source, ...product } = read;
        const written = /tint="([^"]+)"/.exec(modelColorToXlsxAttrs(product, DEFAULT_THEME))![1]!;
        expect(Number(written), String(t)).toBeCloseTo(t, 12);
        const again = color({ theme: "6", tint: written }) as SheetColorRef;
        expect(again.lumMod ?? 1, String(t)).toBeCloseTo(read.lumMod ?? 1, 12);
        expect(again.lumOff ?? 0, String(t)).toBeCloseTo(read.lumOff ?? 0, 12);
        expect(resolveColor(again, DEFAULT_THEME)).toBe(resolveColor(read, DEFAULT_THEME));
      }
    });
  });
});

describe("fonts", () => {
  it("reads b, i and strike with val absent or true", () => {
    expect(fontOf(`<font><b/><i/><strike/></font>`)).toEqual({ bold: true, italic: true, strike: true });
    expect(fontOf(`<font><b val="1"/><i val="true"/><strike val="1"/></font>`)).toEqual({ bold: true, italic: true, strike: true });
    expect(fontOf(`<font><b val="0"/><i val="false"/><strike val="0"/></font>`)).toEqual({});
  });

  it("reads u@val: absent → single, none → absent, the other three as written", () => {
    expect(fontOf(`<font><u/></font>`)).toEqual({ underline: "single" });
    for (const val of ["single", "double", "singleAccounting", "doubleAccounting"]) {
      expect(fontOf(`<font><u val="${val}"/></font>`)).toEqual({ underline: val });
    }
    expect(fontOf(`<font><u val="none"/></font>`)).toEqual({});
  });

  it("reads vertAlign superscript/subscript, baseline as absent", () => {
    expect(fontOf(`<font><vertAlign val="superscript"/></font>`)).toEqual({ vertAlign: "superscript" });
    expect(fontOf(`<font><vertAlign val="subscript"/></font>`)).toEqual({ vertAlign: "subscript" });
    expect(fontOf(`<font><vertAlign val="baseline"/></font>`)).toEqual({});
  });

  it("reads sz, color and name → family", () => {
    expect(fontOf(`<font><sz val="10.5"/><color rgb="FF00B050"/><name val="Arial"/><family val="2"/></font>`)).toEqual({
      family: "Arial",
      sizePt: 10.5,
      color: "#00b050",
    });
    expect(fontOf(`<font><color theme="1"/></font>`)).toEqual({ color: { theme: "dk1" } });
    expect(fontOf(`<font><color auto="1"/></font>`)).toEqual({}); // automatic = absent
  });

  it("lets scheme minor/major win over the name, in either order", () => {
    expect(fontOf(`<font><name val="Calibri"/><scheme val="minor"/></font>`)).toEqual({ family: { theme: "minor" } });
    expect(fontOf(`<font><scheme val="major"/><name val="Calibri Light"/></font>`)).toEqual({ family: { theme: "major" } });
    expect(fontOf(`<font><name val="Calibri"/><scheme val="none"/></font>`)).toEqual({ family: "Calibri" });
  });

  it("does not model condense, extend, outline, shadow, family or charset", () => {
    expect(fontOf(`<font><condense/><extend/><outline/><shadow/><sz val="9"/><family val="3"/><charset val="128"/></font>`)).toEqual({ sizePt: 9 });
  });

  it("maps a run's rPr (rFont) through the same reader", () => {
    const reader = new XlsxFontReader();
    reader.add("rFont", (name) => (name === "val" ? "Cambria" : undefined));
    reader.add("i", () => undefined);
    expect(reader.build()).toEqual({ family: "Cambria", italic: true });
  });
});

describe("fills", () => {
  it("reads patternType none or absent as no fill", () => {
    expect(fillOf(`<fill><patternFill patternType="none"/></fill>`)).toBeUndefined();
    expect(fillOf(`<fill><patternFill><fgColor rgb="FFFF0000"/></patternFill></fill>`)).toBeUndefined();
    expect(fillOf(`<fill/>`)).toBeUndefined();
  });

  it("reads solid as a solid fill of fgColor, black when it has none or it is automatic", () => {
    expect(fillOf(`<fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/><bgColor indexed="64"/></patternFill></fill>`)).toEqual({
      kind: "solid",
      color: "#ffff00",
    });
    expect(fillOf(`<fill><patternFill patternType="solid"><fgColor theme="4" tint="-0.25"/></patternFill></fill>`)).toEqual({
      kind: "solid",
      color: { theme: "accent1", lumMod: 0.75, xlsxTint: "-0.25" },
    });
    expect(fillOf(`<fill><patternFill patternType="solid"/></fill>`)).toEqual({ kind: "solid", color: "#000000" });
    expect(fillOf(`<fill><patternFill patternType="solid"><fgColor indexed="64"/></patternFill></fill>`)).toEqual({ kind: "solid", color: "#000000" });
  });

  it("reads every other pattern with its fg and bg", () => {
    expect(fillOf(`<fill><patternFill patternType="gray125"/></fill>`)).toEqual({ kind: "pattern", pattern: "gray125" });
    expect(fillOf(`<fill><patternFill patternType="darkTrellis"><fgColor theme="5"/><bgColor rgb="FFFFFFCC"/></patternFill></fill>`)).toEqual({
      kind: "pattern",
      pattern: "darkTrellis",
      fg: { theme: "accent2" },
      bg: "#ffffcc",
    });
    expect(fillOf(`<fill><patternFill patternType="lightGrid"><bgColor indexed="65"/></patternFill></fill>`)).toEqual({
      kind: "pattern",
      pattern: "lightGrid",
      bg: "#ffffff",
    });
  });

  it("reads a linear gradient (the default type) with its degree and stops", () => {
    expect(fillOf(`<fill><gradientFill degree="90"><stop position="0"><color theme="0"/></stop><stop position="1"><color theme="4"/></stop></gradientFill></fill>`)).toEqual({
      kind: "gradient",
      type: "linear",
      degree: 90,
      stops: [
        { position: 0, color: { theme: "lt1" } },
        { position: 1, color: { theme: "accent1" } },
      ],
    });
  });

  it("reads a path gradient's focus", () => {
    expect(
      fillOf(
        `<fill><gradientFill type="path" left="0.5" right="0.5" top="0.25" bottom="0.75"><stop position="0"><color rgb="FFFFFFFF"/></stop>` +
          `<stop position="0.5"><color auto="1"/></stop></gradientFill></fill>`,
      ),
    ).toEqual({
      kind: "gradient",
      type: "path",
      left: 0.5,
      right: 0.5,
      top: 0.25,
      bottom: 0.75,
      stops: [
        { position: 0, color: "#ffffff" },
        { position: 0.5, color: "#000000" },
      ],
    });
  });
});

describe("borders", () => {
  it("reads all thirteen line styles", () => {
    const styles = [
      "thin",
      "medium",
      "thick",
      "dashed",
      "dotted",
      "double",
      "hair",
      "mediumDashed",
      "dashDot",
      "mediumDashDot",
      "dashDotDot",
      "mediumDashDotDot",
      "slantDashDot",
    ];
    for (const style of styles) {
      expect(bordersOf(`<border><left style="${style}"/><right/><top/><bottom/><diagonal/></border>`), style).toEqual({ left: { style } });
    }
  });

  it("reads each edge's colour (automatic absent) and no edge for a style absent or none", () => {
    expect(
      bordersOf(
        `<border><left style="thin"><color indexed="64"/></left><right style="medium"><color theme="4"/></right>` +
          `<top style="none"><color rgb="FFFF0000"/></top><bottom style="thick"><color rgb="FFFF0000"/></bottom><diagonal/></border>`,
      ),
    ).toEqual({ left: { style: "thin" }, right: { style: "medium", color: { theme: "accent1" } }, bottom: { style: "thick", color: "#ff0000" } });
    expect(bordersOf(`<border><left/><right/><top/><bottom/><diagonal/></border>`)).toBeUndefined();
  });

  it("reads diagonalUp/diagonalDown onto the diagonal edge", () => {
    expect(bordersOf(`<border diagonalUp="1" diagonalDown="1"><left/><right/><top/><bottom/><diagonal style="thin"><color auto="1"/></diagonal></border>`)).toEqual({
      diagonal: { style: "thin", up: true, down: true },
    });
    expect(bordersOf(`<border diagonalDown="1"><diagonal style="dashed"/></border>`)).toEqual({ diagonal: { style: "dashed", down: true } });
    // The flags without a diagonal line draw nothing.
    expect(bordersOf(`<border diagonalUp="1"><diagonal/></border>`)).toBeUndefined();
  });

  it("reads start/end as left/right", () => {
    expect(bordersOf(`<border><start style="thin"/><end style="double"/><top/><bottom/><diagonal/></border>`)).toEqual({
      left: { style: "thin" },
      right: { style: "double" },
    });
  });
});

describe("cellXfs", () => {
  const xml = stylesXml({
    fonts: [`<font><sz val="11"/><name val="Calibri"/><scheme val="minor"/></font>`, `<font><b/><sz val="11"/><color rgb="FFFF0000"/></font>`],
    fills: [`<fill><patternFill patternType="none"/></fill>`, `<fill><patternFill patternType="gray125"/></fill>`, `<fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/></patternFill></fill>`],
    borders: [`<border/>`, `<border><bottom style="thin"/></border>`],
    xfs: [
      `<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>`,
      `<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>`,
      `<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>`,
      `<xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyFont="0" applyNumberFormat="1"/>`,
    ],
  });

  it("maps each xf to one style, in order, never deduplicated (styles[i].sourceXf === i)", () => {
    const { styles } = readXlsxStyles(xml, undefined);
    expect(styles).toHaveLength(4);
    styles.forEach((style, i) => expect(style.sourceXf).toBe(i));
    expect(styles[1]).toEqual({
      font: { bold: true, sizePt: 11, color: "#ff0000" },
      fill: { kind: "solid", color: "#ffff00" },
      borders: { bottom: { style: "thin" } },
      sourceXf: 1,
    });
    expect(styles[2]).toEqual({ ...styles[1], sourceXf: 2 });
  });

  it("gives every style its own copies of the font, fill and border — plain data, no shared objects", () => {
    const { styles } = readXlsxStyles(xml, undefined);
    expect(styles[1]!.font).not.toBe(styles[2]!.font);
    expect(styles[1]!.fill).not.toBe(styles[2]!.fill);
    expect(styles[1]!.borders).not.toBe(styles[2]!.borders);
    expect(styles[1]!.borders!.bottom).not.toBe(styles[2]!.borders!.bottom);
    expect(JSON.parse(JSON.stringify(styles))).toEqual(styles);
  });

  it("copies a font equal to the default style's, ignores the apply* flags and xfId; numFmtId 14 is the locale-dependent built-in (SPEC-118)", () => {
    const { styles } = readXlsxStyles(xml, undefined);
    expect(styles[0]).toEqual({ font: { family: { theme: "minor" }, sizePt: 11 }, sourceXf: 0 });
    expect(styles[3]).toEqual({ font: { family: { theme: "minor" }, sizePt: 11 }, numberFormat: "m/d/yyyy", builtinFormatId: 14, sourceXf: 3 });
    expect(styles[3]!.font).not.toBe(styles[0]!.font);
  });

  it("reads the alignment child", () => {
    expect(
      xfStyle(`<xf fontId="0" fillId="0" borderId="0"><alignment horizontal="centerContinuous" vertical="top" wrapText="1" indent="3" textRotation="45" shrinkToFit="1"/></xf>`)
        .alignment,
    ).toEqual({ horizontal: "centerContinuous", vertical: "top", wrap: true, indent: 3, textRotation: 45, shrinkToFit: true });
    expect(xfStyle(`<xf fontId="0" fillId="0" borderId="0"><alignment textRotation="255"/></xf>`).alignment).toEqual({ textRotation: 255 });
    expect(xfStyle(`<xf fontId="0" fillId="0" borderId="0"><alignment textRotation="135" vertical="justify"/></xf>`).alignment).toEqual({
      vertical: "justify",
      textRotation: 135,
    });
    // Defaults and nonsense store nothing.
    expect(xfStyle(`<xf fontId="0" fillId="0" borderId="0"><alignment indent="0" textRotation="0" wrapText="0" horizontal="sideways"/></xf>`).alignment).toBeUndefined();
  });

  it("reads the protection child: locked=\"0\" → locked: false, hidden=\"1\" → hidden", () => {
    expect(xfStyle(`<xf fontId="0" fillId="0" borderId="0" applyProtection="1"><protection locked="0" hidden="1"/></xf>`).protection).toEqual({
      locked: false,
      hidden: true,
    });
    expect(xfStyle(`<xf fontId="0" fillId="0" borderId="0"><protection locked="1"/></xf>`).protection).toBeUndefined();
  });

  it("reads quotePrefix", () => {
    expect(xfStyle(`<xf fontId="0" fillId="0" borderId="0" xfId="0" quotePrefix="1"/>`).quotePrefix).toBe(true);
    expect(xfStyle(`<xf fontId="0" fillId="0" borderId="0" xfId="0" quotePrefix="0"/>`).quotePrefix).toBeUndefined();
  });

  it("omits what an xf points at but the part lacks", () => {
    expect(xfStyle(`<xf fontId="9" fillId="9" borderId="9"/>`)).toEqual({ sourceXf: 1 });
  });
});

describe("number formats (SPEC-118 \"Reading xlsx number formats\")", () => {
  /** The number-format fields of each xf, over a part declaring `numFmts`
   * (`<numFmt>` elements) and one xf per `numFmtId` attribute text. */
  function formatsOf(numFmtIds: Array<string | null>, numFmts = ""): Array<Pick<CellStyle, "numberFormat" | "builtinFormatId">> {
    const xfs = numFmtIds.map((id) => `<xf${id === null ? "" : ` numFmtId="${id}"`} fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="0"/>`);
    const xml = stylesXml({ xfs }).replace("<fonts", `${numFmts ? `<numFmts count="1">${numFmts}</numFmts>` : ""}<fonts`);
    return readXlsxStyles(xml, undefined).styles.map(({ numberFormat, builtinFormatId }) => ({
      ...(numberFormat !== undefined ? { numberFormat } : {}),
      ...(builtinFormatId !== undefined ? { builtinFormatId } : {}),
    }));
  }

  it("a locale-dependent built-in sets builtinFormatId with its en-US code; another built-in sets only its code", () => {
    expect(formatsOf(["14", "22", "5", "44", "10", "49", "46", "2"])).toEqual([
      { numberFormat: "m/d/yyyy", builtinFormatId: 14 },
      { numberFormat: "m/d/yyyy h:mm", builtinFormatId: 22 },
      { numberFormat: '"$"#,##0_);\\("$"#,##0\\)', builtinFormatId: 5 },
      { numberFormat: '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)', builtinFormatId: 44 },
      { numberFormat: "0.00%" },
      { numberFormat: "@" },
      { numberFormat: "[h]:mm:ss" },
      { numberFormat: "0.00" },
    ]);
  });

  it("General, an absent numFmtId, an unknown id and the CJK ids leave both fields absent", () => {
    expect(formatsOf(["0", null, "163", "200", "23", "36", "50", "81"])).toEqual([{}, {}, {}, {}, {}, {}, {}, {}]);
  });

  it("an explicit numFmt wins for any id — Excel writes 43 and 44 itself — and never sets builtinFormatId", () => {
    const numFmts =
      `<numFmt numFmtId="44" formatCode="_-* #,##0.00\\ &quot;€&quot;_-;\\-* #,##0.00\\ &quot;€&quot;_-;_-* &quot;-&quot;??\\ &quot;€&quot;_-;_-@_-"/>` +
      `<numFmt numFmtId="14" formatCode="dd/mm/yy"/><numFmt numFmtId="164" formatCode="[$-F800]dddd\\,\\ mmmm\\ dd\\,\\ yyyy"/>`;
    expect(formatsOf(["44", "14", "164"], numFmts)).toEqual([
      { numberFormat: '_-* #,##0.00\\ "€"_-;\\-* #,##0.00\\ "€"_-;_-* "-"??\\ "€"_-;_-@_-' },
      { numberFormat: "dd/mm/yy" },
      { numberFormat: "[$-F800]dddd\\,\\ mmmm\\ dd\\,\\ yyyy" },
    ]);
  });

  it("a code equal to General (any case) is stored absent; applyNumberFormat is ignored", () => {
    expect(formatsOf(["164", "165"], `<numFmt numFmtId="164" formatCode="GENERAL"/><numFmt numFmtId="165" formatCode="0.0"/>`)).toEqual([
      {},
      { numberFormat: "0.0" },
    ]);
  });
});

describe("the parts left unread and the override palette", () => {
  // SPEC-123 reads these two (see "SPEC-123 named cell styles" below).
  const named =
    `<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>` +
    `<cellStyles count="2"><cellStyle name="Good" xfId="0" builtinId="26"/><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>`;
  const unread =
    `<dxfs count="1"><dxf><font><b/></font><fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf></dxfs>` +
    `<tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>` +
    `<colors><mruColors><color rgb="FF66FF33"/></mruColors></colors>` +
    `<extLst><ext uri="{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}"><slicerStyles defaultSlicerStyle="SlicerStyleLight1"/></ext></extLst>`;

  it("leaves tableStyles, extLst and mruColors unread (SPEC-129 reads the dxfs); numFmts are SPEC-118's and read, cellStyleXfs/cellStyles SPEC-123's", () => {
    const xfs = [`<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>`];
    const numFmts = `<numFmts count="1"><numFmt numFmtId="164" formatCode="0.0%"/></numFmts><fonts`;
    const bare = readXlsxStyles(stylesXml({ xfs }).replace("<fonts", numFmts), undefined);
    const full = readXlsxStyles(stylesXml({ xfs, extra: unread }).replace("<fonts", numFmts), undefined);
    expect({ ...full, dxfs: [] }).toEqual(bare);
    expect(full.dxfs.map(({ sourceXml: _xml, ...dxf }) => dxf)).toEqual([{ font: { bold: true }, fill: { kind: "solid", color: "#ffc7ce" } }]);
    expect(full.styles[0]!.numberFormat).toBe("0.0%");
    expect(full.styles[0]!.namedStyle).toBeUndefined();
    expect(full.indexedColors).toBeUndefined();
    // SPEC-123: the named styles are read. Good, the first cellStyle on
    // xfId 0, wins it (Normal's, on the same xfId, is skipped), so the xf
    // pointing there is in Good — an absent applyX includes every aspect.
    const withNamed = readXlsxStyles(stylesXml({ xfs, extra: named + unread }).replace("<fonts", numFmts), undefined);
    expect(withNamed).toEqual({
      ...bare,
      dxfs: full.dxfs,
      styles: [{ ...bare.styles[0], namedStyle: "Good" }],
      namedStyles: [{ name: "Good", builtinId: 26, style: { font: { family: "Calibri", sizePt: 11 } }, includes: [...STYLE_ASPECTS], sourceXfId: 0 }],
    });
  });

  it("reads <colors><indexedColors> as 64 lowercase entries, the fonts and fills resolving through them", () => {
    const palette = INDEXED_COLORS.map((hex, i) => (i === 44 ? "FFB6C1" : hex.slice(1).toUpperCase()));
    palette[3] = "not a colour";
    const xml = stylesXml({
      fonts: [`<font><color indexed="44"/></font>`],
      fills: [`<fill><patternFill patternType="none"/></fill>`, `<fill><patternFill patternType="solid"><fgColor indexed="44"/></patternFill></fill>`],
      xfs: [`<xf fontId="0" fillId="1" borderId="0"/>`],
      extra: `<colors><indexedColors>${palette.map((rgb) => `<rgbColor rgb="00${rgb}"/>`).join("")}</indexedColors></colors>`,
    });
    const table = readXlsxStyles(xml, undefined);
    expect(table.indexedColors).toHaveLength(64);
    expect(table.indexedColors![44]).toBe("#ffb6c1");
    expect(table.indexedColors![3]).toBe(INDEXED_COLORS[3]); // unreadable → the legacy entry
    expect(table.indexedColors!.filter((hex, i) => hex !== INDEXED_COLORS[i])).toEqual(["#ffb6c1"]);
    expect(table.styles[0]!.font!.color).toBe("#ffb6c1");
    expect(table.styles[0]!.fill).toEqual({ kind: "solid", color: "#ffb6c1" });
  });
});

describe("the default font and missing parts", () => {
  it("is fonts[cellXfs[0].fontId], its scheme family resolved against the theme", () => {
    const xml = stylesXml({
      fonts: [`<font><sz val="9"/><name val="Arial"/></font>`, `<font><sz val="12"/><name val="Calibri"/><scheme val="minor"/></font>`],
      xfs: [`<xf fontId="1" fillId="0" borderId="0"/>`],
    });
    expect(readXlsxStyles(xml, undefined).defaultFont).toEqual({ family: "Calibri", sizePt: 12 });
    const aptos: DocumentTheme = { ...DEFAULT_THEME, fonts: { major: "Aptos Display", minor: "Aptos Narrow" } };
    expect(readXlsxStyles(xml, aptos).defaultFont).toEqual({ family: "Aptos Narrow", sizePt: 12 });
    const named = stylesXml({ fonts: [`<font><sz val="10"/><name val="Arial"/></font>`] });
    expect(readXlsxStyles(named, aptos).defaultFont).toEqual({ family: "Arial", sizePt: 10 });
    const major = stylesXml({ fonts: [`<font><scheme val="major"/></font>`] });
    expect(readXlsxStyles(major, undefined).defaultFont).toEqual({ family: "Calibri Light", sizePt: 11 });
  });

  it("gives a workbook with no styles part one empty style and Calibri 11", () => {
    // SPEC-129: and no differential styles.
    expect(readXlsxStyles(undefined, undefined)).toEqual({ styles: [{}], defaultFont: { family: "Calibri", sizePt: 11 }, dxfs: [] });
    // A part that is not well-formed, or names no xf, reads the same.
    expect(readXlsxStyles("<styleSheet", undefined)).toEqual({ styles: [{}], defaultFont: { family: "Calibri", sizePt: 11 }, dxfs: [] });
    expect(readXlsxStyles(`<styleSheet xmlns="${MAIN}"/>`, undefined)).toEqual({ styles: [{}], defaultFont: { family: "Calibri", sizePt: 11 }, dxfs: [] });
  });

  it("reads a strict (ISO 29500) styles part", () => {
    const xml = stylesXml({ ns: STRICT, fonts: [`<font><b/><sz val="10"/><name val="Arial"/></font>`] });
    expect(readXlsxStyles(xml, undefined)).toEqual({
      styles: [{ font: { family: "Arial", sizePt: 10, bold: true }, sourceXf: 0 }],
      defaultFont: { family: "Arial", sizePt: 10 },
      dxfs: [],
    });
  });
});

async function parsed() {
  return parseXlsx(await buildSpec117StylesXlsxBytes());
}

describe("spec117-styles.xlsx (Excel-shaped)", () => {

  it("reads the classic Office theme from the workbook's theme relationship", async () => {
    const { theme } = await parsed();
    expect(theme).toEqual({ name: SPEC117_STYLES_XLSX.themeName, colors: DEFAULT_THEME.colors, fonts: DEFAULT_THEME.fonts });
  });

  it("maps cellXfs one-to-one, each style knowing its xf", async () => {
    const { styles } = await parsed();
    expect(styles).toHaveLength(34);
    styles.forEach((style, i) => expect(style.sourceXf).toBe(i));
    expect(styles[0]).toEqual({ font: { family: { theme: "minor" }, sizePt: 11, color: { theme: "dk1" } }, sourceXf: 0 });
    // SPEC-123: the two xfs whose xfId names a cell style carry its name;
    // every other names xfId 0, Normal (absent).
    expect(styles.flatMap((style, i) => (style.namedStyle === undefined ? [] : [[i, style.namedStyle]]))).toEqual([
      [SPEC117_XF.title, "Heading 1"],
      [SPEC117_XF.good, "Good"],
    ]);
    // SPEC-118: the one custom numFmt (164, "0.0%") is the only format.
    expect(styles.filter((style) => style.numberFormat !== undefined).map((style) => style.numberFormat)).toEqual(["0.0%"]);
    expect(styles.some((style) => style.builtinFormatId !== undefined)).toBe(false);
  });

  it("makes B2's style bold red on a solid yellow fill", async () => {
    const { styles } = await parsed();
    const header = styles[SPEC117_STYLES_XLSX.header.styleId]!;
    expect(header.font).toEqual({ family: { theme: "minor" }, sizePt: 11, bold: true, color: SPEC117_STYLES_XLSX.header.color });
    expect(header.fill).toEqual({ kind: "solid", color: SPEC117_STYLES_XLSX.header.fill });
  });

  it("reads the tinted, patterned, indexed and gradient fills", async () => {
    const { styles, theme, indexedColors } = await parsed();
    const tint = styles[SPEC117_XF.tinted]!.fill as { kind: "solid"; color: ModelColor };
    expectNearColor(resolveColor(tint.color, theme!), "#8ea9db");
    expect(styles[SPEC117_XF.trellis]!.fill).toEqual({ kind: "pattern", pattern: "darkTrellis", fg: { theme: "accent2" }, bg: "#ffffcc" });
    expect(styles[SPEC117_XF.gradient]!.fill).toEqual({
      kind: "gradient",
      type: "linear",
      degree: 90,
      stops: [
        { position: 0, color: { theme: "lt1" } },
        { position: 1, color: { theme: "accent1" } },
      ],
    });
    expect(styles[SPEC117_XF.indexed]!.font!.color).toBe("#ff0000"); // indexed 10, not overridden
    expect(styles[SPEC117_XF.indexed]!.fill).toEqual({ kind: "solid", color: SPEC117_STYLES_XLSX.indexedOverride.color });
    expect(indexedColors).toHaveLength(64);
    expect(indexedColors![SPEC117_STYLES_XLSX.indexedOverride.index]).toBe(SPEC117_STYLES_XLSX.indexedOverride.color);
    const italic = styles[SPEC117_XF.italicTint]!.font!;
    expect(italic.italic).toBe(true);
    expectNearColor(resolveColor(italic.color!, theme!), "#2f5597");
    expect(styles[SPEC117_XF.doubleUnderline]!.font!.underline).toBe("double");
    expect(styles[SPEC117_XF.strike]!.font).toMatchObject({ strike: true, vertAlign: "superscript" });
    expect(styles[SPEC117_XF.good]).toMatchObject({ font: { color: "#006100" }, fill: { kind: "solid", color: "#c6efce" } });
  });

  it("covers all thirteen border styles, the Heading 1 border and a diagonal", async () => {
    const { styles } = await parsed();
    SPEC117_BORDER_STYLES.forEach((style, k) => {
      const borders = styles[SPEC117_XF.borders + k]!.borders!;
      for (const side of ["left", "right", "top", "bottom"] as const) {
        expect(borders[side]?.style, `${style} ${side}`).toBe(style);
      }
    });
    expect(styles[SPEC117_XF.borders + 1]!.borders!.left!.color).toEqual({ theme: "accent1" }); // medium
    expect(styles[SPEC117_XF.borders + 4]!.borders!.left!.color).toBe("#ff0000"); // thick
    expect(styles[SPEC117_XF.borders]!.borders!.left!.color).toBeUndefined(); // indexed 64: automatic
    expect(styles[SPEC117_XF.borders + 6]!.borders!.left!.color).toBeUndefined(); // auto="1"
    expect(styles[SPEC117_XF.title]).toMatchObject({
      font: { bold: true, sizePt: 15, color: { theme: "dk2" } },
      borders: { bottom: { style: "thick", color: { theme: "accent1" } } },
      alignment: { horizontal: "center" },
    });
    expect(styles[SPEC117_XF.diagonal]!.borders).toEqual({ diagonal: { style: "thin", color: "#ff0000", up: true, down: true } });
  });

  it("reads alignment, protection and quotePrefix", async () => {
    const { styles } = await parsed();
    expect(styles[SPEC117_XF.wrap]!.alignment).toEqual({ vertical: "top", wrap: true });
    expect(styles[SPEC117_XF.rotate90]!.alignment).toEqual({ textRotation: 90 });
    expect(styles[SPEC117_XF.indent]!.alignment).toEqual({ horizontal: "left", indent: 2 });
    expect(styles[SPEC117_XF.shrink]!.alignment).toEqual({ vertical: "center", shrinkToFit: true });
    expect(styles[SPEC117_XF.stacked]!.alignment).toEqual({ horizontal: "center", textRotation: 255 });
    expect(styles[SPEC117_XF.unlocked]!.protection).toEqual({ locked: false, hidden: true });
    expect(styles[SPEC117_XF.quotePrefix]!.quotePrefix).toBe(true);
    expect(styles.filter((style) => style.quotePrefix)).toHaveLength(1);
  });
});

// --- SPEC-122: the style writer -------------------------------------------------------

const X14AC = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac";
const DECLARATION = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`;

/** An Excel-shaped styles part: two custom number formats, a bold font, a
 * solid fill, the Heading 1 cell style, and the elements the writer never
 * touches (cellStyles, dxfs, tableStyles, extLst). */
const SOURCE_STYLES =
  `${DECLARATION}\r\n<styleSheet xmlns="${MAIN}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="${X14AC}">` +
  `<numFmts count="2"><numFmt numFmtId="164" formatCode="0.000"/><numFmt numFmtId="170" formatCode="&quot;kg&quot;0"/></numFmts>` +
  `<fonts count="2" x14ac:knownFonts="1"><font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>` +
  `<font><b/><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font></fonts>` +
  `<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>` +
  `<fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/><bgColor indexed="64"/></patternFill></fill></fills>` +
  `<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>` +
  `<cellStyleXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" applyNumberFormat="0"/></cellStyleXfs>` +
  `<cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>` +
  `<xf numFmtId="164" fontId="1" fillId="2" borderId="0" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1"/>` +
  `<xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs>` +
  `<cellStyles count="2"><cellStyle name="Heading 1" xfId="1" builtinId="16"/><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>` +
  `<dxfs count="0"/><tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>` +
  `<extLst><ext uri="{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:slicerStyles defaultSlicerStyle="SlicerStyleLight1"/></ext></extLst></styleSheet>`;

const BOLD_MINOR = { family: { theme: "minor" as const }, sizePt: 11, bold: true as const, color: { theme: "dk1" as const } };
const CODE_44 = BUILTIN_CODE_44();

/** Built-in 44's en-US code (a locale-dependent accounting format). */
function BUILTIN_CODE_44(): string {
  return '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)';
}

/** A writer over `source` for the source's own styles plus `added`. */
function styleWriter(added: CellStyle[], source = SOURCE_STYLES): { writer: XlsxStyleWriter; styles: CellStyle[] } {
  const styles = [...readXlsxStyles(source, undefined).styles, ...added];
  return { writer: new XlsxStyleWriter(source, { styles, theme: undefined }), styles };
}

describe("XlsxStyleWriter: number formats", () => {
  /** The numFmtId a new style with these fields gets, and the numFmts the part then holds. */
  function numFmt(fields: CellStyle, source = SOURCE_STYLES): { id: number; numFmts: string } {
    const { writer } = styleWriter([fields], source);
    const xf = writer.xfFor(readXlsxStyles(source, undefined).styles.length);
    const xml = writer.xml();
    const xfs = /<cellXfs[^>]*>(.*)<\/cellXfs>/.exec(xml)![1]!.match(/<xf\b[^>]*?numFmtId="(\d+)"/g)!;
    return { id: Number(/numFmtId="(\d+)"/.exec(xfs[xf]!)![1]), numFmts: /<numFmts\b.*<\/numFmts>/.exec(xml)?.[0] ?? "" };
  }
  const SOURCE_NUMFMTS = `<numFmts count="2"><numFmt numFmtId="164" formatCode="0.000"/><numFmt numFmtId="170" formatCode="&quot;kg&quot;0"/></numFmts>`;

  it("writes builtinFormatId as the numFmtId, with no numFmt entry — whatever the locale's code", () => {
    expect(numFmt({ numberFormat: "m/d/yyyy", builtinFormatId: 14 })).toEqual({ id: 14, numFmts: SOURCE_NUMFMTS });
    expect(numFmt({ numberFormat: "dd.mm.yyyy hh:mm", builtinFormatId: 22 })).toEqual({ id: 22, numFmts: SOURCE_NUMFMTS });
    expect(numFmt({ numberFormat: CODE_44, builtinFormatId: 44 })).toEqual({ id: 44, numFmts: SOURCE_NUMFMTS });
  });

  it("finds a locale-independent built-in by its code (0–49)", () => {
    const cases: [string, number][] = [["0", 1], ["#,##0", 3], ["0%", 9], ["0.00%", 10], ["0.00E+00", 11], ["# ?/?", 12], ["d-mmm-yy", 15], ["h:mm AM/PM", 18], ["h:mm:ss", 21], ["[h]:mm:ss", 46], ["@", 49]];
    for (const [code, id] of cases) {
      expect(numFmt({ numberFormat: code }), code).toEqual({ id, numFmts: SOURCE_NUMFMTS });
    }
    expect(numFmt({ numberFormat: "General" }).id).toBe(0);
    expect(numFmt({}).id).toBe(0);
  });

  it("never maps a code to a LOCALE_DEPENDENT_BUILTINS id without builtinFormatId: 14's, 22's and 44's codes become custom formats past the highest id", () => {
    expect(numFmt({ numberFormat: "m/d/yyyy" })).toEqual({ id: 171, numFmts: SOURCE_NUMFMTS.replace(`count="2"`, `count="3"`).replace("</numFmts>", `<numFmt numFmtId="171" formatCode="m/d/yyyy"/></numFmts>`) });
    expect(numFmt({ numberFormat: "m/d/yyyy h:mm" }).id).toBe(171);
    expect(numFmt({ numberFormat: CODE_44 })).toEqual({
      id: 171,
      numFmts: SOURCE_NUMFMTS.replace(`count="2"`, `count="3"`).replace("</numFmts>", () => `<numFmt numFmtId="171" formatCode="_(&quot;$&quot;* #,##0.00_);_(&quot;$&quot;* \\(#,##0.00\\);_(&quot;$&quot;* &quot;-&quot;??_);_(@_)"/></numFmts>`),
    });
  });

  it("reuses an existing custom id with the same code", () => {
    expect(numFmt({ numberFormat: "0.000" })).toEqual({ id: 164, numFmts: SOURCE_NUMFMTS });
    expect(numFmt({ numberFormat: `"kg"0` })).toEqual({ id: 170, numFmts: SOURCE_NUMFMTS });
  });

  it("numbers custom formats from 164 in a part with none, inserting <numFmts> first in CT_Stylesheet order", () => {
    const bare = `<styleSheet xmlns="${MAIN}"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><cellXfs count="1"><xf numFmtId="0" fontId="0" xfId="0"/></cellXfs></styleSheet>`;
    expect(numFmt({ numberFormat: "0.0000" }, bare)).toEqual({ id: 164, numFmts: `<numFmts count="1"><numFmt numFmtId="164" formatCode="0.0000"/></numFmts>` });
    expect(numFmt({ numberFormat: "0.0000" }, bare).numFmts).toBe(/^.*<styleSheet[^>]*>(<numFmts.*<\/numFmts>)<fonts/.exec(styleWriterXml(bare, { numberFormat: "0.0000" }))![1]);
  });
});

/** The part a writer over `source` writes for one new style. */
function styleWriterXml(source: string, style: CellStyle): string {
  const { writer, styles } = styleWriter([style], source);
  writer.xfFor(styles.length - 1);
  return writer.xml();
}

describe("XlsxStyleWriter: append-only", () => {
  it("writes a style with sourceXf as that index and rewrites nothing when nothing was appended", () => {
    const { writer, styles } = styleWriter([]);
    expect(styles.map((_, index) => writer.xfFor(index))).toEqual([0, 1, 2]);
    expect(writer.changed).toBe(false);
    expect(writer.xml()).toBe(SOURCE_STYLES);
  });

  it("reuses equal source entries, appends the rest, and leaves the source bytes unchanged outside the inserted runs and count values", () => {
    const italic = { family: "Cambria", sizePt: 14, italic: true as const, strike: true as const, underline: "double" as const, vertAlign: "superscript" as const, color: "#c00000" };
    const added: CellStyle[] = [
      // A: every part found in the source — but a new xf (no sourceXf).
      { font: BOLD_MINOR, fill: { kind: "solid", color: "#ffff00" }, numberFormat: "0.000" },
      // B: a new font, border and custom format; alignment, protection, quotePrefix.
      {
        font: italic,
        borders: { top: { style: "thin" }, bottom: { style: "double", color: "#ff0000" } },
        alignment: { horizontal: "center", vertical: "top", wrap: true, indent: 2, textRotation: 45, shrinkToFit: true },
        protection: { locked: false, hidden: true },
        quotePrefix: true,
        numberFormat: "m/d/yyyy",
      },
      // C: a locale-dependent built-in, no font (font 0).
      { numberFormat: "m/d/yyyy", builtinFormatId: 14 },
      // D: the Heading 1 cell style — apply* against its parent xf.
      { font: BOLD_MINOR, namedStyle: "Heading 1", numberFormat: "0.00%" },
      // E: a theme-colour fill, an existing custom format, hidden only.
      { fill: { kind: "solid", color: { theme: "accent1", lumMod: 0.6, lumOff: 0.4 } }, numberFormat: `"kg"0`, protection: { hidden: true } },
      // F: equal to A — one xf.
      { font: { ...BOLD_MINOR }, fill: { kind: "solid", color: "#ffff00" }, numberFormat: "0.000" },
      // G: built-in 44's code without builtinFormatId — a custom format.
      { numberFormat: CODE_44 },
    ];
    const { writer, styles } = styleWriter(added);
    expect(styles.slice(3).map((_, index) => writer.xfFor(index + 3))).toEqual([3, 4, 5, 6, 7, 3, 8]);
    expect(writer.changed).toBe(true);
    const inserted = {
      numFmts: `<numFmt numFmtId="171" formatCode="m/d/yyyy"/><numFmt numFmtId="172" formatCode="_(&quot;$&quot;* #,##0.00_);_(&quot;$&quot;* \\(#,##0.00\\);_(&quot;$&quot;* &quot;-&quot;??_);_(@_)"/>`,
      fonts: `<font><i/><strike/><u val="double"/><vertAlign val="superscript"/><sz val="14"/><color rgb="FFC00000"/><name val="Cambria"/></font>`,
      fills: `<fill><patternFill patternType="solid"><fgColor theme="4" tint="0.4"/><bgColor indexed="64"/></patternFill></fill>`,
      borders: `<border><left/><right/><top style="thin"><color indexed="64"/></top><bottom style="double"><color rgb="FFFF0000"/></bottom><diagonal/></border>`,
      cellXfs:
        `<xf numFmtId="164" fontId="1" fillId="2" borderId="0" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1"/>` +
        `<xf numFmtId="171" fontId="2" fillId="0" borderId="1" xfId="0" quotePrefix="1" applyNumberFormat="1" applyFont="1" applyBorder="1" applyAlignment="1" applyProtection="1">` +
        `<alignment horizontal="center" vertical="top" textRotation="45" wrapText="1" indent="2" shrinkToFit="1"/><protection locked="0" hidden="1"/></xf>` +
        `<xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>` +
        `<xf numFmtId="10" fontId="1" fillId="0" borderId="0" xfId="1" applyNumberFormat="1"/>` +
        `<xf numFmtId="170" fontId="0" fillId="3" borderId="0" xfId="0" applyNumberFormat="1" applyFill="1" applyProtection="1"><protection hidden="1"/></xf>` +
        `<xf numFmtId="172" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>`,
    };
    const xml = writer.xml();
    expect(xml).toBe(
      SOURCE_STYLES.replace(`<numFmts count="2">`, `<numFmts count="4">`)
        .replace("</numFmts>", () => `${inserted.numFmts}</numFmts>`)
        .replace(`<fonts count="2" `, `<fonts count="3" `)
        .replace("</fonts>", () => `${inserted.fonts}</fonts>`)
        .replace(`<fills count="3">`, `<fills count="4">`)
        .replace("</fills>", () => `${inserted.fills}</fills>`)
        .replace(`<borders count="1">`, `<borders count="2">`)
        .replace("</borders>", () => `${inserted.borders}</borders>`)
        .replace(`<cellXfs count="3">`, `<cellXfs count="9">`)
        .replace("</cellXfs>", () => `${inserted.cellXfs}</cellXfs>`),
    );
    // Removing the inserted runs and restoring the counts gives the source exactly.
    let restored = xml;
    for (const run of Object.values(inserted)) {
      restored = restored.replace(run, () => "");
    }
    restored = restored
      .replace(`<numFmts count="4">`, `<numFmts count="2">`)
      .replace(`<fonts count="3" `, `<fonts count="2" `)
      .replace(`<fills count="4">`, `<fills count="3">`)
      .replace(`<borders count="2">`, `<borders count="1">`)
      .replace(`<cellXfs count="9">`, `<cellXfs count="3">`);
    expect(restored).toBe(SOURCE_STYLES);
    // A written style reads back as the style it was written for (an absent font as font 0).
    const reread = readXlsxStyles(xml, undefined).styles;
    const font0 = reread[0]!.font;
    expect(styleKey(reread[3]!)).toBe(styleKey(added[0]!));
    expect(styleKey(reread[4]!)).toBe(styleKey(added[1]!));
    expect(styleKey(reread[5]!)).toBe(styleKey({ ...added[2]!, font: font0 }));
    // (A theme colour reads back with the tint text written for it.)
    const tinted: SheetColorRef = { theme: "accent1", lumMod: 0.6, lumOff: 0.4, xlsxTint: "0.4" };
    expect(styleKey(reread[7]!)).toBe(styleKey({ ...added[4]!, fill: { kind: "solid", color: tinted }, font: font0 }));
    expect(styleKey(reread[8]!)).toBe(styleKey({ ...added[6]!, font: font0 }));
  });

  it("writes fonts in Excel's child order, a theme font as its typeface + <scheme>", () => {
    const appendedFont = (font: CellStyle["font"]): string => /<fonts[^>]*>.*(<font>.*?<\/font>)<\/fonts>/.exec(styleWriterXml(SOURCE_STYLES, { font }))![1]!;
    expect(appendedFont({ family: "Arial", sizePt: 10, bold: true, italic: true, strike: true, underline: "single", vertAlign: "subscript", color: "#336699" })).toBe(
      `<font><b/><i/><strike/><u/><vertAlign val="subscript"/><sz val="10"/><color rgb="FF336699"/><name val="Arial"/></font>`,
    );
    expect(appendedFont({ family: { theme: "major" }, sizePt: 18, color: { theme: "dk2" } })).toBe(
      `<font><sz val="18"/><color theme="3"/><name val="${resolveFont({ theme: "major" }, DEFAULT_THEME)}"/><family val="2"/><scheme val="major"/></font>`,
    );
  });

  it("inserts an absent collection in CT_Stylesheet order (here borders, before cellXfs; with no parent xf, border 0 needs no applyBorder)", () => {
    const bare = `<styleSheet xmlns="${MAIN}"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="1"><fill><patternFill/></fill></fills><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" xfId="0"/></cellXfs></styleSheet>`;
    expect(styleWriterXml(bare, { borders: { left: { style: "thin" } } })).toBe(
      `<styleSheet xmlns="${MAIN}"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="1"><fill><patternFill/></fill></fills>` +
        `<borders count="1"><border><left style="thin"><color indexed="64"/></left><right/><top/><bottom/><diagonal/></border></borders>` +
        `<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" xfId="0"/><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs></styleSheet>`,
    );
  });
});

describe("XlsxStyleWriter: from scratch", () => {
  it("builds font 0 and xf 0 from model.styles[0]: the blank workbook's is SPEC-121's FROM_SCRATCH_STYLES_XML", () => {
    const blank = new XlsxStyleWriter(undefined, createBlankWorkbook());
    expect(blank.xfFor(0)).toBe(0);
    expect(blank.changed).toBe(false);
    expect(blank.xml()).toBe(FROM_SCRATCH_STYLES_XML);
  });

  it("keeps another workbook's default font — and ignores a sourceXf, which names no xf of this part", () => {
    const styles: CellStyle[] = [{ font: { family: "Arial", sizePt: 10 } }, { font: { family: "Arial", sizePt: 10, bold: true }, sourceXf: 7 }];
    const writer = new XlsxStyleWriter(undefined, { styles, theme: undefined });
    expect([writer.xfFor(0), writer.xfFor(1)]).toEqual([0, 1]);
    const xml = writer.xml();
    expect(xml).toContain(`<fonts count="2"><font><sz val="10"/><name val="Arial"/></font><font><b/><sz val="10"/><name val="Arial"/></font></fonts>`);
    expect(xml).toContain(`<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs>`);
    const read = readXlsxStyles(xml, undefined);
    expect(read.defaultFont).toEqual({ family: "Arial", sizePt: 10 });
    expect(read.styles.map(styleKey)).toEqual(styles.map(styleKey));
  });
});

// --- SPEC-123: named cell styles -------------------------------------------------------

/** An Excel-shaped styles part with named styles: the built-ins Normal, Good
 * and Heading 1; the custom "Invoice total" (a number format, font, borders
 * and alignment — its fill excluded with "false", its protection with "0");
 * a hidden built-in; a second cellStyle on Good's xfId (skipped). Its
 * cellXfs: Normal, Good, Heading 1 centred, Invoice total, an xf without
 * xfId. */
const NAMED_SOURCE =
  `${DECLARATION}\r\n<styleSheet xmlns="${MAIN}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="${X14AC}">` +
  `<numFmts count="1"><numFmt numFmtId="164" formatCode="0.0%"/></numFmts>` +
  `<fonts count="4" x14ac:knownFonts="1"><font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>` +
  `<font><sz val="11"/><color rgb="FF006100"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>` +
  `<font><b/><sz val="15"/><color theme="3"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>` +
  `<font><b/><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font></fonts>` +
  `<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>` +
  `<fill><patternFill patternType="solid"><fgColor rgb="FFC6EFCE"/><bgColor indexed="64"/></patternFill></fill></fills>` +
  `<borders count="3"><border><left/><right/><top/><bottom/><diagonal/></border>` +
  `<border><left/><right/><top/><bottom style="thick"><color theme="4"/></bottom><diagonal/></border>` +
  `<border><left/><right/><top style="thin"><color theme="4"/></top><bottom style="double"><color theme="4"/></bottom><diagonal/></border></borders>` +
  `<cellStyleXfs count="5"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>` +
  `<xf numFmtId="0" fontId="1" fillId="2" borderId="0" applyNumberFormat="0" applyBorder="0" applyAlignment="0" applyProtection="0"/>` +
  `<xf numFmtId="0" fontId="2" fillId="0" borderId="1" applyNumberFormat="0" applyFill="0" applyAlignment="0" applyProtection="0"/>` +
  `<xf numFmtId="164" fontId="3" fillId="2" borderId="2" applyFill="false" applyProtection="0"><alignment horizontal="right"/></xf>` +
  `<xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyFont="0" applyFill="0" applyBorder="0" applyAlignment="0" applyProtection="0"/></cellStyleXfs>` +
  `<cellXfs count="5"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>` +
  `<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="1"/>` +
  `<xf numFmtId="0" fontId="2" fillId="0" borderId="1" xfId="2" applyAlignment="1"><alignment horizontal="center"/></xf>` +
  `<xf numFmtId="164" fontId="3" fillId="0" borderId="2" xfId="3"><alignment horizontal="right"/></xf>` +
  `<xf numFmtId="0" fontId="3" fillId="0" borderId="0" applyFont="1"/></cellXfs>` +
  `<cellStyles count="6"><cellStyle name="Good" xfId="1" builtinId="26"/><cellStyle name="Heading 1" xfId="2" builtinId="16"/>` +
  `<cellStyle name="Invoice total" xfId="3"/><cellStyle name="Followed Hyperlink" xfId="4" builtinId="9" hidden="1"/>` +
  `<cellStyle name="Good 2" xfId="1"/><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>` +
  `<dxfs count="0"/><tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>` +
  `<extLst><ext uri="{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:slicerStyles defaultSlicerStyle="SlicerStyleLight1"/></ext></extLst></styleSheet>`;

const NORMAL_FONT: CellFont = { family: { theme: "minor" }, sizePt: 11, color: { theme: "dk1" } };
const ACCENT1 = { theme: "accent1" as const };
const INVOICE_TOTAL: CellStyle = {
  numberFormat: "0.0%",
  font: BOLD_MINOR,
  borders: { top: { style: "thin", color: ACCENT1 }, bottom: { style: "double", color: ACCENT1 } },
  alignment: { horizontal: "right" },
};

/** NAMED_SOURCE's named styles, as read. */
const NAMED_SOURCE_STYLES: NamedCellStyle[] = [
  {
    name: "Good",
    builtinId: 26,
    style: { font: { ...NORMAL_FONT, color: "#006100" }, fill: { kind: "solid", color: "#c6efce" } },
    includes: ["font", "fill"],
    sourceXfId: 1,
  },
  {
    name: "Heading 1",
    builtinId: 16,
    style: { font: { family: { theme: "minor" }, sizePt: 15, bold: true, color: { theme: "dk2" } }, borders: { bottom: { style: "thick", color: ACCENT1 } } },
    includes: ["font", "borders"],
    sourceXfId: 2,
  },
  { name: "Invoice total", style: INVOICE_TOTAL, includes: ["numberFormat", "font", "borders", "alignment"], sourceXfId: 3 },
  { name: "Followed Hyperlink", builtinId: 9, style: {}, includes: ["numberFormat"], hidden: true, sourceXfId: 4 },
  { name: "Normal", builtinId: 0, style: { font: NORMAL_FONT }, includes: [...STYLE_ASPECTS], sourceXfId: 0 },
];

/** A built-in as formatting.ts appends it on first use: the pinned definition. */
function appendedBuiltin(name: string): NamedCellStyle {
  const builtin = builtinCellStyle(name)!;
  return { name, builtinId: builtin.builtinId, style: structuredClone(builtin.style), includes: [...builtin.includes] };
}

/** A custom definition — an included number format, font, alignment and
 * protection, its fill and borders excluded — with a name to escape. */
const SUBTOTAL: NamedCellStyle = {
  name: `Sub & "total"`,
  style: { numberFormat: "#,##0.00", font: { italic: true }, alignment: { horizontal: "right" }, protection: { locked: false } },
  includes: ["numberFormat", "font", "alignment", "protection"],
};
const SUBTOTAL_NAME_ATTR = `Sub &amp; &quot;total&quot;`;

/** Cells in Bad and SUBTOTAL as applyNamedStyle leaves them (a font is the
 * definition's over Normal's). */
const BAD_CELL: CellStyle = { font: { ...NORMAL_FONT, color: "#9c0006" }, fill: { kind: "solid", color: "#ffc7ce" }, namedStyle: "Bad" };
const SUBTOTAL_CELL: CellStyle = {
  numberFormat: "#,##0.00",
  font: { ...NORMAL_FONT, italic: true },
  alignment: { horizontal: "right" },
  protection: { locked: false },
  namedStyle: SUBTOTAL.name,
};

/** A definition as it reads back once written at `sourceXfId`: its font
 * completed by Normal's. */
function writtenDefinition(definition: NamedCellStyle, sourceXfId: number): NamedCellStyle {
  const { sourceXfId: _source, ...rest } = definition;
  const style = definition.includes.includes("font") ? { ...rest.style, font: { ...NORMAL_FONT, ...rest.style.font } } : rest.style;
  return { ...rest, style, sourceXfId };
}

/** A writer over `source` whose model holds the source's styles and named
 * styles plus `added` and `named`; `first` is the first added style's id. */
function namedWriter(added: CellStyle[], named: NamedCellStyle[] = [], source = NAMED_SOURCE): { writer: XlsxStyleWriter; first: number } {
  const table = readXlsxStyles(source, undefined);
  const model = { styles: [...table.styles, ...added], theme: undefined, namedStyles: [...(table.namedStyles ?? []), ...named] };
  return { writer: new XlsxStyleWriter(source, model), first: table.styles.length };
}

/** A collection's content: the text between its start and close tags. */
function contentOf(xml: string, local: string): string {
  return new RegExp(`<${local}\\b[^>]*>(.*?)</${local}>`, "s").exec(xml)?.[1] ?? "";
}

/** The collections the writer appends to, and the child each counts. */
const COLLECTION_CHILD: Readonly<Record<string, string>> = {
  numFmts: "numFmt",
  fonts: "font",
  fills: "fill",
  borders: "border",
  cellStyleXfs: "xf",
  cellXfs: "xf",
  cellStyles: "cellStyle",
};

/** NAMED_SOURCE with `runs` appended to its collections, their counts updated. */
function namedSourceWith(runs: Partial<Record<string, string>>): string {
  let xml = NAMED_SOURCE;
  for (const [local, run] of Object.entries(runs)) {
    const total = (contentOf(xml, local) + run).match(new RegExp(`<${COLLECTION_CHILD[local]}\\b`, "g"))!.length;
    xml = xml.replace(new RegExp(`(<${local}\\b[^>]*?count=")\\d+(")`), `$1${total}$2`).replace(`</${local}>`, () => `${run}</${local}>`);
  }
  return xml;
}

describe("SPEC-123 named cell styles: reading cellStyleXfs and cellStyles", () => {
  it("reads one named style per cellStyle in file order — name, builtinId, hidden, sourceXfId = xfId — the first per xfId winning", () => {
    const { namedStyles } = readXlsxStyles(NAMED_SOURCE, undefined);
    expect(namedStyles).toEqual(NAMED_SOURCE_STYLES);
    // "Good 2" names Good's xfId after Good: skipped.
    expect(namedStyles!.map((entry) => entry.name)).not.toContain("Good 2");
  });

  it('defines each by its cellStyleXfs xf: an absent applyX includes the aspect, "0"/"false" excludes it; style keeps the included aspects with a value', () => {
    /** One cellStyleXfs xf's definition, over fonts [Calibri 11, bold red],
     * fills [none, gray125, yellow], borders [empty, bottom thin] and numFmt 164. */
    const definitionOf = (styleXf: string): Pick<NamedCellStyle, "style" | "includes"> => {
      const xml = stylesXml({
        fonts: [`<font><sz val="11"/><name val="Calibri"/></font>`, `<font><b/><sz val="11"/><color rgb="FFFF0000"/></font>`],
        fills: [`<fill><patternFill patternType="none"/></fill>`, `<fill><patternFill patternType="gray125"/></fill>`, `<fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/></patternFill></fill>`],
        borders: [`<border/>`, `<border><bottom style="thin"/></border>`],
        xfs: [`<xf fontId="0" fillId="0" borderId="0" xfId="0"/>`],
        extra: `<cellStyleXfs count="1">${styleXf}</cellStyleXfs><cellStyles count="1"><cellStyle name="S" xfId="0"/></cellStyles>`,
      }).replace("<fonts", `<numFmts count="1"><numFmt numFmtId="164" formatCode="0.0%"/></numFmts><fonts`);
      const { style, includes } = readXlsxStyles(xml, undefined).namedStyles![0]!;
      return { style, includes };
    };
    const ids = `numFmtId="164" fontId="1" fillId="2" borderId="1"`;
    const children = `<alignment horizontal="center"/><protection locked="0"/>`;
    const everything: CellStyle = {
      numberFormat: "0.0%",
      font: { bold: true, sizePt: 11, color: "#ff0000" },
      fill: { kind: "solid", color: "#ffff00" },
      borders: { bottom: { style: "thin" } },
      alignment: { horizontal: "center" },
      protection: { locked: false },
    };
    expect(definitionOf(`<xf ${ids}>${children}</xf>`)).toEqual({ style: everything, includes: [...STYLE_ASPECTS] });
    const flags: [StyleAspect, string][] = [
      ["numberFormat", "applyNumberFormat"],
      ["font", "applyFont"],
      ["fill", "applyFill"],
      ["borders", "applyBorder"],
      ["alignment", "applyAlignment"],
      ["protection", "applyProtection"],
    ];
    for (const [aspect, flag] of flags) {
      const { [aspect]: _excluded, ...rest } = everything;
      for (const off of ["0", "false"]) {
        expect(definitionOf(`<xf ${ids} ${flag}="${off}">${children}</xf>`), `${flag}="${off}"`).toEqual({
          style: rest,
          includes: STYLE_ASPECTS.filter((other) => other !== aspect),
        });
      }
      for (const on of ["1", "true"]) {
        expect(definitionOf(`<xf ${ids} ${flag}="${on}">${children}</xf>`), `${flag}="${on}"`).toEqual({ style: everything, includes: [...STYLE_ASPECTS] });
      }
    }
    // An included aspect without a value is listed and holds nothing; a
    // locale-dependent built-in keeps its builtinFormatId; a style xf's
    // quotePrefix is never read.
    expect(definitionOf(`<xf numFmtId="0" fontId="9" fillId="0" borderId="0"/>`)).toEqual({ style: {}, includes: [...STYLE_ASPECTS] });
    expect(definitionOf(`<xf numFmtId="14" fontId="9" fillId="0" borderId="0" quotePrefix="1"/>`).style).toEqual({ numberFormat: "m/d/yyyy", builtinFormatId: 14 });
    expect(definitionOf(`<xf numFmtId="14" fontId="9" fillId="0" borderId="0" applyNumberFormat="0"/>`).style).toEqual({});
  });

  it('gives a cellStyle whose xfId has no cellStyleXfs entry an empty definition, skips one without a name or an xfId, and reads hidden "1"/"true"', () => {
    const xml = stylesXml({
      extra:
        `<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>` +
        `<cellStyles count="6"><cellStyle name="Normal" xfId="0" builtinId="0" hidden="0"/><cellStyle xfId="2"/><cellStyle name="Two" xfId="2" hidden="true"/>` +
        `<cellStyle name="No xfId"/><cellStyle name="Bad xfId" xfId="x"/><cellStyle name="Far" xfId="7" builtinId="53" customBuiltin="1" hidden="1"/></cellStyles>`,
    });
    expect(readXlsxStyles(xml, undefined).namedStyles).toEqual([
      { name: "Normal", builtinId: 0, style: { font: { family: "Calibri", sizePt: 11 } }, includes: [...STYLE_ASPECTS], sourceXfId: 0 },
      { name: "Two", style: {}, includes: [], hidden: true, sourceXfId: 2 },
      { name: "Far", builtinId: 53, style: {}, includes: [], hidden: true, sourceXfId: 7 },
    ]);
  });

  it("names each cellXfs style by its xfId — absent with no xfId, an xfId naming no cellStyle, or Normal's; the cellXfs apply flags stay ignored", () => {
    const { styles } = readXlsxStyles(NAMED_SOURCE, undefined);
    expect(styles.map((style) => style.namedStyle)).toEqual([undefined, "Good", "Heading 1", "Invoice total", undefined]);
    expect(styles[3]).toEqual({ ...INVOICE_TOTAL, namedStyle: "Invoice total", sourceXf: 3 });
    expect(styles[2]).toEqual({ ...NAMED_SOURCE_STYLES[1]!.style, alignment: { horizontal: "center" }, namedStyle: "Heading 1", sourceXf: 2 });
    // An xfId with an entry but no cellStyle, one with neither, and one a
    // cellStyle names without an entry (its empty definition's name).
    const xfs = [3, 9, 7].map((xfId) => `<xf fontId="0" fillId="0" borderId="0" xfId="${xfId}"/>`);
    const extra =
      `<cellStyleXfs count="4"><xf/><xf/><xf/><xf/></cellStyleXfs>` +
      `<cellStyles count="2"><cellStyle name="Normal" xfId="0" builtinId="0"/><cellStyle name="Far" xfId="7"/></cellStyles>`;
    expect(readXlsxStyles(stylesXml({ xfs, extra }), undefined).styles.map((style) => style.namedStyle)).toEqual([undefined, undefined, "Far"]);
  });

  it('takes Normal to be builtinId 0 whatever its name, else — when no cellStyle has builtinId 0 — the one named "Normal"', () => {
    const xfs = [`<xf fontId="0" fillId="0" borderId="0" xfId="0"/>`, `<xf fontId="0" fillId="0" borderId="0" xfId="1"/>`];
    const styleXfs = `<cellStyleXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/><xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyFont="0"/></cellStyleXfs>`;
    const namesOf = (cellStyles: string): (string | undefined)[] =>
      readXlsxStyles(stylesXml({ xfs, extra: `${styleXfs}<cellStyles>${cellStyles}</cellStyles>` }), undefined).styles.map((style) => style.namedStyle);
    expect(namesOf(`<cellStyle name="Standard" xfId="0" builtinId="0"/><cellStyle name="Akzent" xfId="1"/>`)).toEqual([undefined, "Akzent"]);
    expect(namesOf(`<cellStyle name="Normal" xfId="0"/><cellStyle name="Accent" xfId="1"/>`)).toEqual([undefined, "Accent"]);
    expect(namesOf(`<cellStyle name="Accent" xfId="0"/><cellStyle name="Normal" xfId="1" builtinId="0"/>`)).toEqual(["Accent", undefined]);
  });

  it("reads no named styles and no namedStyle from a part without cellStyles", () => {
    const table = readXlsxStyles(stylesXml({ xfs: [`<xf fontId="0" fillId="0" borderId="0" xfId="1"/>`], extra: `<cellStyleXfs count="2"><xf/><xf/></cellStyleXfs>` }), undefined);
    expect(table).not.toHaveProperty("namedStyles");
    expect(table.styles[0]!.namedStyle).toBeUndefined();
  });

  it("hands spec117-styles.xlsx's Good, Heading 1 and Normal on through parseXlsx", async () => {
    const { namedStyles } = await parsed();
    expect(namedStyles).toEqual([
      {
        name: "Good",
        builtinId: 26,
        style: { font: { ...NORMAL_FONT, color: "#006100" }, fill: { kind: "solid", color: "#c6efce" } },
        includes: ["font", "fill"],
        sourceXfId: 1,
      },
      NAMED_SOURCE_STYLES[1],
      { name: "Normal", builtinId: 0, style: { font: NORMAL_FONT }, includes: [...STYLE_ASPECTS], sourceXfId: 0 },
    ]);
  });
});

describe("SPEC-123 named cell styles: the style writer", () => {
  it("writes a style naming a source style with its sourceXfId, the apply flags against that parent — appending no named style", () => {
    const { writer, first } = namedWriter([
      { font: { ...NORMAL_FONT, color: "#006100", bold: true }, fill: { kind: "solid", color: "#c6efce" }, namedStyle: "Good" },
      { ...INVOICE_TOTAL, numberFormat: "0.00%", namedStyle: "Invoice total" },
      { font: NORMAL_FONT, namedStyle: "Heading 1" },
    ]);
    expect([0, 1, 2].map((k) => writer.xfFor(first + k))).toEqual([5, 6, 7]);
    expect(writer.xml()).toBe(
      namedSourceWith({
        fonts: `<font><b/><sz val="11"/><color rgb="FF006100"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>`,
        cellXfs:
          `<xf numFmtId="0" fontId="4" fillId="2" borderId="0" xfId="1" applyFont="1"/>` +
          // The parent xf points at fill 2 (excluded there): the style's none fill differs.
          `<xf numFmtId="10" fontId="3" fillId="0" borderId="2" xfId="3" applyNumberFormat="1" applyFill="1"><alignment horizontal="right"/></xf>` +
          `<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="2" applyFont="1" applyBorder="1"/>`,
      }),
    );
  });

  it('appends a built-in first applied: a cellStyleXfs xf (applyX="0" and Normal\'s ids for each excluded aspect) and a cellStyle with its builtinId, counts updated', () => {
    const { writer, first } = namedWriter([BAD_CELL], [appendedBuiltin("Bad")]);
    expect(writer.xfFor(first)).toBe(5);
    expect(writer.changed).toBe(true);
    expect(writer.xml()).toBe(
      namedSourceWith({
        fonts: `<font><sz val="11"/><color rgb="FF9C0006"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>`,
        fills: `<fill><patternFill patternType="solid"><fgColor rgb="FFFFC7CE"/><bgColor indexed="64"/></patternFill></fill>`,
        cellStyleXfs: `<xf numFmtId="0" fontId="4" fillId="3" borderId="0" applyNumberFormat="0" applyBorder="0" applyAlignment="0" applyProtection="0"/>`,
        // The style says what its parent says: no apply flag.
        cellXfs: `<xf numFmtId="0" fontId="4" fillId="3" borderId="0" xfId="5"/>`,
        cellStyles: `<cellStyle name="Bad" xfId="5" builtinId="27"/>`,
      }),
    );
  });

  it("appends a custom-named definition without builtinId — its number format, alignment and protection written, the excluded fill and border Normal's — and flags what a style formatted beyond it changes", () => {
    const formatted: CellStyle = { ...SUBTOTAL_CELL, numberFormat: "0.000", fill: { kind: "solid", color: "#ffff00" } };
    const { writer, first } = namedWriter([SUBTOTAL_CELL, formatted], [SUBTOTAL]);
    expect([writer.xfFor(first), writer.xfFor(first + 1)]).toEqual([5, 6]);
    const children = `<alignment horizontal="right"/><protection locked="0"/>`;
    expect(writer.xml()).toBe(
      namedSourceWith({
        numFmts: `<numFmt numFmtId="165" formatCode="0.000"/>`,
        fonts: `<font><i/><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>`,
        fills: `<fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/><bgColor indexed="64"/></patternFill></fill>`,
        cellStyleXfs: `<xf numFmtId="4" fontId="4" fillId="0" borderId="0" applyFill="0" applyBorder="0">${children}</xf>`,
        cellXfs:
          `<xf numFmtId="4" fontId="4" fillId="0" borderId="0" xfId="5">${children}</xf>` +
          `<xf numFmtId="165" fontId="4" fillId="3" borderId="0" xfId="5" applyNumberFormat="1" applyFill="1">${children}</xf>`,
        cellStyles: `<cellStyle name="${SUBTOTAL_NAME_ATTR}" xfId="5"/>`,
      }),
    );
  });

  it("appends a named style once, however many styles name it", () => {
    const { writer, first } = namedWriter([BAD_CELL, { ...BAD_CELL, font: { ...BAD_CELL.font, bold: true } }, BAD_CELL], [appendedBuiltin("Bad")]);
    expect([0, 1, 2].map((k) => writer.xfFor(first + k))).toEqual([5, 6, 5]);
    const xml = writer.xml();
    expect(contentOf(xml, "cellStyleXfs").match(/<xf\b/g)).toHaveLength(6);
    expect(contentOf(xml, "cellStyles").match(/<cellStyle\b/g)).toHaveLength(7);
    expect(contentOf(xml, "cellXfs").endsWith(`<xf numFmtId="0" fontId="5" fillId="3" borderId="0" xfId="5" applyFont="1"/>`)).toBe(true);
  });

  it('writes an absent or "Normal" namedStyle, or one whose definition has builtinId 0, with Normal\'s xfId — the builtinId-0 cellStyle\'s, wherever it is', () => {
    const plain = namedWriter([{ font: BOLD_MINOR }, { font: BOLD_MINOR, namedStyle: "Normal" }]);
    expect([plain.writer.xfFor(plain.first), plain.writer.xfFor(plain.first + 1)]).toEqual([5, 5]);
    expect(plain.writer.xml()).toBe(namedSourceWith({ cellXfs: `<xf numFmtId="0" fontId="3" fillId="0" borderId="0" xfId="0" applyFont="1"/>` }));
    // A localised Normal on xfId 1 (Heading 1 on 0): the apply flags compare against it.
    const localised = SOURCE_STYLES.replace(
      `<cellStyle name="Heading 1" xfId="1" builtinId="16"/><cellStyle name="Normal" xfId="0" builtinId="0"/>`,
      `<cellStyle name="Heading 1" xfId="0" builtinId="16"/><cellStyle name="Standard" xfId="1" builtinId="0"/>`,
    );
    const moved = namedWriter(
      [{ font: BOLD_MINOR }, { font: BOLD_MINOR, namedStyle: "Normal" }, { font: BOLD_MINOR, namedStyle: "Standard" }, { font: BOLD_MINOR, numberFormat: "0.000" }],
      [],
      localised,
    );
    expect([0, 1, 2, 3].map((k) => moved.writer.xfFor(moved.first + k))).toEqual([3, 3, 3, 4]);
    expect(moved.writer.xml()).toBe(
      localised
        .replace(`<cellXfs count="3">`, `<cellXfs count="5">`)
        .replace("</cellXfs>", `<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="1"/><xf numFmtId="164" fontId="1" fillId="0" borderId="0" xfId="1" applyNumberFormat="1"/></cellXfs>`),
    );
  });

  it("falls back, for a name the workbook does not define, to the source cellStyle of that name, else a built-in's pinned definition, else Normal", () => {
    // A model without namedStyles (stored before they were read) still finds the source's cellStyle by name.
    const styles = [...readXlsxStyles(NAMED_SOURCE, undefined).styles, { font: BOLD_MINOR, namedStyle: "Invoice total" }];
    const unlisted = new XlsxStyleWriter(NAMED_SOURCE, { styles, theme: undefined });
    expect(unlisted.xfFor(5)).toBe(5);
    expect(unlisted.xml()).toBe(
      namedSourceWith({ cellXfs: `<xf numFmtId="0" fontId="3" fillId="0" borderId="0" xfId="3" applyNumberFormat="1" applyFill="1" applyBorder="1" applyAlignment="1"/>` }),
    );
    // A built-in nothing appended yet is written with its pinned definition;
    // a name nothing defines is Normal's, and appends no named style.
    const neutral: CellStyle = { font: { ...NORMAL_FONT, color: "#9c5700" }, fill: { kind: "solid", color: "#ffeb9c" }, namedStyle: "Neutral" };
    const { writer, first } = namedWriter([neutral, { font: BOLD_MINOR, namedStyle: "Mystery" }]);
    expect([writer.xfFor(first), writer.xfFor(first + 1)]).toEqual([5, 6]);
    expect(writer.xml()).toBe(
      namedSourceWith({
        fonts: `<font><sz val="11"/><color rgb="FF9C5700"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>`,
        fills: `<fill><patternFill patternType="solid"><fgColor rgb="FFFFEB9C"/><bgColor indexed="64"/></patternFill></fill>`,
        cellStyleXfs: `<xf numFmtId="0" fontId="4" fillId="3" borderId="0" applyNumberFormat="0" applyBorder="0" applyAlignment="0" applyProtection="0"/>`,
        cellXfs: `<xf numFmtId="0" fontId="4" fillId="3" borderId="0" xfId="5"/><xf numFmtId="0" fontId="3" fillId="0" borderId="0" xfId="0" applyFont="1"/>`,
        cellStyles: `<cellStyle name="Neutral" xfId="5" builtinId="28"/>`,
      }),
    );
  });

  it("gives a source without cellStyles one holding Normal first, and one without cellStyleXfs Normal's xf first — each in CT_Stylesheet order", () => {
    const head =
      `<styleSheet xmlns="${MAIN}"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>` +
      `<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>` +
      `<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>`;
    const styleXfs = `<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>`;
    const cellXfs = `<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>`;
    const tail = `<dxfs count="0"/></styleSheet>`;
    const good: CellStyle = { font: { family: "Calibri", sizePt: 11, color: "#006100" }, fill: { kind: "solid", color: "#c6efce" }, namedStyle: "Good" };
    const expected =
      `<styleSheet xmlns="${MAIN}"><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><sz val="11"/><color rgb="FF006100"/><name val="Calibri"/></font></fonts>` +
      `<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>` +
      `<fill><patternFill patternType="solid"><fgColor rgb="FFC6EFCE"/><bgColor indexed="64"/></patternFill></fill></fills>` +
      `<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>` +
      `<cellStyleXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>` +
      `<xf numFmtId="0" fontId="1" fillId="2" borderId="0" applyNumberFormat="0" applyBorder="0" applyAlignment="0" applyProtection="0"/></cellStyleXfs>` +
      `<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="1"/></cellXfs>` +
      `<cellStyles count="2"><cellStyle name="Normal" xfId="0" builtinId="0"/><cellStyle name="Good" xfId="1" builtinId="26"/></cellStyles>` +
      tail;
    for (const source of [head + styleXfs + cellXfs + tail, head + cellXfs + tail]) {
      const { writer, first } = namedWriter([good], [appendedBuiltin("Good")], source);
      expect(writer.xfFor(first)).toBe(1);
      expect(writer.xml()).toBe(expected);
    }
  });

  it("fills an index a source cellStyle or cellXfs xf names without an entry with an every-aspect-excluded xf before appending, so neither changes meaning", () => {
    // "Orphan" names xfId 5 and a cellXfs xf xfId 6 — past the five entries.
    const dangling = NAMED_SOURCE.replace(`<cellStyles count="6">`, `<cellStyles count="7">`)
      .replace(`<cellStyle name="Normal" xfId="0"`, `<cellStyle name="Orphan" xfId="5"/><cellStyle name="Normal" xfId="0"`)
      .replace(`<cellXfs count="5">`, `<cellXfs count="6">`)
      .replace("</cellXfs>", `<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="6"/></cellXfs>`);
    const before = readXlsxStyles(dangling, undefined);
    expect(before.namedStyles!.find((entry) => entry.name === "Orphan")).toEqual({ name: "Orphan", style: {}, includes: [], sourceXfId: 5 });
    const { writer, first } = namedWriter([BAD_CELL], [appendedBuiltin("Bad")], dangling);
    expect(writer.xfFor(first)).toBe(6);
    const xml = writer.xml();
    const empty = `<xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyNumberFormat="0" applyFont="0" applyFill="0" applyBorder="0" applyAlignment="0" applyProtection="0"/>`;
    expect(contentOf(xml, "cellStyleXfs")).toBe(
      `${contentOf(dangling, "cellStyleXfs")}${empty}${empty}<xf numFmtId="0" fontId="4" fillId="3" borderId="0" applyNumberFormat="0" applyBorder="0" applyAlignment="0" applyProtection="0"/>`,
    );
    expect(contentOf(xml, "cellStyles").endsWith(`<cellStyle name="Bad" xfId="7" builtinId="27"/>`)).toBe(true);
    // Read back, the orphan keeps its empty definition and the dangling xf
    // its absent namedStyle; Bad is not shadowed.
    const after = readXlsxStyles(xml, undefined);
    expect(after.namedStyles).toEqual([...before.namedStyles!, writtenDefinition(appendedBuiltin("Bad"), 7)]);
    expect(after.styles.slice(0, before.styles.length)).toEqual(before.styles);
    expect(after.styles[6]!.namedStyle).toBe("Bad");
  });

  it("from scratch, appends every definition used after the base's Normal — a read definition's sourceXfId names nothing of the new part", () => {
    const goodRead = NAMED_SOURCE_STYLES[0]!; // sourceXfId 1, of another part
    const model = { styles: [...createBlankWorkbook().styles, { ...goodRead.style, namedStyle: "Good" }, SUBTOTAL_CELL], theme: undefined, namedStyles: [goodRead, SUBTOTAL] };
    const writer = new XlsxStyleWriter(undefined, model);
    expect([0, 1, 2].map((id) => writer.xfFor(id))).toEqual([0, 1, 2]);
    const xml = writer.xml();
    const children = `<alignment horizontal="right"/><protection locked="0"/>`;
    expect(contentOf(xml, "cellStyleXfs")).toBe(
      `<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>` +
        `<xf numFmtId="0" fontId="1" fillId="2" borderId="0" applyNumberFormat="0" applyBorder="0" applyAlignment="0" applyProtection="0"/>` +
        `<xf numFmtId="4" fontId="2" fillId="0" borderId="0" applyFill="0" applyBorder="0">${children}</xf>`,
    );
    expect(contentOf(xml, "cellXfs")).toBe(
      `<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="1"/>` +
        `<xf numFmtId="4" fontId="2" fillId="0" borderId="0" xfId="2">${children}</xf>`,
    );
    expect(xml).toContain(
      `<cellStyles count="3"><cellStyle name="Normal" xfId="0" builtinId="0"/><cellStyle name="Good" xfId="1" builtinId="26"/>` +
        `<cellStyle name="${SUBTOTAL_NAME_ATTR}" xfId="2"/></cellStyles>`,
    );
    const reread = readXlsxStyles(xml, undefined);
    expect(reread.namedStyles).toEqual([
      { name: "Normal", builtinId: 0, style: { font: NORMAL_FONT }, includes: [...STYLE_ASPECTS], sourceXfId: 0 },
      writtenDefinition(goodRead, 1),
      writtenDefinition(SUBTOTAL, 2),
    ]);
    expect(reread.styles.map((style) => style.namedStyle)).toEqual([undefined, "Good", SUBTOTAL.name]);
  });

  it("appends a built-in a file named in its language under that name, with the file's builtinId, once for either name", () => {
    const gut: NamedCellStyle = { ...NAMED_SOURCE_STYLES[0]!, name: "Gut" }; // builtinId 26, read from another part
    const styles = [...createBlankWorkbook().styles, { ...gut.style, namedStyle: "Gut" }, { ...gut.style, namedStyle: "Good" }];
    const writer = new XlsxStyleWriter(undefined, { styles, theme: undefined, namedStyles: [gut] });
    expect([1, 2].map((id) => writer.xfFor(id))).toEqual([1, 1]);
    const xml = writer.xml();
    expect(contentOf(xml, "cellStyles")).toBe(`<cellStyle name="Normal" xfId="0" builtinId="0"/><cellStyle name="Gut" xfId="1" builtinId="26"/>`);
    expect(readXlsxStyles(xml, undefined).styles[1]!.namedStyle).toBe("Gut");
  });

  it("writes what reads back: the named styles — each appended one at its new index, its font completed by Normal's — and every style's namedStyle; every source entry keeps its index and bytes", () => {
    const named = [appendedBuiltin("Bad"), appendedBuiltin("Title"), SUBTOTAL];
    const added: CellStyle[] = [
      BAD_CELL,
      { font: { family: { theme: "major" }, sizePt: 18, color: { theme: "dk2" } }, namedStyle: "Title" },
      SUBTOTAL_CELL,
      { ...SUBTOTAL_CELL, font: { ...SUBTOTAL_CELL.font, bold: true } }, // direct formatting over it
      { ...NAMED_SOURCE_STYLES[0]!.style, alignment: { wrap: true }, namedStyle: "Good" }, // a source style
      { ...INVOICE_TOTAL, fill: { kind: "solid", color: "#ffc7ce" }, namedStyle: "Invoice total" },
      { font: BOLD_MINOR }, // Normal
    ];
    const { writer, first } = namedWriter(added, named);
    const xfs = added.map((_, k) => writer.xfFor(first + k));
    const xml = writer.xml();
    const reread = readXlsxStyles(xml, undefined);
    expect(reread.namedStyles).toEqual([...NAMED_SOURCE_STYLES, writtenDefinition(named[0]!, 5), writtenDefinition(named[1]!, 6), writtenDefinition(named[2]!, 7)]);
    added.forEach((style, k) => {
      expect(reread.styles[xfs[k]!]!.namedStyle, String(k)).toBe(style.namedStyle);
      expect(styleKey(reread.styles[xfs[k]!]!), String(k)).toBe(styleKey(style));
    });
    // Append-only: each collection's source content heads the written one,
    // its count is the new total, and nothing outside the collections moved.
    for (const [local, child] of Object.entries(COLLECTION_CHILD)) {
      const content = contentOf(xml, local);
      expect(content.startsWith(contentOf(NAMED_SOURCE, local)), local).toBe(true);
      expect(new RegExp(`<${local}\\b[^>]*?count="(\\d+)"`).exec(xml)![1], local).toBe(String(content.match(new RegExp(`<${child}\\b`, "g"))!.length));
    }
    const outside = (text: string): string => Object.keys(COLLECTION_CHILD).reduce((rest, local) => rest.replace(new RegExp(`<${local}\\b.*?</${local}>`, "s"), ""), text);
    expect(outside(xml)).toBe(outside(NAMED_SOURCE));
  });

  it("survives a whole save: spec117-styles.xlsx with Bad first applied to B2 re-opens with Bad defined and B2 in it, Heading 1 and Good kept", async () => {
    const source = await buildSpec117StylesXlsxBytes();
    const model = await readSpreadsheetBytes(source, "xlsx", "book");
    expect(model.namedStyles?.map((entry) => entry.name)).toEqual(["Good", "Heading 1", "Normal"]);
    const bad = appendedBuiltin("Bad");
    const namedStyles = [...model.namedStyles!, bad];
    model.namedStyles = namedStyles;
    const badId = internStyle(model, BAD_CELL);
    const sheet = model.sheets[0]!;
    expect(getCell(sheet, 1, 1)?.styleId).toBe(SPEC117_XF.header);
    model.sheets[0] = {
      ...sheet,
      dirty: true,
      rows: sheet.rows.map((row) => (row.index !== 1 ? row : { ...row, cells: row.cells.map((cell) => (cell.col === 1 ? { ...cell, styleId: badId } : cell)) })),
    };
    const bytes = await writeXlsx(model);
    expect(await packageIntegrity(bytes)).toEqual([]);
    const reread = await readSpreadsheetBytes(bytes, "xlsx", "reread");
    expect(reread.namedStyles).toEqual([...namedStyles.slice(0, 3), writtenDefinition(bad, 3)]);
    const styleAt = (row: number, col: number): CellStyle => cellStyle(reread, getCell(reread.sheets[0]!, row, col)?.styleId ?? 0);
    expect(styleAt(1, 1).namedStyle).toBe("Bad");
    expect(styleAt(4, 2).namedStyle).toBe("Heading 1"); // C5
    expect(styleAt(6, 5).namedStyle).toBe("Good"); // F7
    // The source's own entries are untouched.
    const decoder = new TextDecoder();
    const before = decoder.decode((await unzip(source)).get("xl/styles.xml"));
    const after = decoder.decode((await unzip(bytes)).get("xl/styles.xml"));
    for (const local of Object.keys(COLLECTION_CHILD)) {
      expect(contentOf(after, local).startsWith(contentOf(before, local)), local).toBe(true);
    }
    expect(contentOf(after, "cellStyles")).toBe(`${contentOf(before, "cellStyles")}<cellStyle name="Bad" xfId="3" builtinId="27"/>`);
  });
});

// --- SPEC-129: differential styles ------------------------------------------------------

describe("SPEC-129 dxfs: Excel's quirks on read, each dxf kept verbatim", () => {
  const DXFS = [
    `<dxf><font><b val="0"/><i/><strike val="1"/><u val="none"/><color theme="4" tint="-0.249977111117893"/></font></dxf>`,
    `<dxf><font><b/><u/><color rgb="FF9C0006"/></font><fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf>`,
    `<dxf><fill><patternFill patternType="solid"><fgColor rgb="FF00B050"/></patternFill></fill></dxf>`,
    `<dxf><fill><patternFill patternType="darkGray"><fgColor rgb="FF0000FF"/><bgColor indexed="13"/></patternFill></fill></dxf>`,
    `<dxf><numFmt numFmtId="164" formatCode="0.0%"/><border><left style="thin"><color rgb="FFFF0000"/></left><bottom style="double"/></border></dxf>`,
    `<dxf><alignment horizontal="center"/><protection locked="0"/></dxf>`,
  ];
  const part = stylesXml({ extra: `<dxfs count="${DXFS.length}">${DXFS.join("")}</dxfs>` });

  it("reads tri-state flags, underline none, tinted theme colours, the bgColor fill quirk, patterns, number formats and borders by index", () => {
    const dxfs = readXlsxDxfs(part);
    expect(dxfs.map(({ sourceXml: _xml, ...dxf }) => dxf)).toEqual([
      { font: { bold: false, italic: true, strike: true, underline: "none", color: { theme: "accent1", lumMod: 1 - 0.249977111117893, xlsxTint: "-0.249977111117893" } } },
      { font: { bold: true, underline: "single", color: "#9c0006" }, fill: { kind: "solid", color: "#ffc7ce" } },
      { fill: { kind: "solid", color: "#00b050" } },
      { fill: { kind: "pattern", pattern: "darkGray", fg: "#0000ff", bg: INDEXED_COLORS[13] } },
      { numberFormat: "0.0%", borders: { left: { style: "thin", color: "#ff0000" }, bottom: { style: "double" } } },
      {},
    ]);
    // Each keeps its element verbatim — alignment and protection included.
    expect(dxfs.map((dxf) => dxf.sourceXml)).toEqual(DXFS);
    // The tinted theme colour resolves as SPEC-117's does (accent 1, darker 25%).
    expectNearColor(resolveColor(dxfs[0]!.font!.color!, DEFAULT_THEME), "#2f5597");
  });

  it("reads the same through readXlsxStyles; a part without dxfs, or not well-formed, reads none", () => {
    expect(readXlsxStyles(part, undefined).dxfs).toEqual(readXlsxDxfs(part));
    expect(readXlsxDxfs(stylesXml({}))).toEqual([]);
    expect(readXlsxDxfs("<styleSheet")).toEqual([]);
    expect(readXlsxDxfs(undefined)).toEqual([]);
  });
});

/** A blank workbook holding `dxfs`. */
function modelWith(dxfs: Dxf[]) {
  return { ...createBlankWorkbook(), dxfs };
}

describe("SPEC-129 dxfs: the writer appends past the source's, append-only", () => {
  const SOURCE_DXF = `<dxf><font><color rgb="FF9C0006"/></font><fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf>`;
  const source =
    `<styleSheet xmlns="${MAIN}"><numFmts count="1"><numFmt numFmtId="170" formatCode="0.0%"/></numFmts>` +
    `<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>` +
    `<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>` +
    `<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>` +
    `<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>` +
    `<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>` +
    `<dxfs count="1">${SOURCE_DXF}</dxfs><tableStyles count="0"/></styleSheet>`;
  const appended: Dxf[] = [
    { font: { color: "#006100", bold: false, underline: "none", strike: true }, fill: { kind: "solid", color: "#c6efce" } },
    { numberFormat: "0.0%" },
    { numberFormat: "#,##0.000", borders: { top: { style: "thin" }, right: { style: "medium", color: "#ff0000" } } },
    { fill: { kind: "pattern", pattern: "lightGray", fg: "#ff0000" } },
  ];

  it("leaves the part untouched until a rule writes a dxfId past the source's — an interned but unwritten dxf is not appended", () => {
    const writer = new XlsxStyleWriter(source, modelWith([...readXlsxDxfs(source), ...appended]));
    expect(writer.changed).toBe(false);
    expect(writer.xml()).toBe(source);
    expect(writer.dxfFor(0)).toBe(0);
    expect(writer.changed).toBe(false);
    expect(writer.xml()).toBe(source);
  });

  it("appends every dxf up to the highest written id, count updated, a solid fill as bgColor, numFmt ids reused or the next free ≥ 164; the source dxf verbatim", () => {
    const writer = new XlsxStyleWriter(source, modelWith([...readXlsxDxfs(source), ...appended]));
    expect(writer.dxfFor(3)).toBe(3);
    expect(writer.changed).toBe(true);
    const xml = writer.xml();
    expect(xml).toBe(
      source.replace(
        `<dxfs count="1">${SOURCE_DXF}</dxfs>`,
        `<dxfs count="4">${SOURCE_DXF}` +
          `<dxf><font><b val="0"/><strike/><u val="none"/><color rgb="FF006100"/></font><fill><patternFill><bgColor rgb="FFC6EFCE"/></patternFill></fill></dxf>` +
          `<dxf><numFmt numFmtId="170" formatCode="0.0%"/></dxf>` +
          `<dxf><numFmt numFmtId="171" formatCode="#,##0.000"/><border><right style="medium"><color rgb="FFFF0000"/></right><top style="thin"/></border></dxf></dxfs>`,
      ),
    );
    // Read back, the appended entries are the model's.
    expect(readXlsxDxfs(xml).map(({ sourceXml: _xml, ...dxf }) => dxf)).toEqual([readXlsxDxfs(source)[0]!, ...appended.slice(0, 3)].map(({ sourceXml: _xml, ...dxf }) => dxf));
  });

  it("creates <dxfs> after cellStyles, before tableStyles, when the part has none; from scratch too", () => {
    const bare = source.replace(`<dxfs count="1">${SOURCE_DXF}</dxfs>`, "");
    const writer = new XlsxStyleWriter(bare, modelWith([appended[0]!]));
    writer.dxfFor(0);
    expect(writer.xml()).toContain(`</cellStyles><dxfs count="1"><dxf><font>`);
    expect(writer.xml()).toMatch(/<\/dxf><\/dxfs><tableStyles count="0"\/><\/styleSheet>$/);
    const scratch = new XlsxStyleWriter(undefined, modelWith([appended[3]!, appended[1]!]));
    scratch.dxfFor(1);
    const xml = scratch.xml();
    expect(xml).toContain(`<dxfs count="2"><dxf><fill><patternFill patternType="lightGray"><fgColor rgb="FFFF0000"/>`);
    expect(xml).toContain(`<dxf><numFmt numFmtId="164" formatCode="0.0%"/></dxf></dxfs><tableStyles count="0"/>`);
    expect(readXlsxDxfs(xml).map(({ sourceXml: _xml, ...dxf }) => dxf)).toEqual([appended[3], appended[1]]);
  });
});
