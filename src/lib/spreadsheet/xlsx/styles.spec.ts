// SPEC-117 "Reading xlsx styles" and "Colours": the colour codec
// (xlsx/themeColor.ts) — Excel's theme-index swap, the tint ↔ lumMod/lumOff
// arithmetic both ways, the indexed palette with its override, 64, 65, rgb
// ARGB and 6-digit values, auto — with the pinned "Accent 1, Lighter 40%" =
// Excel's #8ea9db; and the style table (xlsx/styles.ts) row by row: fonts,
// fills (pattern and gradient), borders (all thirteen styles, diagonals,
// start/end), cellXfs one-to-one with copies, alignment, protection,
// quotePrefix, the override palette, the parts left unread, the default
// font — over synthetic parts and spec117-styles.xlsx.
import { describe, expect, it } from "vitest";
import { DEFAULT_THEME, hexToRgb, resolveColor, type DocumentTheme, type ModelColor } from "../../core/theme";
import { buildSpec117StylesXlsxBytes, SPEC117_BORDER_STYLES, SPEC117_STYLES_XLSX, SPEC117_XF } from "../fixtures/build";
import type { CellStyle, SheetColorRef } from "../model";
import { INDEXED_COLORS } from "../styles";
import { parseXlsx } from "./reader";
import { readXlsxStyles, XlsxFontReader } from "./styles";
import {
  modelColorToXlsxAttrs,
  tintTransforms,
  XLSX_THEME_SLOTS,
  xlsxColorToModel,
  xlsxThemeIndex,
  xlsxThemeSlot,
  type XlsxColorAttributes,
} from "./themeColor";

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

  it("copies a font equal to the default style's, ignores the apply* flags, xfId and the number format (SPEC-118)", () => {
    const { styles } = readXlsxStyles(xml, undefined);
    expect(styles[0]).toEqual({ font: { family: { theme: "minor" }, sizePt: 11 }, sourceXf: 0 });
    expect(styles[3]).toEqual({ font: { family: { theme: "minor" }, sizePt: 11 }, sourceXf: 3 });
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

describe("the parts left unread and the override palette", () => {
  const unread =
    `<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>` +
    `<cellStyles count="2"><cellStyle name="Good" xfId="0" builtinId="26"/><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>` +
    `<dxfs count="1"><dxf><font><b/></font><fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf></dxfs>` +
    `<tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>` +
    `<colors><mruColors><color rgb="FF66FF33"/></mruColors></colors>` +
    `<extLst><ext uri="{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}"><slicerStyles defaultSlicerStyle="SlicerStyleLight1"/></ext></extLst>`;

  it("leaves numFmts, cellStyleXfs/cellStyles, dxfs, tableStyles, extLst and mruColors unread", () => {
    const xfs = [`<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>`];
    const bare = readXlsxStyles(stylesXml({ xfs }), undefined);
    const full = readXlsxStyles(stylesXml({ xfs, extra: unread }).replace("<fonts", `<numFmts count="1"><numFmt numFmtId="164" formatCode="0.0%"/></numFmts><fonts`), undefined);
    expect(full).toEqual(bare);
    expect(full.styles[0]!.numberFormat).toBeUndefined();
    expect(full.styles[0]!.namedStyle).toBeUndefined();
    expect(full.indexedColors).toBeUndefined();
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
    expect(readXlsxStyles(undefined, undefined)).toEqual({ styles: [{}], defaultFont: { family: "Calibri", sizePt: 11 } });
    // A part that is not well-formed, or names no xf, reads the same.
    expect(readXlsxStyles("<styleSheet", undefined)).toEqual({ styles: [{}], defaultFont: { family: "Calibri", sizePt: 11 } });
    expect(readXlsxStyles(`<styleSheet xmlns="${MAIN}"/>`, undefined)).toEqual({ styles: [{}], defaultFont: { family: "Calibri", sizePt: 11 } });
  });

  it("reads a strict (ISO 29500) styles part", () => {
    const xml = stylesXml({ ns: STRICT, fonts: [`<font><b/><sz val="10"/><name val="Arial"/></font>`] });
    expect(readXlsxStyles(xml, undefined)).toEqual({
      styles: [{ font: { family: "Arial", sizePt: 10, bold: true }, sourceXf: 0 }],
      defaultFont: { family: "Arial", sizePt: 10 },
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
    for (const style of styles) {
      expect(style.numberFormat).toBeUndefined();
      expect(style.namedStyle).toBeUndefined();
    }
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
