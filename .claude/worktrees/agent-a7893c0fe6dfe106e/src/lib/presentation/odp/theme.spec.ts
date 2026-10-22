// SPEC-094: the odp theme wire mapping — golden equality both ways, the
// splice leaving the rest of styles.xml byte-identical, and the parent-slot
// capture the reader's ref mapping rides on.
import { describe, expect, it } from "vitest";
import { DEFAULT_THEME } from "../../core/theme";
import { odpThemeBlockXml, readOdpTheme, spliceOdpThemeStyles, themeParentSlot } from "./theme";
import { GOLDEN_THEME, ODP_THEME_GOLDEN } from "./themeGolden";

const STYLES_PRELUDE =
  `<?xml version="1.0" encoding="UTF-8"?>` +
  `<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"` +
  ` xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"` +
  ` xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"` +
  ` xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" office:version="1.2">`;

const SOURCE_STYLES =
  `${STYLES_PRELUDE}<office:styles>` +
  `<style:style style:name="standard" style:family="graphic"><style:graphic-properties draw:fill="solid" draw:fill-color="#ffffff"/></style:style>` +
  `</office:styles>` +
  `<office:automatic-styles><style:page-layout style:name="PM1"/></office:automatic-styles>` +
  `<office:master-styles><style:master-page style:name="Default" style:page-layout-name="PM1"/></office:master-styles>` +
  `</office:document-styles>`;

describe("the de-theme-* block (golden, both directions)", () => {
  it("writes GOLDEN_THEME exactly as ODP_THEME_GOLDEN", () => {
    expect(odpThemeBlockXml(GOLDEN_THEME)).toBe(ODP_THEME_GOLDEN);
  });

  it("reads the golden block back as GOLDEN_THEME", () => {
    const styles = spliceOdpThemeStyles(SOURCE_STYLES, GOLDEN_THEME);
    const theme = readOdpTheme(styles, DEFAULT_THEME);
    expect(theme?.name).toBe("Golden");
    expect(theme?.colors).toEqual(GOLDEN_THEME.colors);
    expect(theme?.fonts).toEqual(GOLDEN_THEME.fonts);
    expect(theme?.partName).toBe("styles.xml");
  });
});

describe("spliceOdpThemeStyles", () => {
  it("splices into office:styles leaving every other byte identical", () => {
    const out = spliceOdpThemeStyles(SOURCE_STYLES, GOLDEN_THEME);
    expect(out).toBe(SOURCE_STYLES.replace("</office:styles>", `${ODP_THEME_GOLDEN}</office:styles>`));
  });

  it("replaces an existing de-theme block rather than duplicating it", () => {
    const once = spliceOdpThemeStyles(SOURCE_STYLES, DEFAULT_THEME);
    const twice = spliceOdpThemeStyles(once, GOLDEN_THEME);
    expect(twice).toBe(spliceOdpThemeStyles(SOURCE_STYLES, GOLDEN_THEME));
    expect(twice.match(/de-theme-accent1"/g)).toHaveLength(1);
  });

  it("returns the input unchanged when there is no office:styles", () => {
    const bare = `${STYLES_PRELUDE}</office:document-styles>`;
    expect(spliceOdpThemeStyles(bare, GOLDEN_THEME)).toBe(bare);
  });
});

describe("readOdpTheme", () => {
  it("returns undefined for a deck with no theme block (colours stay literal)", () => {
    expect(readOdpTheme(SOURCE_STYLES, DEFAULT_THEME)).toBeUndefined();
  });

  it("ignores an unrecognised extension theme block (preserved, never parsed)", () => {
    // A LibreOffice-style extension block whose shape this spec does not claim
    // to know: it must neither produce a theme nor break the read.
    const withExtension = SOURCE_STYLES.replace(
      "</office:styles>",
      `<style:style style:name="vendor-theme" style:family="graphic"/></office:styles>`,
    );
    expect(readOdpTheme(withExtension, DEFAULT_THEME)).toBeUndefined();
  });
});

describe("themeParentSlot (the reader's ref capture)", () => {
  it("maps both the graphic and the text parent forms to their slot", () => {
    expect(themeParentSlot("de-theme-accent1")).toBe("accent1");
    expect(themeParentSlot("de-theme-text-dk2")).toBe("dk2");
  });

  it("rejects non-theme parents and unknown slots", () => {
    expect(themeParentSlot("standard")).toBeUndefined();
    expect(themeParentSlot("de-theme-major")).toBeUndefined();
    expect(themeParentSlot("de-theme-nosuch")).toBeUndefined();
    expect(themeParentSlot(undefined)).toBeUndefined();
  });
});
