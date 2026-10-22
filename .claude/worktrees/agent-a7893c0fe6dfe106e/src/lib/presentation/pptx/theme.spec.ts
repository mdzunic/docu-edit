// SPEC-094: the pptx theme wire mapping — golden equality both ways, the
// clrMap reverse mapping (including the `p:clrMapOvr` opt-out), and the splice
// leaving `a:fmtScheme` byte-identical.
import { describe, expect, it } from "vitest";
import { parseXml } from "../../core/xml";
import { DEFAULT_THEME } from "../../core/theme";
import {
  colorRefFrom,
  fontRefFrom,
  readDeckTheme,
  schemeClrXml,
  slideClrMapMode,
  slideSpaceName,
  spliceThemeXml,
} from "./theme";
import { GOLDEN_THEME, THEME_FMT_SCHEME, THEME_GOLDEN, THEME_SOURCE } from "./themeGolden";

const A = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';
const DEFAULT_MAP: Record<string, string> = {
  bg1: "lt1",
  tx1: "dk1",
  bg2: "lt2",
  tx2: "dk2",
  accent1: "accent1",
  accent2: "accent2",
  accent3: "accent3",
  accent4: "accent4",
  accent5: "accent5",
  accent6: "accent6",
  hlink: "hlink",
  folHlink: "folHlink",
};

function el(inner: string): Element {
  return parseXml(`<a:wrap ${A}>${inner}</a:wrap>`).documentElement.firstElementChild as Element;
}

describe("spliceThemeXml (golden, both directions)", () => {
  it("writes GOLDEN_THEME into the source part exactly as THEME_GOLDEN", () => {
    expect(spliceThemeXml(THEME_SOURCE, GOLDEN_THEME)).toBe(THEME_GOLDEN);
  });

  it("reads THEME_GOLDEN back as GOLDEN_THEME's scheme", () => {
    const root = parseXml(THEME_GOLDEN).documentElement;
    const theme = readDeckTheme(root, undefined, "ppt/theme/theme1.xml", DEFAULT_THEME);
    expect(theme?.name).toBe("Golden");
    expect(theme?.colors).toEqual(GOLDEN_THEME.colors);
    expect(theme?.fonts).toEqual(GOLDEN_THEME.fonts);
    expect(theme?.partName).toBe("ppt/theme/theme1.xml");
  });

  it("keeps a:fmtScheme (and objectDefaults/extraClrSchemeLst) byte-identical", () => {
    const out = spliceThemeXml(THEME_SOURCE, GOLDEN_THEME);
    expect(out).toContain(THEME_FMT_SCHEME);
    expect(out).toContain("<a:objectDefaults/><a:extraClrSchemeLst/>");
  });

  it("keeps dk1/lt1's a:sysClr form with lastClr updated; own slots write a:srgbClr", () => {
    const out = spliceThemeXml(THEME_SOURCE, GOLDEN_THEME);
    expect(out).toContain('<a:dk1><a:sysClr val="windowText" lastClr="101010"/></a:dk1>');
    expect(out).toContain('<a:accent1><a:srgbClr val="AA1122"/></a:accent1>');
  });

  it("splicing the same theme twice is idempotent", () => {
    const once = spliceThemeXml(THEME_SOURCE, GOLDEN_THEME);
    expect(spliceThemeXml(once, GOLDEN_THEME)).toBe(once);
  });
});

describe("colorRefFrom (AC-14/15)", () => {
  it("captures a schemeClr with canonical-order transforms as 0..1 fractions", () => {
    const ref = colorRefFrom(el('<a:schemeClr val="tx1"><a:lumMod val="60000"/><a:lumOff val="40000"/></a:schemeClr>'), DEFAULT_MAP);
    expect(ref).toEqual({ theme: "dk1", lumMod: 0.6, lumOff: 0.4 });
  });

  it("maps the slide-space name through the clrMap", () => {
    expect(colorRefFrom(el('<a:schemeClr val="bg2"/>'), DEFAULT_MAP)).toEqual({ theme: "lt2" });
    expect(colorRefFrom(el('<a:schemeClr val="accent3"/>'), DEFAULT_MAP)).toEqual({ theme: "accent3" });
  });

  it("stays a literal for any other transform (alpha, satMod, …)", () => {
    expect(colorRefFrom(el('<a:schemeClr val="accent1"><a:alpha val="50000"/></a:schemeClr>'), DEFAULT_MAP)).toBeUndefined();
    expect(
      colorRefFrom(el('<a:schemeClr val="accent1"><a:lumMod val="60000"/><a:satMod val="30000"/></a:schemeClr>'), DEFAULT_MAP),
    ).toBeUndefined();
  });

  it("stays a literal for an unknown slot or the phClr placeholder", () => {
    expect(colorRefFrom(el('<a:schemeClr val="phClr"/>'), DEFAULT_MAP)).toBeUndefined();
  });
});

describe("fontRefFrom (AC-16)", () => {
  it("maps the placeholder typefaces and nothing else", () => {
    expect(fontRefFrom("+mj-lt")).toEqual({ theme: "major" });
    expect(fontRefFrom("+mn-lt")).toEqual({ theme: "minor" });
    expect(fontRefFrom("Calibri")).toBeUndefined();
  });
});

describe("the clrMap reverse mapping (AC-18)", () => {
  it("writes the attribute whose value is the slot", () => {
    expect(slideSpaceName("dk1", DEFAULT_MAP)).toBe("tx1");
    expect(slideSpaceName("lt1", DEFAULT_MAP)).toBe("bg1");
    expect(slideSpaceName("accent1", DEFAULT_MAP)).toBe("accent1");
  });

  it("falls back to the slot name when nothing maps to it", () => {
    expect(slideSpaceName("accent4", { bg1: "lt1" })).toBe("accent4");
    expect(slideSpaceName("dk1", undefined)).toBe("dk1");
  });
});

describe("schemeClrXml (AC-17)", () => {
  it("emits transforms as val/100000 children in canonical order", () => {
    expect(schemeClrXml({ theme: "accent1", lumMod: 0.6, lumOff: 0.4 }, DEFAULT_MAP)).toBe(
      '<a:schemeClr val="accent1"><a:lumMod val="60000"/><a:lumOff val="40000"/></a:schemeClr>',
    );
    expect(schemeClrXml({ theme: "dk1" }, DEFAULT_MAP)).toBe('<a:schemeClr val="tx1"/>');
    expect(schemeClrXml({ theme: "lt1", shade: 0.5, tint: 0.25 }, DEFAULT_MAP)).toBe(
      '<a:schemeClr val="bg1"><a:shade val="50000"/><a:tint val="25000"/></a:schemeClr>',
    );
  });
});

describe("slideClrMapMode — the p:clrMapOvr opt-out (AC-22/23)", () => {
  it("uses the deck map when the slide has no override or a masterClrMapping", () => {
    expect(slideClrMapMode(undefined, DEFAULT_MAP)).toBe(DEFAULT_MAP);
    expect(slideClrMapMode("<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>", DEFAULT_MAP)).toBe(DEFAULT_MAP);
  });

  it("uses a parseable override's own mapping", () => {
    const mode = slideClrMapMode(
      '<p:clrMapOvr><a:overrideClrMapping bg1="dk1" tx1="lt1" accent1="accent2"/></p:clrMapOvr>',
      DEFAULT_MAP,
    );
    expect(mode).toEqual({ bg1: "dk1", tx1: "lt1", accent1: "accent2" });
  });

  it("opts the slide out (literals) when the override cannot be parsed", () => {
    expect(slideClrMapMode("<p:clrMapOvr><p:garbage/></p:clrMapOvr>", DEFAULT_MAP)).toBe("literal");
    expect(slideClrMapMode("<p:clrMapOvr><a:overrideClrMapping/></p:clrMapOvr>", DEFAULT_MAP)).toBe("literal");
  });
});
