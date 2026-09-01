// SPEC-094 golden (the SPEC-086 `p:timing` discipline): the canonical theme
// part this editor's splice produces for `GOLDEN_THEME`, hand-authored against
// ECMA-376 §20.1.6 — not screen-scraped. `theme.spec.ts` asserts the splice in
// both directions: writing GOLDEN_THEME into THEME_SOURCE yields THEME_GOLDEN,
// and reading THEME_GOLDEN back yields GOLDEN_THEME's scheme. The `a:fmtScheme`
// block is deliberately non-trivial so byte-verbatim survival is provable.
import type { DocumentTheme } from "../../core/theme";

/** A non-Office scheme, so no golden value can pass by coincidence. */
export const GOLDEN_THEME: DocumentTheme = {
  name: "Golden",
  colors: {
    dk1: "#101010",
    lt1: "#fefefe",
    dk2: "#123456",
    lt2: "#edcba9",
    accent1: "#aa1122",
    accent2: "#bb2233",
    accent3: "#cc3344",
    accent4: "#dd4455",
    accent5: "#ee5566",
    accent6: "#ff6677",
    hlink: "#0000ee",
    folHlink: "#551a8b",
  },
  fonts: { major: "Georgia", minor: "Verdana" },
};

const XML_DECL = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`;
const A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";

/** The untouched remainder every splice must carry through byte-verbatim. */
export const THEME_FMT_SCHEME =
  `<a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>` +
  `<a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:lumMod val="110000"/></a:schemeClr></a:gs>` +
  `<a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="78000"/></a:schemeClr></a:gs></a:gsLst>` +
  `<a:lin ang="5400000" scaled="0"/></a:gradFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>` +
  `<a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>` +
  `<a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>` +
  `<a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst>` +
  `<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle>` +
  `<a:effectStyle><a:effectLst><a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" rotWithShape="0">` +
  `<a:srgbClr val="000000"><a:alpha val="63000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle></a:effectStyleLst>` +
  `<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>` +
  `<a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme>`;

/** A source theme part in PowerPoint's shape: sysClr dk1/lt1, Office colours,
 * Calibri fonts, plus the fmtScheme and objectDefaults the splice must never
 * touch. */
export const THEME_SOURCE =
  `${XML_DECL}<a:theme xmlns:a="${A_NS}" name="Office Theme"><a:themeElements>` +
  `<a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>` +
  `<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="44546A"/></a:dk2>` +
  `<a:lt2><a:srgbClr val="E7E6E6"/></a:lt2><a:accent1><a:srgbClr val="4472C4"/></a:accent1>` +
  `<a:accent2><a:srgbClr val="ED7D31"/></a:accent2><a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>` +
  `<a:accent4><a:srgbClr val="FFC000"/></a:accent4><a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>` +
  `<a:accent6><a:srgbClr val="70AD47"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink>` +
  `<a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme>` +
  `<a:fontScheme name="Office"><a:majorFont><a:latin typeface="Calibri Light" panose="020F0302020204030204"/>` +
  `<a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="游ゴシック Light"/></a:majorFont>` +
  `<a:minorFont><a:latin typeface="Calibri" panose="020F0502020204030204"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme>` +
  THEME_FMT_SCHEME +
  `</a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>`;

/** THEME_SOURCE after splicing GOLDEN_THEME in: the clrScheme is the canonical
 * serialisation (dk1/lt1 keep their sysClr form with `lastClr` updated), the
 * two latin typefaces move, and every other byte survives verbatim. */
export const THEME_GOLDEN =
  `${XML_DECL}<a:theme xmlns:a="${A_NS}" name="Office Theme"><a:themeElements>` +
  `<a:clrScheme name="Golden"><a:dk1><a:sysClr val="windowText" lastClr="101010"/></a:dk1>` +
  `<a:lt1><a:sysClr val="window" lastClr="FEFEFE"/></a:lt1><a:dk2><a:srgbClr val="123456"/></a:dk2>` +
  `<a:lt2><a:srgbClr val="EDCBA9"/></a:lt2><a:accent1><a:srgbClr val="AA1122"/></a:accent1>` +
  `<a:accent2><a:srgbClr val="BB2233"/></a:accent2><a:accent3><a:srgbClr val="CC3344"/></a:accent3>` +
  `<a:accent4><a:srgbClr val="DD4455"/></a:accent4><a:accent5><a:srgbClr val="EE5566"/></a:accent5>` +
  `<a:accent6><a:srgbClr val="FF6677"/></a:accent6><a:hlink><a:srgbClr val="0000EE"/></a:hlink>` +
  `<a:folHlink><a:srgbClr val="551A8B"/></a:folHlink></a:clrScheme>` +
  `<a:fontScheme name="Office"><a:majorFont><a:latin typeface="Georgia" panose="020F0302020204030204"/>` +
  `<a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="游ゴシック Light"/></a:majorFont>` +
  `<a:minorFont><a:latin typeface="Verdana" panose="020F0502020204030204"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme>` +
  THEME_FMT_SCHEME +
  `</a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>`;
