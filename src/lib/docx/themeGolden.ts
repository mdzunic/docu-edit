// SPEC-077 golden (the SPEC-094 themeGolden discipline): a hand-authored
// word/theme/theme1.xml source, the exact bytes this product's splice must
// produce for DOCX_GOLDEN_THEME, and the template a product-created theme
// part is built from (AC-36). Authored against ECMA-376 §20.1.6, not
// screen-scraped; theme.spec.ts asserts the splice in both directions. The
// `a:fmtScheme` block is deliberately non-trivial so byte-verbatim survival
// (AC-34) is provable.
import type { DocumentTheme } from "../core/theme";

/** A non-Office scheme, so no golden value can pass by coincidence. dk1
 * deliberately differs from the source's windowText black: AC-35 requires a
 * CHANGED slot to write `a:srgbClr` even where the source used `a:sysClr`,
 * while the unchanged lt1 keeps its sysClr form byte-verbatim. */
export const DOCX_GOLDEN_THEME: DocumentTheme = {
  name: "Golden",
  colors: {
    dk1: "#101010",
    lt1: "#ffffff",
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
export const DOCX_THEME_FMT_SCHEME =
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

/** A source theme part in Word's shape: sysClr dk1/lt1, Office colours,
 * Calibri fonts with panose attributes and a Jpan script font, plus the
 * fmtScheme and objectDefaults the splice must never touch. Also the
 * template a product-created word/theme/theme1.xml starts from (AC-36) —
 * its scheme matches DEFAULT_THEME, so splicing any theme in yields a
 * complete, valid part. */
export const DOCX_THEME_SOURCE =
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
  DOCX_THEME_FMT_SCHEME +
  `</a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>`;

/** DOCX_THEME_SOURCE after splicing DOCX_GOLDEN_THEME in: the changed dk1
 * becomes `a:srgbClr` (AC-35), the unchanged lt1 keeps its `a:sysClr` form
 * byte-verbatim, the two latin typefaces move with their panose attributes
 * preserved, the root name attribute carries the edited theme's name (it is
 * where AC-25 reads the name back from), and every byte outside
 * `a:clrScheme`/the typefaces/that one attribute survives verbatim (AC-34). */
export const DOCX_THEME_GOLDEN =
  `${XML_DECL}<a:theme xmlns:a="${A_NS}" name="Golden"><a:themeElements>` +
  `<a:clrScheme name="Golden"><a:dk1><a:srgbClr val="101010"/></a:dk1>` +
  `<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="123456"/></a:dk2>` +
  `<a:lt2><a:srgbClr val="EDCBA9"/></a:lt2><a:accent1><a:srgbClr val="AA1122"/></a:accent1>` +
  `<a:accent2><a:srgbClr val="BB2233"/></a:accent2><a:accent3><a:srgbClr val="CC3344"/></a:accent3>` +
  `<a:accent4><a:srgbClr val="DD4455"/></a:accent4><a:accent5><a:srgbClr val="EE5566"/></a:accent5>` +
  `<a:accent6><a:srgbClr val="FF6677"/></a:accent6><a:hlink><a:srgbClr val="0000EE"/></a:hlink>` +
  `<a:folHlink><a:srgbClr val="551A8B"/></a:folHlink></a:clrScheme>` +
  `<a:fontScheme name="Office"><a:majorFont><a:latin typeface="Georgia" panose="020F0302020204030204"/>` +
  `<a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="游ゴシック Light"/></a:majorFont>` +
  `<a:minorFont><a:latin typeface="Verdana" panose="020F0502020204030204"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme>` +
  DOCX_THEME_FMT_SCHEME +
  `</a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>`;
