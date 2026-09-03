// SPEC-077 — the .docx theme part's write lifecycle (AC-33..36): an
// untouched part survives a save byte-identical as a preserved opaque part;
// an edited theme is spliced back byte-conservatively; a package (or fresh
// build) that never had a theme part gains word/theme/theme1.xml — built
// from the golden template — with its content-type override and main-part
// relationship declared, and only when the theme was actually touched.
import { describe, it, expect } from "vitest";
import { parseDocx, type ParsedDocxDocument } from "./reader";
import { writeDocx } from "./writer";
import { validateDocx } from "./validate";
import { zip } from "../core/zipWriter";
import { unzip } from "../core/zip";
import { equalBytes } from "../core/equality";
import { readDocxTheme, spliceDocxThemeXml, THEME_CONTENT_TYPE, THEME_RELATIONSHIP } from "./theme";
import { DOCX_GOLDEN_THEME, DOCX_THEME_FMT_SCHEME, DOCX_THEME_GOLDEN, DOCX_THEME_SOURCE } from "./themeGolden";
import { createDocument, type DocumentModel } from "../documentModel";
import { WORD_NAMESPACE } from "./xml";

const encoder = new TextEncoder();
const decoder = new TextDecoder();

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types";
const REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
const OFFICE_DOCUMENT_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
const CORE_PROPERTIES_REL = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
const STYLES_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";

/** DOCX_THEME_SOURCE with distinctive whitespace and a producer comment —
 * bytes no regeneration would reproduce, so byte-equality after a save
 * proves the part was copied, not rebuilt (AC-33). */
const DISTINCTIVE_THEME_XML = DOCX_THEME_SOURCE.replace(
  "<a:themeElements>",
  "<!-- producer fingerprint, kept verbatim -->\n  <a:themeElements>\n ",
).replace("</a:theme>", "\n</a:theme>\n");

/** A minimal, validator-clean in-memory .docx; `themeXml` adds a
 * word/theme/theme1.xml part with its override and relationship. */
function buildPackage(themeXml?: string): Promise<Uint8Array> {
  const themeOverride =
    themeXml !== undefined
      ? `<Override PartName="/word/theme/theme1.xml" ContentType="${THEME_CONTENT_TYPE}"/>`
      : "";
  const themeRel =
    themeXml !== undefined
      ? `<Relationship Id="rId9" Type="${THEME_RELATIONSHIP}" Target="theme/theme1.xml"/>`
      : "";
  const contentTypes =
    `${XML_DECL}<Types xmlns="${CT_NS}">` +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
    themeOverride +
    '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' +
    "</Types>";
  const packageRels =
    `${XML_DECL}<Relationships xmlns="${REL_NS}">` +
    `<Relationship Id="rId1" Type="${OFFICE_DOCUMENT_REL}" Target="word/document.xml"/>` +
    `<Relationship Id="rId2" Type="${CORE_PROPERTIES_REL}" Target="docProps/core.xml"/>` +
    "</Relationships>";
  const documentRels =
    `${XML_DECL}<Relationships xmlns="${REL_NS}">` +
    `<Relationship Id="rId1" Type="${STYLES_REL}" Target="styles.xml"/>` +
    themeRel +
    "</Relationships>";
  const documentXml =
    `${XML_DECL}<w:document xmlns:w="${WORD_NAMESPACE}"><w:body>` +
    "<w:p><w:r><w:t>Hello theme</w:t></w:r></w:p>" +
    "</w:body></w:document>";
  const stylesXml =
    `${XML_DECL}<w:styles xmlns:w="${WORD_NAMESPACE}">` +
    '<w:style w:type="paragraph" w:styleId="Normal" w:default="1"><w:name w:val="Normal"/></w:style>' +
    "</w:styles>";
  const entries = [
    { name: "[Content_Types].xml", data: encoder.encode(contentTypes) },
    { name: "_rels/.rels", data: encoder.encode(packageRels) },
    { name: "word/document.xml", data: encoder.encode(documentXml) },
    { name: "word/_rels/document.xml.rels", data: encoder.encode(documentRels) },
    { name: "word/styles.xml", data: encoder.encode(stylesXml) },
  ];
  if (themeXml !== undefined) {
    entries.push({ name: "word/theme/theme1.xml", data: encoder.encode(themeXml) });
  }
  return zip(entries);
}

/** The model an opened document holds — retained package included, theme
 * carried across the way openDocument.ts does. */
function modelFrom(parsed: ParsedDocxDocument): DocumentModel {
  const model: DocumentModel = {
    id: "theme-doc",
    title: "Theme Doc",
    sections: parsed.sections,
    defaultStyles: parsed.defaultStyles,
    format: "docx",
    opaque: parsed.opaque,
    sourceXml: parsed.sourceXml,
  };
  if (parsed.theme !== undefined) {
    model.theme = parsed.theme;
  }
  return model;
}

describe("writeDocx — untouched theme part stays a preserved opaque part (AC-33)", () => {
  it("writes the theme part back byte-identical when the theme was not edited", async () => {
    const pkg = await buildPackage(DISTINCTIVE_THEME_XML);
    const parsed = await parseDocx(pkg);
    expect(parsed.theme).toBeDefined();
    expect(parsed.theme!.partName).toBe("word/theme/theme1.xml");
    expect(parsed.theme!.dirty).toBeUndefined();

    const { bytes } = await writeDocx(modelFrom(parsed));
    const parts = await unzip(bytes);
    expect(parts.has("word/theme/theme1.xml")).toBe(true);
    expect(equalBytes(parts.get("word/theme/theme1.xml")!, encoder.encode(DISTINCTIVE_THEME_XML))).toBe(true);
  });

  it("declares nothing twice: the existing rel and override are reused, not re-added", async () => {
    const pkg = await buildPackage(DISTINCTIVE_THEME_XML);
    const parsed = await parseDocx(pkg);
    const { bytes } = await writeDocx(modelFrom(parsed));
    const parts = await unzip(bytes);
    const rels = decoder.decode(parts.get("word/_rels/document.xml.rels")!);
    const contentTypes = decoder.decode(parts.get("[Content_Types].xml")!);
    expect(rels.split(THEME_RELATIONSHIP).length - 1).toBe(1);
    expect(contentTypes.split(THEME_CONTENT_TYPE).length - 1).toBe(1);
  });
});

describe("spliceDocxThemeXml — golden in both directions (AC-34, AC-35)", () => {
  it("splicing the golden scheme into the source yields the golden bytes exactly", () => {
    expect(spliceDocxThemeXml(DOCX_THEME_SOURCE, DOCX_GOLDEN_THEME)).toBe(DOCX_THEME_GOLDEN);
  });

  it("reading the golden bytes back yields the golden scheme", () => {
    const theme = readDocxTheme(DOCX_THEME_GOLDEN, "word/theme/theme1.xml");
    expect(theme).toBeDefined();
    expect(theme!.colors).toEqual(DOCX_GOLDEN_THEME.colors);
    expect(theme!.fonts).toEqual(DOCX_GOLDEN_THEME.fonts);
    // The part's own name attribute survives the splice verbatim (AC-34) —
    // the edited name lands on a:clrScheme — so the read-back name is the
    // preserved root's, not the scheme's.
    expect(theme!.name).toBe("Office Theme");
    expect(DOCX_THEME_GOLDEN).toContain('<a:clrScheme name="Golden">');
  });

  it("an unchanged slot keeps the source's element form; a changed one writes a:srgbClr (AC-35)", () => {
    const spliced = spliceDocxThemeXml(DOCX_THEME_SOURCE, DOCX_GOLDEN_THEME);
    // lt1 (#ffffff) matches the source's sysClr — the sysClr form survives.
    expect(spliced).toContain('<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>');
    // dk1 changed from windowText black — the slot becomes a:srgbClr.
    expect(spliced).toContain('<a:dk1><a:srgbClr val="101010"/></a:dk1>');
    expect(spliced).not.toContain('<a:sysClr val="windowText"');
  });

  it("a theme differing in every slot writes a:srgbClr for all twelve (AC-35)", () => {
    const allChanged = {
      ...DOCX_GOLDEN_THEME,
      colors: { ...DOCX_GOLDEN_THEME.colors, lt1: "#fafafa" },
    };
    const spliced = spliceDocxThemeXml(DOCX_THEME_SOURCE, allChanged);
    const clrScheme = /<a:clrScheme[\s\S]*?<\/a:clrScheme>/.exec(spliced)![0];
    expect(clrScheme).not.toContain("a:sysClr");
    expect((clrScheme.match(/<a:srgbClr /g) ?? []).length).toBe(12);
  });
});

describe("writeDocx — an edited theme is spliced into the preserved part (AC-34)", () => {
  it("replaces only clrScheme and the latin typefaces; everything else survives byte-verbatim", async () => {
    const pkg = await buildPackage(DISTINCTIVE_THEME_XML);
    const parsed = await parseDocx(pkg);
    const model = modelFrom(parsed);
    model.theme = {
      ...parsed.theme!,
      name: DOCX_GOLDEN_THEME.name,
      colors: { ...DOCX_GOLDEN_THEME.colors },
      fonts: { ...DOCX_GOLDEN_THEME.fonts },
      dirty: true,
    };

    const { bytes } = await writeDocx(model);
    const parts = await unzip(bytes);
    const themeXml = decoder.decode(parts.get("word/theme/theme1.xml")!);
    expect(themeXml).toBe(spliceDocxThemeXml(DISTINCTIVE_THEME_XML, model.theme));
    // The untouched remainder — fmtScheme (effectStyleLst/bgFillStyleLst
    // included), the producer comment and the distinctive whitespace — is
    // still byte-verbatim inside the spliced part.
    expect(themeXml).toContain(DOCX_THEME_FMT_SCHEME);
    expect(themeXml).toContain("<!-- producer fingerprint, kept verbatim -->\n  <a:themeElements>\n ");
    expect(themeXml).toContain("<a:objectDefaults/><a:extraClrSchemeLst/>");
    expect(themeXml).toContain("\n</a:theme>\n");
    // The edited scheme is in place and reads back exactly.
    const reopened = await parseDocx(bytes);
    expect(reopened.theme!.colors).toEqual(DOCX_GOLDEN_THEME.colors);
    expect(reopened.theme!.fonts).toEqual(DOCX_GOLDEN_THEME.fonts);
  });
});

describe("writeDocx — a package without a theme part (AC-36)", () => {
  it("gains word/theme/theme1.xml (golden template + splice), its override and relationship when the theme was touched", async () => {
    const pkg = await buildPackage();
    const parsed = await parseDocx(pkg);
    expect(parsed.theme).toBeUndefined();
    const model = modelFrom(parsed);
    model.theme = { ...DOCX_GOLDEN_THEME, dirty: true };

    const { bytes } = await writeDocx(model);
    const parts = await unzip(bytes);
    expect(parts.has("word/theme/theme1.xml")).toBe(true);
    expect(decoder.decode(parts.get("word/theme/theme1.xml")!)).toBe(DOCX_THEME_GOLDEN);
    const contentTypes = decoder.decode(parts.get("[Content_Types].xml")!);
    expect(contentTypes).toContain(
      `<Override PartName="/word/theme/theme1.xml" ContentType="${THEME_CONTENT_TYPE}"/>`,
    );
    const rels = decoder.decode(parts.get("word/_rels/document.xml.rels")!);
    expect(rels).toMatch(
      new RegExp(`<Relationship Id="rId\\d+" Type="${THEME_RELATIONSHIP}" Target="theme/theme1.xml"/>`),
    );
    expect(await validateDocx(bytes)).toEqual([]);
    const reopened = await parseDocx(bytes);
    expect(reopened.theme!.colors).toEqual(DOCX_GOLDEN_THEME.colors);
    expect(reopened.theme!.fonts).toEqual(DOCX_GOLDEN_THEME.fonts);
  });

  it("gains no part and declares nothing when the theme was never touched", async () => {
    const pkg = await buildPackage();
    const parsed = await parseDocx(pkg);
    const { bytes } = await writeDocx(modelFrom(parsed));
    const parts = await unzip(bytes);
    expect(parts.has("word/theme/theme1.xml")).toBe(false);
    expect(decoder.decode(parts.get("[Content_Types].xml")!)).not.toContain(THEME_CONTENT_TYPE);
    expect(decoder.decode(parts.get("word/_rels/document.xml.rels")!)).not.toContain(THEME_RELATIONSHIP);
    expect(await validateDocx(bytes)).toEqual([]);
  });
});

describe("writeDocx — fresh build (AC-36, new/cross-format documents)", () => {
  it("emits word/theme/theme1.xml with override and relationship when the model carries a theme", async () => {
    const model = createDocument("fresh-theme");
    model.theme = { ...DOCX_GOLDEN_THEME, dirty: true };

    const { bytes } = await writeDocx(model);
    const parts = await unzip(bytes);
    expect(decoder.decode(parts.get("word/theme/theme1.xml")!)).toBe(DOCX_THEME_GOLDEN);
    expect(decoder.decode(parts.get("[Content_Types].xml")!)).toContain(
      `<Override PartName="/word/theme/theme1.xml" ContentType="${THEME_CONTENT_TYPE}"/>`,
    );
    expect(decoder.decode(parts.get("word/_rels/document.xml.rels")!)).toMatch(
      new RegExp(`<Relationship Id="rId\\d+" Type="${THEME_RELATIONSHIP}" Target="theme/theme1.xml"/>`),
    );
    expect(await validateDocx(bytes)).toEqual([]);
    const reopened = await parseDocx(bytes);
    expect(reopened.theme!.colors).toEqual(DOCX_GOLDEN_THEME.colors);
    expect(reopened.theme!.fonts).toEqual(DOCX_GOLDEN_THEME.fonts);
  });

  it("emits nothing theme-shaped when the model has no theme", async () => {
    const { bytes } = await writeDocx(createDocument("fresh-plain"));
    const parts = await unzip(bytes);
    expect(parts.has("word/theme/theme1.xml")).toBe(false);
    expect(decoder.decode(parts.get("[Content_Types].xml")!)).not.toContain(THEME_CONTENT_TYPE);
    expect(decoder.decode(parts.get("word/_rels/document.xml.rels")!)).not.toContain(THEME_RELATIONSHIP);
  });
});
