// Presentation fixture builders: a small, structurally real `.pptx` and
// `.odp` produced from strings + core/zipWriter alone — deliberately
// DOM-free so e2e specs can call them from Node to generate the checked-in
// e2e/fixtures files (the same constraint as the from-scratch docx writer).
// Both decks carry the same three slides so the reader specs can assert one
// expected outline against either format:
//   1. "Quarterly Review" + subtitle line
//   2. "Roadmap" + three body lines (the pptx variant folds the last two
//      into one a:p split by a:br)
//   3. no title; a plain text box plus, in the pptx, a nested group shape
//      (proves the reader walks p:grpSp)
import { zip } from "../../core/zipWriter";

const P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main";
const A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

const encoder = new TextEncoder();

function pptxSp(id: number, name: string, phType: string | undefined, paragraphs: string[]): string {
  const ph = phType ? `<p:ph type="${phType}"${phType === "subTitle" ? ' idx="1"' : ""}/>` : "";
  const body = paragraphs.map((p) => `<a:p>${p}</a:p>`).join("");
  return (
    `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr/><p:nvPr>${ph}</p:nvPr></p:nvSpPr>` +
    `<p:spPr/><p:txBody><a:bodyPr/>${body}</p:txBody></p:sp>`
  );
}

function pptxSlideXml(shapes: string): string {
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<p:sld xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}"><p:cSld><p:spTree>` +
    `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>` +
    shapes +
    `</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`
  );
}

const SLIDE_RELS =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
  `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
  `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>` +
  `</Relationships>`;

const PPTX_THEME =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
  `<a:theme xmlns:a="${A_NS}" name="Fixture"><a:themeElements>` +
  `<a:clrScheme name="Fixture"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>` +
  `<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="44546A"/></a:dk2>` +
  `<a:lt2><a:srgbClr val="E7E6E6"/></a:lt2><a:accent1><a:srgbClr val="4472C4"/></a:accent1>` +
  `<a:accent2><a:srgbClr val="ED7D31"/></a:accent2><a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>` +
  `<a:accent4><a:srgbClr val="FFC000"/></a:accent4><a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>` +
  `<a:accent6><a:srgbClr val="70AD47"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink>` +
  `<a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme>` +
  `<a:fontScheme name="Fixture"><a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>` +
  `<a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme>` +
  `<a:fmtScheme name="Fixture"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>` +
  `<a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>` +
  `<a:lnStyleLst><a:ln><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>` +
  `<a:ln><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst>` +
  `<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle>` +
  `<a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>` +
  `<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>` +
  `<a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme>` +
  `</a:themeElements></a:theme>`;

/** The pptx fixture's expected preview outline — what the reader specs (and
 * the e2e assertions) compare against. */
export const FIXTURE_SLIDES = [
  { title: "Quarterly Review", lines: ["FY26 results and outlook"] },
  { title: "Roadmap", lines: ["Ship presentations MVP", "Broaden format support", "Hire two engineers", "Open Berlin office"] },
  { lines: ["Thank you", "questions@example.com", "Grouped note"] },
] as const;

export const FIXTURE_DECK_TITLE = "Deck fixture";
export const FIXTURE_DECK_AUTHOR = "Fixture Author";

/** A minimal, structurally real `.pptx`: content types, root rels, the
 * presentation part with a 3-slide sldIdLst, per-slide parts + rels, one
 * master/layout/theme chain, and docProps. 15 parts. */
export async function buildPptxFixtureBytes(): Promise<Uint8Array> {
  const slide1 = pptxSlideXml(
    pptxSp(2, "Title 1", "ctrTitle", ["<a:r><a:t>Quarterly Review</a:t></a:r>"]) +
      pptxSp(3, "Subtitle 2", "subTitle", ["<a:r><a:t>FY26 results and outlook</a:t></a:r>"]),
  );
  const slide2 = pptxSlideXml(
    pptxSp(2, "Title 1", "title", ["<a:r><a:t>Roadmap</a:t></a:r>"]) +
      pptxSp(3, "Content 2", "body", [
        "<a:r><a:t>Ship presentations MVP</a:t></a:r>",
        "<a:r><a:t>Broaden format support</a:t></a:r>",
        "<a:r><a:t>Hire two engineers</a:t></a:r><a:br/><a:r><a:t>Open Berlin office</a:t></a:r>",
      ]),
  );
  const slide3 = pptxSlideXml(
    pptxSp(2, "TextBox 1", undefined, [
      "<a:r><a:t>Thank you</a:t></a:r>",
      "<a:r><a:t>questions@example.com</a:t></a:r>",
    ]) +
      `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="4" name="Group 3"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>` +
      pptxSp(5, "Grouped 4", undefined, ["<a:r><a:t>Grouped note</a:t></a:r>"]) +
      `</p:grpSp>`,
  );

  const contentTypes =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
    `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
    `<Default Extension="xml" ContentType="application/xml"/>` +
    `<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>` +
    `<Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>` +
    `<Override PartName="/ppt/slides/slide2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>` +
    `<Override PartName="/ppt/slides/slide3.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>` +
    `<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>` +
    `<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>` +
    `<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>` +
    `<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>` +
    `<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>` +
    `</Types>`;

  const rootRels =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>` +
    `<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>` +
    `<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>` +
    `</Relationships>`;

  const presentation =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<p:presentation xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}">` +
    `<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>` +
    `<p:sldIdLst><p:sldId id="256" r:id="rId2"/><p:sldId id="257" r:id="rId3"/><p:sldId id="258" r:id="rId4"/></p:sldIdLst>` +
    `<p:sldSz cx="12192000" cy="6858000"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>`;

  const presentationRels =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>` +
    `<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>` +
    `<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide2.xml"/>` +
    `<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide3.xml"/>` +
    `<Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>` +
    `</Relationships>`;

  const slideMaster =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<p:sldMaster xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}"><p:cSld><p:spTree>` +
    `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld>` +
    `<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>` +
    `<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst></p:sldMaster>`;

  const slideMasterRels =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>` +
    `<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>` +
    `</Relationships>`;

  const slideLayout =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<p:sldLayout xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}" type="title"><p:cSld><p:spTree>` +
    `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld>` +
    `<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`;

  const slideLayoutRels =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>` +
    `</Relationships>`;

  const core =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"` +
    ` xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/"` +
    ` xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">` +
    `<dc:title>${FIXTURE_DECK_TITLE}</dc:title><dc:creator>${FIXTURE_DECK_AUTHOR}</dc:creator>` +
    `<dcterms:created xsi:type="dcterms:W3CDTF">2026-01-05T10:00:00Z</dcterms:created>` +
    `<dcterms:modified xsi:type="dcterms:W3CDTF">2026-02-10T09:30:00Z</dcterms:modified>` +
    `</cp:coreProperties>`;

  const app =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"` +
    ` xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">` +
    `<Application>docu-edit fixture</Application><Slides>3</Slides></Properties>`;

  const entry = (name: string, text: string) => ({ name, data: encoder.encode(text) });
  return zip([
    entry("[Content_Types].xml", contentTypes),
    entry("_rels/.rels", rootRels),
    entry("ppt/presentation.xml", presentation),
    entry("ppt/_rels/presentation.xml.rels", presentationRels),
    entry("ppt/slides/slide1.xml", slide1),
    entry("ppt/slides/_rels/slide1.xml.rels", SLIDE_RELS),
    entry("ppt/slides/slide2.xml", slide2),
    entry("ppt/slides/_rels/slide2.xml.rels", SLIDE_RELS),
    entry("ppt/slides/slide3.xml", slide3),
    entry("ppt/slides/_rels/slide3.xml.rels", SLIDE_RELS),
    entry("ppt/slideMasters/slideMaster1.xml", slideMaster),
    entry("ppt/slideMasters/_rels/slideMaster1.xml.rels", slideMasterRels),
    entry("ppt/slideLayouts/slideLayout1.xml", slideLayout),
    entry("ppt/slideLayouts/_rels/slideLayout1.xml.rels", slideLayoutRels),
    entry("ppt/theme/theme1.xml", PPTX_THEME),
    entry("docProps/core.xml", core),
    entry("docProps/app.xml", app),
  ]);
}

const ODF_OFFICE_NS = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const ODF_TEXT_NS = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const ODF_DRAW_NS = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
const ODF_PRESENTATION_NS = "urn:oasis:names:tc:opendocument:xmlns:presentation:1.0";

function odpFrame(cls: string | undefined, paragraphs: string[]): string {
  const classAttr = cls ? ` presentation:class="${cls}"` : "";
  const text = paragraphs.map((p) => `<text:p>${p}</text:p>`).join("");
  return `<draw:frame${classAttr}><draw:text-box>${text}</draw:text-box></draw:frame>`;
}

/** The `.odp` twin of the pptx fixture: same three slides (the a:br split
 * becomes two text:p paragraphs; the grouped shape a draw:g). 5 parts. */
export async function buildOdpFixtureBytes(): Promise<Uint8Array> {
  const page1 =
    `<draw:page draw:name="page1">` +
    odpFrame("title", ["Quarterly Review"]) +
    odpFrame("subtitle", ["FY26 results and outlook"]) +
    `</draw:page>`;
  const page2 =
    `<draw:page draw:name="page2">` +
    odpFrame("title", ["Roadmap"]) +
    odpFrame("outline", ["Ship presentations MVP", "Broaden format support", "Hire two engineers", "Open Berlin office"]) +
    `</draw:page>`;
  const page3 =
    `<draw:page draw:name="page3">` +
    odpFrame(undefined, ["Thank you", "questions@example.com"]) +
    `<draw:g>${odpFrame(undefined, ["Grouped note"])}</draw:g>` +
    `</draw:page>`;

  const content =
    `<?xml version="1.0" encoding="UTF-8"?>` +
    `<office:document-content xmlns:office="${ODF_OFFICE_NS}" xmlns:text="${ODF_TEXT_NS}"` +
    ` xmlns:draw="${ODF_DRAW_NS}" xmlns:presentation="${ODF_PRESENTATION_NS}" office:version="1.2">` +
    `<office:body><office:presentation>${page1}${page2}${page3}</office:presentation></office:body>` +
    `</office:document-content>`;

  const styles =
    `<?xml version="1.0" encoding="UTF-8"?>` +
    `<office:document-styles xmlns:office="${ODF_OFFICE_NS}" office:version="1.2"><office:styles/></office:document-styles>`;

  const meta =
    `<?xml version="1.0" encoding="UTF-8"?>` +
    `<office:document-meta xmlns:office="${ODF_OFFICE_NS}"` +
    ` xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" xmlns:dc="http://purl.org/dc/elements/1.1/" office:version="1.2">` +
    `<office:meta><dc:title>${FIXTURE_DECK_TITLE}</dc:title><meta:initial-creator>${FIXTURE_DECK_AUTHOR}</meta:initial-creator>` +
    `<meta:creation-date>2026-01-05T10:00:00Z</meta:creation-date><dc:date>2026-02-10T09:30:00Z</dc:date>` +
    `</office:meta></office:document-meta>`;

  const manifest =
    `<?xml version="1.0" encoding="UTF-8"?>` +
    `<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.2">` +
    `<manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.presentation"/>` +
    `<manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>` +
    `<manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/>` +
    `<manifest:file-entry manifest:full-path="meta.xml" manifest:media-type="text/xml"/>` +
    `</manifest:manifest>`;

  return zip([
    // ODF requires the mimetype entry first and STOREd so format sniffers
    // can read it from the raw bytes (same rule as odf/writer.ts).
    { name: "mimetype", data: encoder.encode("application/vnd.oasis.opendocument.presentation"), store: true },
    { name: "content.xml", data: encoder.encode(content) },
    { name: "styles.xml", data: encoder.encode(styles) },
    { name: "meta.xml", data: encoder.encode(meta) },
    { name: "META-INF/manifest.xml", data: encoder.encode(manifest) },
  ]);
}
