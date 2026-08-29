// Presentation fixture builders: a small, structurally real `.pptx` and
// `.odp` produced from strings + core/zipWriter alone — deliberately
// DOM-free so e2e specs can call them from Node to generate the checked-in
// e2e/fixtures files (the same constraint as the from-scratch docx writer).
// Both decks carry the same FOUR slides so the reader specs can assert one
// expected outline and one cross-format frame structure against either
// format:
//   1. "Quarterly Review" + subtitle line
//   2. "Roadmap" + four body lines (the pptx variant folds the last two into
//      one a:p split by a:br)
//   3. no title; a plain text box plus a group (a nested shape / draw:g)
//   4. "Extras": three placeholders exercising all inheritance levels, an
//      explicitly-positioned themed ellipse, an image, a group of two rects
//      and a table (a foreign frame) — the geometry/inheritance/media slide.
import { zip } from "../../core/zipWriter";
import { writePptxTiming } from "../pptx/timing";
import { writeOdpAnim } from "../odp/anim";
import { GOLDEN_PAGE_LAYOUT_NAMES, LAYOUT_GOLDEN } from "../odp/layoutGolden";
import { createBlankPresentation } from "../slides";
import { createTableFrame } from "../tables";
import { writePptx } from "../pptx/writer";
import { writeOdp } from "../odp/writer";
import type { SlideAnimation } from "../model";

const P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main";
const A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

const encoder = new TextEncoder();

/** cm → EMU (914400 per inch). */
const EMU = (cm: number): number => Math.round((cm / 2.54) * 914400);
const xfrm = (xCm: number, yCm: number, wCm: number, hCm: number): string =>
  `<a:xfrm><a:off x="${EMU(xCm)}" y="${EMU(yCm)}"/><a:ext cx="${EMU(wCm)}" cy="${EMU(hCm)}"/></a:xfrm>`;

/** A 2×2 solid-colour PNG (75 bytes) — the "reuse the small PNG convention"
 * the Reference calls for; small enough to embed as a byte literal. */
const IMAGE_PNG = new Uint8Array([
  137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 68, 82, 0, 0, 0, 2, 0, 0, 0, 2, 8, 2, 0, 0, 0, 253, 212, 154,
  115, 0, 0, 0, 18, 73, 68, 65, 84, 120, 156, 99, 48, 238, 56, 99, 220, 113, 134, 1, 66, 1, 0, 40, 110, 6, 29, 254, 136,
  5, 56, 0, 0, 0, 0, 73, 69, 78, 68, 174, 66, 96, 130,
]);

function pptxSp(
  id: number,
  name: string,
  ph: { type?: string; idx?: string } | undefined,
  paragraphs: string[],
  spPrInner = "",
): string {
  const phXml = ph
    ? `<p:ph${ph.type ? ` type="${ph.type}"` : ""}${ph.idx ? ` idx="${ph.idx}"` : ""}/>`
    : "";
  const body = paragraphs.map((p) => `<a:p>${p}</a:p>`).join("");
  return (
    `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr/><p:nvPr>${phXml}</p:nvPr></p:nvSpPr>` +
    `<p:spPr>${spPrInner}</p:spPr><p:txBody><a:bodyPr/>${body}</p:txBody></p:sp>`
  );
}

/** A non-placeholder rect shape at a slide-space box. */
function pptxRect(id: number, xCm: number, yCm: number, wCm: number, hCm: number): string {
  return (
    `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Rect ${id}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
    `<p:spPr>${xfrm(xCm, yCm, wCm, hCm)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:sp>`
  );
}

function pptxSlideXml(shapes: string, layoutRel = "rId1"): string {
  void layoutRel;
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

// Slide 1 references the shared layout AND its notes-slide part (SPEC-083), so
// the reader follows the notesSlide relationship to the notes body.
const SLIDE1_RELS =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
  `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
  `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>` +
  `<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="../notesSlides/notesSlide1.xml"/>` +
  `</Relationships>`;

// A notes-slide part: a slide-image placeholder (ignored on read) and the body
// placeholder carrying slide 1's notes, one with a bold run (flattened) and a
// trailing empty paragraph (stripped) — pinning AC-2/4.
const NOTES_SLIDE1 =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
  `<p:notes xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}"><p:cSld><p:spTree>` +
  `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>` +
  `<p:sp><p:nvSpPr><p:cNvPr id="2" name="Slide Image Placeholder 1"/><p:cNvSpPr/><p:nvPr><p:ph type="sldImg"/></p:nvPr></p:nvSpPr><p:spPr/></p:sp>` +
  `<p:sp><p:nvSpPr><p:cNvPr id="3" name="Notes Placeholder 2"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>` +
  `<p:spPr/><p:txBody><a:bodyPr/>` +
  `<a:p><a:r><a:rPr lang="en-US" b="1"/><a:t>Open with the Q3 highlight</a:t></a:r></a:p>` +
  `<a:p><a:r><a:t>Pause before the revenue chart</a:t></a:r></a:p>` +
  `<a:p/>` +
  `</p:txBody></p:sp>` +
  `</p:spTree></p:cSld></p:notes>`;

const NOTES_SLIDE1_RELS =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
  `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
  `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="../slides/slide1.xml"/>` +
  `</Relationships>`;

// Slide 4 references layout2 (which carries the inheritance-source title
// placeholder) and the shared image part.
const SLIDE4_RELS =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
  `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
  `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout2.xml"/>` +
  `<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>` +
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

/** The fixture's expected outline (title + body lines), per slide — what the
 * reader specs (and the e2e assertions) compare `slideOutline` against. */
export const FIXTURE_SLIDES = [
  { title: "Quarterly Review", lines: ["FY26 results and outlook"] },
  { title: "Roadmap", lines: ["Ship presentations MVP", "Broaden format support", "Hire two engineers", "Open Berlin office"] },
  { lines: ["Thank you", "questions@example.com", "Grouped note"] },
  { title: "Extras", lines: ["Inherited body", "Own geometry"] },
] as const;

/** The cross-format frame kinds, in z-order, both readers must agree on
 * (AC-6). */
export const FIXTURE_FRAME_KINDS: string[][] = [
  ["text", "text"],
  ["text", "text"],
  ["text", "group"],
  // Slide 4's trailing graphicFrame/table:table is a modelled table (SPEC-091);
  // before SPEC-091 it read as a foreign "Table" box.
  ["text", "text", "text", "shape", "image", "group", "table"],
];

/** Slide 4's per-frame geometry in cm ([x, y, w, h]), asserted equal across
 * the two formats within 0.05cm (AC-8). Group children are not listed (the
 * top-level group box is). */
export const FIXTURE_SLIDE4_GEOMETRY: [number, number, number, number][] = [
  [2, 1, 29, 3], // title (pptx: layout-inherited; odp: inline)
  [2, 5, 14, 10], // body (pptx: master-inherited; odp: master-page-inherited)
  [18, 5, 13, 3], // own geometry
  [2, 15, 5, 3], // ellipse
  [25, 15, 4, 3], // image
  [8, 15, 5.5, 3], // group box (union of the two rects)
  [16, 15, 8, 3], // table (modelled TableFrame, SPEC-091)
];

export const FIXTURE_DECK_TITLE = "Deck fixture";
export const FIXTURE_DECK_AUTHOR = "Fixture Author";

/** Slide 1's speaker notes (SPEC-083) — identical across both fixture formats
 * so the reader specs pin extraction to exact strings. Each source carries a
 * trailing empty paragraph the readers must strip (AC-4) and inline formatting
 * the readers must flatten (AC). */
export const FIXTURE_NOTES = ["Open with the Q3 highlight", "Pause before the revenue chart"];

/** The themed ellipse's resolved fill (accent1 + lumMod 60% + lumOff 40% =
 * PowerPoint's "Lighter 40%"). */
export const FIXTURE_SHAPE_FILL = "#8faadc";

/** A minimal, structurally real `.pptx` with two layouts, a master carrying
 * txStyles and an inherited body placeholder, a theme, an image part, and
 * four slides. */
export async function buildPptxFixtureBytes(): Promise<Uint8Array> {
  // Slides 1-3 carry their placeholders'/text boxes' own a:xfrm — as real
  // decks do — so they render at legible slide-space boxes (AC-42/63/65).
  // Slide 4 alone is left to exercise the three placeholder inheritance levels
  // (AC-9/10/11/12).
  const slide1 = pptxSlideXml(
    pptxSp(2, "Title 1", { type: "ctrTitle" }, ["<a:r><a:t>Quarterly Review</a:t></a:r>"], xfrm(2, 6, 29, 3)) +
      pptxSp(3, "Subtitle 2", { type: "subTitle", idx: "1" }, ["<a:r><a:t>FY26 results and outlook</a:t></a:r>"], xfrm(2, 10, 29, 2)),
  );
  const slide2 = pptxSlideXml(
    pptxSp(2, "Title 1", { type: "title" }, ["<a:r><a:t>Roadmap</a:t></a:r>"], xfrm(2, 1, 29, 3)) +
      pptxSp(3, "Content 2", { type: "body" }, [
        "<a:r><a:t>Ship presentations MVP</a:t></a:r>",
        "<a:r><a:t>Broaden format support</a:t></a:r>",
        "<a:r><a:t>Hire two engineers</a:t></a:r><a:br/><a:r><a:t>Open Berlin office</a:t></a:r>",
      ], xfrm(2, 5, 29, 12)),
  );
  const slide3 = pptxSlideXml(
    pptxSp(2, "TextBox 1", undefined, [
      "<a:r><a:t>Thank you</a:t></a:r>",
      "<a:r><a:t>questions@example.com</a:t></a:r>",
    ], xfrm(2, 2, 24, 6)) +
      `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="4" name="Group 3"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>` +
      pptxSp(5, "Grouped 4", undefined, ["<a:r><a:t>Grouped note</a:t></a:r>"], xfrm(2, 10, 20, 3)) +
      `</p:grpSp>`,
  );

  // Slide 4: three placeholders (layout-inherited title, master-inherited
  // body, own-geometry body), a themed ellipse, an image, a group of two
  // rects and a table.
  const ellipseFill =
    `<a:solidFill><a:schemeClr val="accent1"><a:lumMod val="60000"/><a:lumOff val="40000"/></a:schemeClr></a:solidFill>`;
  const ellipse =
    `<p:sp><p:nvSpPr><p:cNvPr id="20" name="Ellipse"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
    `<p:spPr>${xfrm(2, 15, 5, 3)}<a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>${ellipseFill}</p:spPr></p:sp>`;
  const pic =
    `<p:pic><p:nvPicPr><p:cNvPr id="21" name="Photo"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>` +
    `<p:blipFill><a:blip r:embed="rId2"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
    `<p:spPr>${xfrm(25, 15, 4, 3)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`;
  const group =
    `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="22" name="Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
    `<p:grpSpPr><a:xfrm><a:off x="${EMU(8)}" y="${EMU(15)}"/><a:ext cx="${EMU(5.5)}" cy="${EMU(3)}"/>` +
    `<a:chOff x="${EMU(8)}" y="${EMU(15)}"/><a:chExt cx="${EMU(5.5)}" cy="${EMU(3)}"/></a:xfrm></p:grpSpPr>` +
    pptxRect(23, 8, 15, 2.5, 3) +
    pptxRect(24, 11, 15, 2.5, 3) +
    `</p:grpSp>`;
  const table =
    `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="25" name="Table"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>` +
    `<p:xfrm><a:off x="${EMU(16)}" y="${EMU(15)}"/><a:ext cx="${EMU(8)}" cy="${EMU(3)}"/></p:xfrm>` +
    `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">` +
    `<a:tbl><a:tblPr/><a:tblGrid><a:gridCol w="${EMU(8)}"/></a:tblGrid>` +
    `<a:tr h="${EMU(3)}"><a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:t>Cell</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc></a:tr></a:tbl>` +
    `</a:graphicData></a:graphic></p:graphicFrame>`;
  const slide4 = pptxSlideXml(
    pptxSp(2, "Title", { type: "title" }, ["<a:r><a:t>Extras</a:t></a:r>"]) +
      pptxSp(3, "Body", { type: "body", idx: "1" }, ["<a:r><a:t>Inherited body</a:t></a:r>"]) +
      pptxSp(
        4,
        "Own",
        { type: "body", idx: "2" },
        ['<a:r><a:rPr b="1" sz="2400"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:rPr><a:t>Own geometry</a:t></a:r>'],
        xfrm(18, 5, 13, 3),
      ) +
      ellipse +
      pic +
      group +
      table,
    "rId1",
  );

  const contentTypes =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
    `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
    `<Default Extension="xml" ContentType="application/xml"/>` +
    `<Default Extension="png" ContentType="image/png"/>` +
    `<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>` +
    `<Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>` +
    `<Override PartName="/ppt/slides/slide2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>` +
    `<Override PartName="/ppt/slides/slide3.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>` +
    `<Override PartName="/ppt/slides/slide4.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>` +
    `<Override PartName="/ppt/notesSlides/notesSlide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/>` +
    `<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>` +
    `<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>` +
    `<Override PartName="/ppt/slideLayouts/slideLayout2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>` +
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
    `<p:sldIdLst><p:sldId id="256" r:id="rId2"/><p:sldId id="257" r:id="rId3"/><p:sldId id="258" r:id="rId4"/><p:sldId id="259" r:id="rId6"/></p:sldIdLst>` +
    `<p:sldSz cx="12192000" cy="6858000"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>`;

  const presentationRels =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>` +
    `<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>` +
    `<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide2.xml"/>` +
    `<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide3.xml"/>` +
    `<Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>` +
    `<Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide4.xml"/>` +
    `</Relationships>`;

  const slideMaster =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<p:sldMaster xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}"><p:cSld><p:spTree>` +
    `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>` +
    // Master body placeholder (type=body idx=1) — slide 4's body inherits its
    // geometry (AC-10).
    `<p:sp><p:nvSpPr><p:cNvPr id="10" name="Body PH"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>` +
    `<p:spPr>${xfrm(2, 5, 14, 10)}</p:spPr><p:txBody><a:bodyPr/><a:p/></p:txBody></p:sp>` +
    `</p:spTree></p:cSld>` +
    `<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>` +
    `<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/><p:sldLayoutId id="2147483650" r:id="rId3"/></p:sldLayoutIdLst>` +
    `<p:txStyles>` +
    `<p:titleStyle><a:lvl1pPr><a:defRPr sz="4400"/></a:lvl1pPr></p:titleStyle>` +
    `<p:bodyStyle><a:lvl1pPr><a:defRPr sz="1800"/></a:lvl1pPr></p:bodyStyle>` +
    `<p:otherStyle><a:lvl1pPr><a:defRPr sz="1800"/></a:lvl1pPr></p:otherStyle>` +
    `</p:txStyles></p:sldMaster>`;

  const slideMasterRels =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>` +
    `<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>` +
    `<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout2.xml"/>` +
    `</Relationships>`;

  const slideLayout1 =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<p:sldLayout xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}" type="title"><p:cSld><p:spTree>` +
    `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld>` +
    `<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`;

  // Layout 2 carries the title placeholder slide 4's title inherits (AC-9).
  const slideLayout2 =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<p:sldLayout xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}" type="obj"><p:cSld><p:spTree>` +
    `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>` +
    `<p:sp><p:nvSpPr><p:cNvPr id="11" name="Title PH"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>` +
    `<p:spPr>${xfrm(2, 1, 29, 3)}</p:spPr><p:txBody><a:bodyPr/><a:p/></p:txBody></p:sp>` +
    `</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`;

  const layout1Rels =
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
    `<Application>docu-edit fixture</Application><Slides>4</Slides></Properties>`;

  const entry = (name: string, text: string) => ({ name, data: encoder.encode(text) });
  return zip([
    entry("[Content_Types].xml", contentTypes),
    entry("_rels/.rels", rootRels),
    entry("ppt/presentation.xml", presentation),
    entry("ppt/_rels/presentation.xml.rels", presentationRels),
    entry("ppt/slides/slide1.xml", slide1),
    entry("ppt/slides/_rels/slide1.xml.rels", SLIDE1_RELS),
    entry("ppt/notesSlides/notesSlide1.xml", NOTES_SLIDE1),
    entry("ppt/notesSlides/_rels/notesSlide1.xml.rels", NOTES_SLIDE1_RELS),
    entry("ppt/slides/slide2.xml", slide2),
    entry("ppt/slides/_rels/slide2.xml.rels", SLIDE_RELS),
    entry("ppt/slides/slide3.xml", slide3),
    entry("ppt/slides/_rels/slide3.xml.rels", SLIDE_RELS),
    entry("ppt/slides/slide4.xml", slide4),
    entry("ppt/slides/_rels/slide4.xml.rels", SLIDE4_RELS),
    entry("ppt/slideMasters/slideMaster1.xml", slideMaster),
    entry("ppt/slideMasters/_rels/slideMaster1.xml.rels", slideMasterRels),
    entry("ppt/slideLayouts/slideLayout1.xml", slideLayout1),
    entry("ppt/slideLayouts/_rels/slideLayout1.xml.rels", layout1Rels),
    entry("ppt/slideLayouts/slideLayout2.xml", slideLayout2),
    entry("ppt/slideLayouts/_rels/slideLayout2.xml.rels", layout1Rels),
    entry("ppt/theme/theme1.xml", PPTX_THEME),
    { name: "ppt/media/image1.png", data: IMAGE_PNG },
    entry("docProps/core.xml", core),
    entry("docProps/app.xml", app),
  ]);
}

// ============================================================================
// ODP twin
// ============================================================================

const ODF_OFFICE_NS = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const ODF_STYLE_NS = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const ODF_TEXT_NS = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const ODF_DRAW_NS = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
const ODF_FO_NS = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0";
const ODF_SVG_NS = "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0";
const ODF_TABLE_NS = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
const ODF_XLINK_NS = "http://www.w3.org/1999/xlink";
const ODF_PRESENTATION_NS = "urn:oasis:names:tc:opendocument:xmlns:presentation:1.0";

const geom = (xCm: number, yCm: number, wCm: number, hCm: number): string =>
  ` svg:x="${xCm}cm" svg:y="${yCm}cm" svg:width="${wCm}cm" svg:height="${hCm}cm"`;

function odpFrame(cls: string | undefined, paragraphs: string[], geometry = ""): string {
  const classAttr = cls ? ` presentation:class="${cls}"` : "";
  const text = paragraphs.map((p) => `<text:p>${p}</text:p>`).join("");
  return `<draw:frame${classAttr}${geometry}><draw:text-box>${text}</draw:text-box></draw:frame>`;
}

/** The `.odp` twin: same four slides (the a:br split becomes two text:p; the
 * grouped shape a draw:g), with styles.xml carrying the master page and the
 * page size. */
export async function buildOdpFixtureBytes(): Promise<Uint8Array> {
  // The twin's slides 1-3 carry the same slide-space geometry as the .pptx
  // twin so both formats render legibly (AC-42/63/65); slide 4 alone leaves a
  // placeholder geometry-less to exercise the master-page fallback (AC-21).
  // Slide 1 carries speaker notes (SPEC-083): a page-thumbnail (ignored on
  // read), the notes frame with a styled span (flattened) and a trailing empty
  // paragraph (stripped) — pinning AC-3/4.
  const page1Notes =
    `<presentation:notes>` +
    `<draw:page-thumbnail${geom(1, 1, 10, 7.5)}/>` +
    `<draw:frame presentation:class="notes"${geom(1.5, 13, 16, 9)}><draw:text-box>` +
    `<text:p>Open with the Q3 highlight</text:p>` +
    `<text:p><text:span text:style-name="T1">Pause before the revenue chart</text:span></text:p>` +
    `<text:p/>` +
    `</draw:text-box></draw:frame>` +
    `</presentation:notes>`;
  const page1 =
    `<draw:page draw:name="page1">` +
    odpFrame("title", ["Quarterly Review"], geom(2, 6, 29, 3)) +
    odpFrame("subtitle", ["FY26 results and outlook"], geom(2, 10, 29, 2)) +
    page1Notes +
    `</draw:page>`;
  const page2 =
    `<draw:page draw:name="page2">` +
    odpFrame("title", ["Roadmap"], geom(2, 1, 29, 3)) +
    odpFrame("outline", ["Ship presentations MVP", "Broaden format support", "Hire two engineers", "Open Berlin office"], geom(2, 5, 29, 12)) +
    `</draw:page>`;
  const page3 =
    `<draw:page draw:name="page3">` +
    odpFrame(undefined, ["Thank you", "questions@example.com"], geom(2, 2, 24, 6)) +
    `<draw:g>${odpFrame(undefined, ["Grouped note"], geom(2, 10, 20, 3))}</draw:g>` +
    `</draw:page>`;

  // Slide 4: title (inline geometry), body (no geometry → master-page
  // fallback, AC-21, with a parent-style-chained styled run, AC-20), own
  // geometry, a themed ellipse, an image, a group of two rects and a table.
  const page4 =
    `<draw:page draw:name="page4">` +
    odpFrame("title", ["Extras"], geom(2, 1, 29, 3)) +
    // No geometry: the "body"-class placeholder inherits the master page's.
    `<draw:frame presentation:class="body"><draw:text-box><text:p><text:span text:style-name="T1">Inherited body</text:span></text:p></draw:text-box></draw:frame>` +
    odpFrame("outline", ["Own geometry"], geom(18, 5, 13, 3)) +
    `<draw:ellipse draw:style-name="gr1"${geom(2, 15, 5, 3)}/>` +
    `<draw:frame${geom(25, 15, 4, 3)}><draw:image xlink:href="Pictures/image1.png"/></draw:frame>` +
    `<draw:g>` +
    `<draw:rect${geom(8, 15, 2.5, 3)}/>` +
    `<draw:rect${geom(11, 15, 2.5, 3)}/>` +
    `</draw:g>` +
    `<draw:frame${geom(16, 15, 8, 3)}><table:table><table:table-row><table:table-cell><text:p>Cell</text:p></table:table-cell></table:table-row></table:table></draw:frame>` +
    `</draw:page>`;

  const content =
    `<?xml version="1.0" encoding="UTF-8"?>` +
    `<office:document-content xmlns:office="${ODF_OFFICE_NS}" xmlns:style="${ODF_STYLE_NS}" xmlns:text="${ODF_TEXT_NS}"` +
    ` xmlns:draw="${ODF_DRAW_NS}" xmlns:fo="${ODF_FO_NS}" xmlns:svg="${ODF_SVG_NS}" xmlns:table="${ODF_TABLE_NS}"` +
    ` xmlns:xlink="${ODF_XLINK_NS}" xmlns:presentation="${ODF_PRESENTATION_NS}" office:version="1.2">` +
    `<office:automatic-styles>` +
    // A run style whose parent (in styles.xml) sets the colour and which adds
    // bold — exercises the cross-file parent-style-name chain (AC-19/20).
    `<style:style style:name="T1" style:family="text" style:parent-style-name="Emphasis">` +
    `<style:text-properties fo:font-weight="bold"/></style:style>` +
    `<style:style style:name="gr1" style:family="graphic">` +
    `<style:graphic-properties draw:fill="solid" draw:fill-color="#8faadc"/></style:style>` +
    `</office:automatic-styles>` +
    `<office:body><office:presentation>${page1}${page2}${page3}${page4}</office:presentation></office:body>` +
    `</office:document-content>`;

  const styles =
    `<?xml version="1.0" encoding="UTF-8"?>` +
    `<office:document-styles xmlns:office="${ODF_OFFICE_NS}" xmlns:style="${ODF_STYLE_NS}" xmlns:text="${ODF_TEXT_NS}"` +
    ` xmlns:draw="${ODF_DRAW_NS}" xmlns:fo="${ODF_FO_NS}" xmlns:svg="${ODF_SVG_NS}"` +
    ` xmlns:presentation="${ODF_PRESENTATION_NS}" office:version="1.2">` +
    `<office:styles>` +
    `<style:style style:name="Emphasis" style:family="text"><style:text-properties fo:color="#c0392b"/></style:style>` +
    `</office:styles>` +
    `<office:automatic-styles>` +
    `<style:page-layout style:name="PL1"><style:page-layout-properties fo:page-width="33.867cm" fo:page-height="19.05cm"/></style:page-layout>` +
    `</office:automatic-styles>` +
    `<office:master-styles>` +
    `<style:master-page style:name="Default" style:page-layout-name="PL1">` +
    // The master page's "body" placeholder — slide 4's body inherits its
    // geometry (AC-21).
    `<draw:frame presentation:class="body"${geom(2, 5, 14, 10)}><draw:text-box/></draw:frame>` +
    `</style:master-page>` +
    `</office:master-styles>` +
    `</office:document-styles>`;

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
    `<manifest:file-entry manifest:full-path="Pictures/image1.png" manifest:media-type="image/png"/>` +
    `</manifest:manifest>`;

  return zip([
    // ODF requires the mimetype entry first and STOREd so format sniffers can
    // read it from the raw bytes (same rule as odf/writer.ts).
    { name: "mimetype", data: encoder.encode("application/vnd.oasis.opendocument.presentation"), store: true },
    { name: "content.xml", data: encoder.encode(content) },
    { name: "styles.xml", data: encoder.encode(styles) },
    { name: "meta.xml", data: encoder.encode(meta) },
    { name: "Pictures/image1.png", data: IMAGE_PNG },
    { name: "META-INF/manifest.xml", data: encoder.encode(manifest) },
  ]);
}

// ============================================================================
// SPEC-085 per-effect transition fixtures
// ============================================================================
// Compact `.pptx`/`.odp` decks whose slides each carry one mapped transition
// (plus a final out-of-subset one), so the reader specs can assert that
// Impress/PowerPoint-authored markup parses to the expected model (AC-21) and
// the round-trip specs have real source bytes to preserve (AC-22).

interface ExpectedTransition {
  title: string;
  xml: string; // pptx p:transition markup / odp drawing-page-properties attrs
  expected: { type: string; direction?: string; durationMs?: number } | null; // null = out of subset
}

const PPTX_TRANSITION_ROWS: ExpectedTransition[] = [
  { title: "Fade", xml: `<p:transition spd="med"><p:fade/></p:transition>`, expected: { type: "fade", durationMs: 600 } },
  { title: "Push", xml: `<p:transition spd="fast"><p:push dir="l"/></p:transition>`, expected: { type: "push", direction: "left", durationMs: 300 } },
  { title: "Wipe", xml: `<p:transition spd="slow"><p:wipe dir="u"/></p:transition>`, expected: { type: "wipe", direction: "up", durationMs: 1000 } },
  { title: "Cover", xml: `<p:transition spd="med"><p:cover dir="r"/></p:transition>`, expected: { type: "cover", direction: "right", durationMs: 600 } },
  { title: "Uncover", xml: `<p:transition spd="med"><p:pull dir="d"/></p:transition>`, expected: { type: "uncover", direction: "down", durationMs: 600 } },
  { title: "Split", xml: `<p:transition spd="med"><p:split orient="horz" dir="out"/></p:transition>`, expected: { type: "split", durationMs: 600 } },
  { title: "Morph", xml: `<p:transition spd="med"><p:morph/></p:transition>`, expected: null },
];

/** The model each transition-fixture slide is expected to parse to (null =
 * out of subset, i.e. `slide.transition` unset + preserved). */
export const EXPECTED_TRANSITIONS = PPTX_TRANSITION_ROWS.map((row) => row.expected);
export const TRANSITION_FIXTURE_TITLES = PPTX_TRANSITION_ROWS.map((row) => row.title);

function pptxTransitionSlide(title: string, transitionXml: string): string {
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<p:sld xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}"><p:cSld><p:spTree>` +
    `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>` +
    pptxSp(2, "Title", { type: "title" }, [`<a:r><a:t>${title}</a:t></a:r>`], xfrm(2, 2, 29, 3)) +
    `</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>${transitionXml}</p:sld>`
  );
}

/** A minimal, structurally real `.pptx` with one slide per mapped transition
 * (SPEC-085 AC-22). Shares one blank layout/master/theme; each slide carries
 * its `p:transition` in schema position. */
export async function buildPptxTransitionsFixtureBytes(): Promise<Uint8Array> {
  const n = PPTX_TRANSITION_ROWS.length;
  const contentTypes =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
    `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
    `<Default Extension="xml" ContentType="application/xml"/>` +
    `<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>` +
    PPTX_TRANSITION_ROWS.map((_, i) => `<Override PartName="/ppt/slides/slide${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`).join("") +
    `<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>` +
    `<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>` +
    `<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>` +
    `</Types>`;
  const rootRels =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>` +
    `</Relationships>`;
  const sldIds = PPTX_TRANSITION_ROWS.map((_, i) => `<p:sldId id="${256 + i}" r:id="rId${i + 2}"/>`).join("");
  const presentation =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<p:presentation xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}">` +
    `<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>` +
    `<p:sldIdLst>${sldIds}</p:sldIdLst>` +
    `<p:sldSz cx="12192000" cy="6858000"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>`;
  const presentationRels =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>` +
    PPTX_TRANSITION_ROWS.map((_, i) => `<Relationship Id="rId${i + 2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${i + 1}.xml"/>`).join("") +
    `<Relationship Id="rId${n + 2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>` +
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
  const slideLayout1 =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<p:sldLayout xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}" type="blank"><p:cSld><p:spTree>` +
    `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld>` +
    `<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`;
  const layout1Rels =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>` +
    `</Relationships>`;

  const entry = (name: string, text: string) => ({ name, data: encoder.encode(text) });
  const parts = [
    entry("[Content_Types].xml", contentTypes),
    entry("_rels/.rels", rootRels),
    entry("ppt/presentation.xml", presentation),
    entry("ppt/_rels/presentation.xml.rels", presentationRels),
    entry("ppt/slideMasters/slideMaster1.xml", slideMaster),
    entry("ppt/slideMasters/_rels/slideMaster1.xml.rels", slideMasterRels),
    entry("ppt/slideLayouts/slideLayout1.xml", slideLayout1),
    entry("ppt/slideLayouts/_rels/slideLayout1.xml.rels", layout1Rels),
    entry("ppt/theme/theme1.xml", PPTX_THEME),
  ];
  PPTX_TRANSITION_ROWS.forEach((row, i) => {
    parts.push(entry(`ppt/slides/slide${i + 1}.xml`, pptxTransitionSlide(row.title, row.xml)));
    parts.push(entry(`ppt/slides/_rels/slide${i + 1}.xml.rels`, SLIDE_RELS));
  });
  return zip(parts);
}

const ODF_SMIL_NS = "urn:oasis:names:tc:opendocument:xmlns:smil-compatible:1.0";

// The drawing-page-properties attributes each odp transition slide carries.
const ODP_TRANSITION_PROPS: string[] = [
  `smil:type="fade" smil:subtype="crossfade" smil:dur="PT0.8S" presentation:transition-type="manual"`,
  `smil:type="pushWipe" smil:subtype="fromRight" smil:dur="PT0.8S" presentation:transition-type="manual"`,
  `smil:type="barWipe" smil:subtype="topToBottom" smil:direction="reverse" smil:dur="PT0.8S" presentation:transition-type="manual"`,
  `smil:type="slideWipe" smil:subtype="fromLeft" smil:dur="PT0.8S" presentation:transition-type="manual"`,
  `smil:type="slideWipe" smil:subtype="fromTop" smil:direction="reverse" smil:dur="PT0.8S" presentation:transition-type="manual"`,
  `smil:type="barnDoorWipe" smil:subtype="horizontal" smil:dur="PT0.8S" presentation:transition-type="manual"`,
  `smil:type="spiralWipe" smil:subtype="clockwiseTop" smil:dur="PT0.8S" presentation:transition-type="manual"`,
];

/** The odp model each transition-fixture page parses to (durations are 800ms
 * here, unlike the pptx twin's spd buckets). */
export const EXPECTED_ODP_TRANSITIONS = [
  { type: "fade", durationMs: 800 },
  { type: "push", direction: "left", durationMs: 800 },
  { type: "wipe", direction: "up", durationMs: 800 },
  { type: "cover", direction: "right", durationMs: 800 },
  { type: "uncover", direction: "down", durationMs: 800 },
  { type: "split", durationMs: 800 },
  null,
];

/** A minimal `.odp` with one page per mapped transition, each carrying a
 * `de-fixture-dpN` drawing-page automatic style with the smil attributes
 * (SPEC-085 AC-22). */
export async function buildOdpTransitionsFixtureBytes(): Promise<Uint8Array> {
  const styles = ODP_TRANSITION_PROPS.map(
    (props, i) => `<style:style style:name="fdp${i + 1}" style:family="drawing-page"><style:drawing-page-properties ${props}/></style:style>`,
  ).join("");
  const pages = PPTX_TRANSITION_ROWS.map(
    (row, i) => `<draw:page draw:name="page${i + 1}" draw:style-name="fdp${i + 1}">${odpFrame("title", [row.title], geom(2, 2, 29, 3))}</draw:page>`,
  ).join("");
  const content =
    `<?xml version="1.0" encoding="UTF-8"?>` +
    `<office:document-content xmlns:office="${ODF_OFFICE_NS}" xmlns:style="${ODF_STYLE_NS}" xmlns:text="${ODF_TEXT_NS}"` +
    ` xmlns:draw="${ODF_DRAW_NS}" xmlns:fo="${ODF_FO_NS}" xmlns:svg="${ODF_SVG_NS}"` +
    ` xmlns:presentation="${ODF_PRESENTATION_NS}" xmlns:smil="${ODF_SMIL_NS}" office:version="1.2">` +
    `<office:automatic-styles>${styles}</office:automatic-styles>` +
    `<office:body><office:presentation>${pages}</office:presentation></office:body>` +
    `</office:document-content>`;
  const styleParts =
    `<?xml version="1.0" encoding="UTF-8"?>` +
    `<office:document-styles xmlns:office="${ODF_OFFICE_NS}" xmlns:style="${ODF_STYLE_NS}" xmlns:fo="${ODF_FO_NS}"` +
    ` xmlns:svg="${ODF_SVG_NS}" xmlns:draw="${ODF_DRAW_NS}" xmlns:presentation="${ODF_PRESENTATION_NS}" office:version="1.2">` +
    `<office:automatic-styles>` +
    `<style:page-layout style:name="PL1"><style:page-layout-properties fo:page-width="33.867cm" fo:page-height="19.05cm"/></style:page-layout>` +
    `</office:automatic-styles>` +
    `<office:master-styles><style:master-page style:name="Default" style:page-layout-name="PL1"/></office:master-styles>` +
    `</office:document-styles>`;
  const meta =
    `<?xml version="1.0" encoding="UTF-8"?>` +
    `<office:document-meta xmlns:office="${ODF_OFFICE_NS}" xmlns:dc="http://purl.org/dc/elements/1.1/" office:version="1.2">` +
    `<office:meta><dc:title>Transitions fixture</dc:title></office:meta></office:document-meta>`;
  const manifest =
    `<?xml version="1.0" encoding="UTF-8"?>` +
    `<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.2">` +
    `<manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.presentation"/>` +
    `<manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>` +
    `<manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/>` +
    `<manifest:file-entry manifest:full-path="meta.xml" manifest:media-type="text/xml"/>` +
    `</manifest:manifest>`;
  return zip([
    { name: "mimetype", data: encoder.encode("application/vnd.oasis.opendocument.presentation"), store: true },
    { name: "content.xml", data: encoder.encode(content) },
    { name: "styles.xml", data: encoder.encode(styleParts) },
    { name: "meta.xml", data: encoder.encode(meta) },
    { name: "META-INF/manifest.xml", data: encoder.encode(manifest) },
  ]);
}

// ============================================================================
// SPEC-086 animation fixtures
// ============================================================================
// Compact `.pptx`/`.odp` decks exercising all three animation paths: a plain
// slide with a body placeholder + a shape (no timing — the e2e's authoring
// target), a slide carrying the canonical `p:timing`/`anim:*` forms (canonical
// read-back), and a slide with a deliberately exotic timing tree the reader can
// only preserve, not model (AC-29).

/** The animations the canonical fixture slide is expected to parse to (frame
 * ids are minted on read, so callers compare without them). */
export const EXPECTED_FIXTURE_ANIMATIONS: Omit<SlideAnimation, "frameId">[] = [
  { effect: "fadeIn", trigger: "click" },
  { effect: "flyIn", direction: "right", trigger: "afterPrevious", durationMs: 800 },
];

/** The canonical `p:timing` referencing the body (cNvPr id 3) and shape (id 4). */
export const PPTX_CANONICAL_TIMING = writePptxTiming(
  [
    { frameId: "3", effect: "fadeIn", trigger: "click" },
    { frameId: "4", effect: "flyIn", direction: "right", trigger: "afterPrevious", durationMs: 800 },
  ],
  (id) => Number(id),
);

/** An out-of-subset `p:timing` (an emphasis preset the table doesn't map) — the
 * reader preserves it whole. */
const PPTX_EXOTIC_TIMING =
  `<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>` +
  `<p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>` +
  `<p:par><p:cTn presetID="8" presetClass="emph" nodeType="clickEffect"><p:childTnLst>` +
  `<p:animEffect><p:cBhvr><p:cTn dur="500"/><p:tgtEl><p:spTgt spid="2"/></p:tgtEl></p:cBhvr></p:animEffect>` +
  `</p:childTnLst></p:cTn></p:par>` +
  `</p:childTnLst></p:cTn></p:seq></p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>`;

function pptxAnimSlide(shapes: string, timing = ""): string {
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<p:sld xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}"><p:cSld><p:spTree>` +
    `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>` +
    shapes +
    `</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>${timing}</p:sld>`
  );
}

/** Title + body + a rect shape; body is id 3, shape id 4 (the timing targets). */
function pptxAnimShapes(title: string): string {
  return (
    pptxSp(2, "Title", { type: "title" }, [`<a:r><a:t>${title}</a:t></a:r>`], xfrm(2, 1, 29, 3)) +
    pptxSp(3, "Body", { type: "body" }, ["<a:r><a:t>Body line</a:t></a:r>"], xfrm(2, 5, 20, 8)) +
    `<p:sp><p:nvSpPr><p:cNvPr id="4" name="Box"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
    `<p:spPr>${xfrm(24, 5, 6, 4)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>` +
    `<p:txBody><a:bodyPr/><a:p><a:r><a:t>Shape text</a:t></a:r></a:p></p:txBody></p:sp>`
  );
}

/** A minimal `.pptx` with a plain body+shape slide (e2e authoring target), a
 * canonical-timing slide, and an exotic-timing slide (SPEC-086 AC-29). */
export async function buildPptxAnimationsFixtureBytes(): Promise<Uint8Array> {
  const slides = [
    pptxAnimSlide(pptxAnimShapes("Animate me")),
    pptxAnimSlide(pptxAnimShapes("Canonical"), PPTX_CANONICAL_TIMING),
    pptxAnimSlide(pptxSp(2, "Title", { type: "title" }, ["<a:r><a:t>Exotic</a:t></a:r>"], xfrm(2, 1, 29, 3)), PPTX_EXOTIC_TIMING),
  ];
  const n = slides.length;
  const contentTypes =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
    `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
    `<Default Extension="xml" ContentType="application/xml"/>` +
    `<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>` +
    slides.map((_, i) => `<Override PartName="/ppt/slides/slide${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`).join("") +
    `<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>` +
    `<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>` +
    `<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>` +
    `</Types>`;
  const rootRels =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>` +
    `</Relationships>`;
  const sldIds = slides.map((_, i) => `<p:sldId id="${256 + i}" r:id="rId${i + 2}"/>`).join("");
  const presentation =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<p:presentation xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}">` +
    `<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>` +
    `<p:sldIdLst>${sldIds}</p:sldIdLst>` +
    `<p:sldSz cx="12192000" cy="6858000"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>`;
  const presentationRels =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>` +
    slides.map((_, i) => `<Relationship Id="rId${i + 2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${i + 1}.xml"/>`).join("") +
    `<Relationship Id="rId${n + 2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>` +
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
  const slideLayout1 =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<p:sldLayout xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}" type="obj"><p:cSld><p:spTree>` +
    `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld>` +
    `<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`;
  const layout1Rels =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>` +
    `</Relationships>`;
  const entry = (name: string, text: string) => ({ name, data: encoder.encode(text) });
  const parts = [
    entry("[Content_Types].xml", contentTypes),
    entry("_rels/.rels", rootRels),
    entry("ppt/presentation.xml", presentation),
    entry("ppt/_rels/presentation.xml.rels", presentationRels),
    entry("ppt/slideMasters/slideMaster1.xml", slideMaster),
    entry("ppt/slideMasters/_rels/slideMaster1.xml.rels", slideMasterRels),
    entry("ppt/slideLayouts/slideLayout1.xml", slideLayout1),
    entry("ppt/slideLayouts/_rels/slideLayout1.xml.rels", layout1Rels),
    entry("ppt/theme/theme1.xml", PPTX_THEME),
  ];
  slides.forEach((xml, i) => {
    parts.push(entry(`ppt/slides/slide${i + 1}.xml`, xml));
    parts.push(entry(`ppt/slides/_rels/slide${i + 1}.xml.rels`, SLIDE_RELS));
  });
  return zip(parts);
}

/** The canonical odp `anim:*` referencing the body (de-f1) and shape (de-f2). */
const ODP_CANONICAL_ANIM = writeOdpAnim(
  [
    { frameId: "de-f1", effect: "fadeIn", trigger: "click" },
    { frameId: "de-f2", effect: "flyIn", direction: "right", trigger: "afterPrevious", durationMs: 800 },
  ],
  (id) => id,
);

/** An out-of-subset odp `anim:*` (an emphasis preset the table doesn't map). */
const ODP_EXOTIC_ANIM =
  `<anim:par xmlns:anim="urn:oasis:names:tc:opendocument:xmlns:animation:1.0"` +
  ` xmlns:smil="urn:oasis:names:tc:opendocument:xmlns:smil-compatible:1.0"` +
  ` presentation:node-type="timing-root"><anim:seq presentation:node-type="main-sequence">` +
  `<anim:par smil:begin="next"><anim:par smil:begin="0s">` +
  `<anim:par presentation:node-type="on-click" presentation:preset-class="emphasis" presentation:preset-id="ooo-emphasis-spin" smil:begin="0s">` +
  `<anim:animate smil:targetElement="de-f1"/></anim:par>` +
  `</anim:par></anim:par></anim:seq></anim:par>`;

/** Title + body (de-f1) + rect shape (de-f2). */
function odpAnimFrames(title: string): string {
  return (
    odpFrame("title", [title], geom(2, 1, 29, 3)) +
    `<draw:frame presentation:class="body" xml:id="de-f1" draw:id="de-f1"${geom(2, 5, 20, 8)}><draw:text-box><text:p>Body line</text:p></draw:text-box></draw:frame>` +
    `<draw:rect xml:id="de-f2" draw:id="de-f2"${geom(24, 5, 6, 4)}><draw:text-box><text:p>Shape text</text:p></draw:text-box></draw:rect>`
  );
}

/** A minimal `.odp` twin of the animation fixture (SPEC-086 AC-29). */
export async function buildOdpAnimationsFixtureBytes(): Promise<Uint8Array> {
  const pages = [
    `<draw:page draw:name="page1">${odpAnimFrames("Animate me")}</draw:page>`,
    `<draw:page draw:name="page2">${odpAnimFrames("Canonical")}${ODP_CANONICAL_ANIM}</draw:page>`,
    `<draw:page draw:name="page3">${odpFrame("title", ["Exotic"], geom(2, 1, 29, 3))}${ODP_EXOTIC_ANIM}</draw:page>`,
  ].join("");
  const content =
    `<?xml version="1.0" encoding="UTF-8"?>` +
    `<office:document-content xmlns:office="${ODF_OFFICE_NS}" xmlns:style="${ODF_STYLE_NS}" xmlns:text="${ODF_TEXT_NS}"` +
    ` xmlns:draw="${ODF_DRAW_NS}" xmlns:fo="${ODF_FO_NS}" xmlns:svg="${ODF_SVG_NS}"` +
    ` xmlns:presentation="${ODF_PRESENTATION_NS}" office:version="1.2">` +
    `<office:body><office:presentation>${pages}</office:presentation></office:body>` +
    `</office:document-content>`;
  const styleParts =
    `<?xml version="1.0" encoding="UTF-8"?>` +
    `<office:document-styles xmlns:office="${ODF_OFFICE_NS}" xmlns:style="${ODF_STYLE_NS}" xmlns:fo="${ODF_FO_NS}"` +
    ` xmlns:svg="${ODF_SVG_NS}" xmlns:draw="${ODF_DRAW_NS}" xmlns:presentation="${ODF_PRESENTATION_NS}" office:version="1.2">` +
    `<office:automatic-styles>` +
    `<style:page-layout style:name="PL1"><style:page-layout-properties fo:page-width="33.867cm" fo:page-height="19.05cm"/></style:page-layout>` +
    `</office:automatic-styles>` +
    `<office:master-styles><style:master-page style:name="Default" style:page-layout-name="PL1">` +
    `<draw:frame presentation:class="body"${geom(2, 5, 20, 8)}><draw:text-box/></draw:frame>` +
    `</style:master-page></office:master-styles>` +
    `</office:document-styles>`;
  const meta =
    `<?xml version="1.0" encoding="UTF-8"?>` +
    `<office:document-meta xmlns:office="${ODF_OFFICE_NS}" xmlns:dc="http://purl.org/dc/elements/1.1/" office:version="1.2">` +
    `<office:meta><dc:title>Animations fixture</dc:title></office:meta></office:document-meta>`;
  const manifest =
    `<?xml version="1.0" encoding="UTF-8"?>` +
    `<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.2">` +
    `<manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.presentation"/>` +
    `<manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>` +
    `<manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/>` +
    `<manifest:file-entry manifest:full-path="meta.xml" manifest:media-type="text/xml"/>` +
    `</manifest:manifest>`;
  return zip([
    { name: "mimetype", data: encoder.encode("application/vnd.oasis.opendocument.presentation"), store: true },
    { name: "content.xml", data: encoder.encode(content) },
    { name: "styles.xml", data: encoder.encode(styleParts) },
    { name: "meta.xml", data: encoder.encode(meta) },
    { name: "META-INF/manifest.xml", data: encoder.encode(manifest) },
  ]);
}

// ============================================================================
// SPEC-089 layout catalogue fixtures
// ============================================================================
// A `.pptx` carrying the six standard layouts (title/obj/secHead/twoObj/
// titleOnly/blank) plus one exotic `type="cust"` layout (the "other" path), two
// slides — one referencing the title layout, one with NO layout reference at
// all — so the gallery, apply/reset and the no-reference fallback are all
// exercised. The `.odp` twin is built from odp/layoutGolden.ts's LAYOUT_GOLDEN.

interface Spec089Ph {
  type: string;
  idx?: number;
  x: number;
  y: number;
  w: number;
  h: number;
}
interface Spec089Layout {
  type: string;
  name: string;
  phs: Spec089Ph[];
}

const SPEC089_LAYOUTS: Spec089Layout[] = [
  { type: "title", name: "Title Slide", phs: [
    { type: "ctrTitle", x: 2, y: 6, w: 29, h: 3 }, { type: "subTitle", idx: 1, x: 2, y: 10, w: 29, h: 2 } ] },
  { type: "obj", name: "Title and Content", phs: [
    { type: "title", x: 2, y: 1, w: 29, h: 3 }, { type: "body", idx: 1, x: 2, y: 5, w: 29, h: 12 } ] },
  { type: "secHead", name: "Section Header", phs: [
    { type: "title", x: 2, y: 8, w: 29, h: 4 }, { type: "body", idx: 1, x: 2, y: 13, w: 29, h: 3 } ] },
  { type: "twoObj", name: "Two Content", phs: [
    { type: "title", x: 2, y: 1, w: 29, h: 3 }, { type: "body", idx: 1, x: 2, y: 5, w: 13, h: 12 }, { type: "body", idx: 2, x: 17, y: 5, w: 13, h: 12 } ] },
  { type: "titleOnly", name: "Title Only", phs: [{ type: "title", x: 2, y: 1, w: 29, h: 3 }] },
  { type: "blank", name: "Blank", phs: [] },
  // The exotic layout pins the "other" classification path (AC-5/19).
  { type: "cust", name: "Custom Panel", phs: [{ type: "title", x: 2, y: 1, w: 29, h: 3 }, { type: "body", idx: 1, x: 2, y: 5, w: 29, h: 12 }] },
];

function spec089LayoutPh(ph: Spec089Ph, id: number): string {
  const phAttr = `type="${ph.type}"${ph.idx !== undefined ? ` idx="${ph.idx}"` : ""}`;
  return (
    `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="PH ${id}"/><p:cNvSpPr/><p:nvPr><p:ph ${phAttr}/></p:nvPr></p:nvSpPr>` +
    `<p:spPr>${xfrm(ph.x, ph.y, ph.w, ph.h)}</p:spPr><p:txBody><a:bodyPr/><a:p/></p:txBody></p:sp>`
  );
}
function spec089LayoutXml(layout: Spec089Layout): string {
  const shapes = layout.phs.map((ph, i) => spec089LayoutPh(ph, i + 2)).join("");
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<p:sldLayout xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}" type="${layout.type}" preserve="1"><p:cSld name="${layout.name}"><p:spTree>` +
    `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>` +
    shapes +
    `</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`
  );
}

/** A `.pptx` with six standard layouts + one exotic, two slides (one with a
 * layout, one with no layout reference). */
export async function buildSpec089PptxFixtureBytes(): Promise<Uint8Array> {
  const n = SPEC089_LAYOUTS.length;
  const layoutRelsToMaster =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>` +
    `</Relationships>`;
  // Slide 1 references the title layout; slide 2 has no slideLayout relationship.
  const slide1 = pptxSlideXml(
    pptxSp(2, "Title 1", { type: "ctrTitle" }, ["<a:r><a:t>Welcome</a:t></a:r>"], xfrm(2, 6, 29, 3)) +
      pptxSp(3, "Subtitle 2", { type: "subTitle", idx: "1" }, ["<a:r><a:t>Layouts fixture</a:t></a:r>"], xfrm(2, 10, 29, 2)),
  );
  const slide1Rels =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>` +
    `</Relationships>`;
  const slide2 = pptxSlideXml(
    pptxSp(2, "TextBox 1", undefined, ["<a:r><a:t>No layout here</a:t></a:r>"], xfrm(2, 2, 24, 4)),
  );
  const slide2Rels =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    `</Relationships>`;

  const contentTypes =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
    `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
    `<Default Extension="xml" ContentType="application/xml"/>` +
    `<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>` +
    `<Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>` +
    `<Override PartName="/ppt/slides/slide2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>` +
    `<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>` +
    SPEC089_LAYOUTS.map((_, i) => `<Override PartName="/ppt/slideLayouts/slideLayout${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>`).join("") +
    `<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>` +
    `</Types>`;
  const rootRels =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>` +
    `</Relationships>`;
  const presentation =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<p:presentation xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}">` +
    `<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>` +
    `<p:sldIdLst><p:sldId id="256" r:id="rId2"/><p:sldId id="257" r:id="rId3"/></p:sldIdLst>` +
    `<p:sldSz cx="12192000" cy="6858000"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>`;
  const presentationRels =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>` +
    `<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>` +
    `<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide2.xml"/>` +
    `<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>` +
    `</Relationships>`;
  const master =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<p:sldMaster xmlns:a="${A_NS}" xmlns:r="${R_NS}" xmlns:p="${P_NS}"><p:cSld><p:spTree>` +
    `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld>` +
    `<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>` +
    `<p:sldLayoutIdLst>` +
    SPEC089_LAYOUTS.map((_, i) => `<p:sldLayoutId id="${2147483649 + i}" r:id="rId${i + 1}"/>`).join("") +
    `</p:sldLayoutIdLst></p:sldMaster>`;
  const masterRels =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    SPEC089_LAYOUTS.map((_, i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout${i + 1}.xml"/>`).join("") +
    `<Relationship Id="rId${n + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>` +
    `</Relationships>`;

  const entry = (name: string, text: string) => ({ name, data: encoder.encode(text) });
  const parts = [
    entry("[Content_Types].xml", contentTypes),
    entry("_rels/.rels", rootRels),
    entry("ppt/presentation.xml", presentation),
    entry("ppt/_rels/presentation.xml.rels", presentationRels),
    entry("ppt/slides/slide1.xml", slide1),
    entry("ppt/slides/_rels/slide1.xml.rels", slide1Rels),
    entry("ppt/slides/slide2.xml", slide2),
    entry("ppt/slides/_rels/slide2.xml.rels", slide2Rels),
    entry("ppt/slideMasters/slideMaster1.xml", master),
    entry("ppt/slideMasters/_rels/slideMaster1.xml.rels", masterRels),
    entry("ppt/theme/theme1.xml", PPTX_THEME),
  ];
  SPEC089_LAYOUTS.forEach((layout, i) => {
    parts.push(entry(`ppt/slideLayouts/slideLayout${i + 1}.xml`, spec089LayoutXml(layout)));
    parts.push(entry(`ppt/slideLayouts/_rels/slideLayout${i + 1}.xml.rels`, layoutRelsToMaster));
  });
  return zip(parts);
}

/** The `.odp` twin, its layout styles built from LAYOUT_GOLDEN (AC-19): three
 * pages referencing AL-title/AL-two/AL-exotic and one page with no layout
 * reference. */
export async function buildSpec089OdpFixtureBytes(): Promise<Uint8Array> {
  const page = (name: string, layout: string | null, cls: string, text: string): string =>
    `<draw:page draw:name="${name}" draw:master-page-name="Default"${layout ? ` presentation:presentation-page-layout-name="${layout}"` : ""}>` +
    odpFrame(cls, [text], geom(2, 6, 29, 3)) +
    `</draw:page>`;
  const pages =
    page("page1", GOLDEN_PAGE_LAYOUT_NAMES[0], "title", "Title layout") +
    page("page2", GOLDEN_PAGE_LAYOUT_NAMES[1], "title", "Two content layout") +
    page("page3", GOLDEN_PAGE_LAYOUT_NAMES[2], "title", "Exotic layout") +
    page("page4", null, "title", "No layout reference");
  const content =
    `<?xml version="1.0" encoding="UTF-8"?>` +
    `<office:document-content xmlns:office="${ODF_OFFICE_NS}" xmlns:style="${ODF_STYLE_NS}" xmlns:text="${ODF_TEXT_NS}"` +
    ` xmlns:draw="${ODF_DRAW_NS}" xmlns:fo="${ODF_FO_NS}" xmlns:svg="${ODF_SVG_NS}" xmlns:presentation="${ODF_PRESENTATION_NS}" office:version="1.2">` +
    `<office:automatic-styles></office:automatic-styles>` +
    `<office:body><office:presentation>${pages}</office:presentation></office:body>` +
    `</office:document-content>`;
  const styles =
    `<?xml version="1.0" encoding="UTF-8"?>` +
    `<office:document-styles xmlns:office="${ODF_OFFICE_NS}" xmlns:style="${ODF_STYLE_NS}" xmlns:text="${ODF_TEXT_NS}"` +
    ` xmlns:draw="${ODF_DRAW_NS}" xmlns:fo="${ODF_FO_NS}" xmlns:svg="${ODF_SVG_NS}" xmlns:presentation="${ODF_PRESENTATION_NS}" office:version="1.2">` +
    `<office:styles>${LAYOUT_GOLDEN}</office:styles>` +
    `<office:automatic-styles>` +
    `<style:page-layout style:name="PL1"><style:page-layout-properties fo:page-width="33.867cm" fo:page-height="19.05cm"/></style:page-layout>` +
    `</office:automatic-styles>` +
    `<office:master-styles><style:master-page style:name="Default" style:page-layout-name="PL1"/></office:master-styles>` +
    `</office:document-styles>`;
  const meta =
    `<?xml version="1.0" encoding="UTF-8"?>` +
    `<office:document-meta xmlns:office="${ODF_OFFICE_NS}" xmlns:dc="http://purl.org/dc/elements/1.1/" office:version="1.2">` +
    `<office:meta><dc:title>Layouts fixture</dc:title></office:meta></office:document-meta>`;
  const manifest =
    `<?xml version="1.0" encoding="UTF-8"?>` +
    `<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.2">` +
    `<manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.presentation"/>` +
    `<manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>` +
    `<manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/>` +
    `<manifest:file-entry manifest:full-path="meta.xml" manifest:media-type="text/xml"/>` +
    `</manifest:manifest>`;
  return zip([
    { name: "mimetype", data: encoder.encode("application/vnd.oasis.opendocument.presentation"), store: true },
    { name: "content.xml", data: encoder.encode(content) },
    { name: "styles.xml", data: encoder.encode(styles) },
    { name: "meta.xml", data: encoder.encode(meta) },
    { name: "META-INF/manifest.xml", data: encoder.encode(manifest) },
  ]);
}

// ============================================================================
// SPEC-091 table fixtures
// ============================================================================
// Built through the from-scratch writers (createBlankPresentation + insert +
// write) rather than hand-authored package bytes: each deck carries one
// modelled TableFrame (a 3×3 grid with a header row and banded rows) and one
// deliberately out-of-subset table (an irregular a:tbl / table:table) that must
// stay a foreign frame — pinning the preserve-don't-parse path (AC-17c).

/** The modelled table the SPEC-091 fixtures carry, for the unit spec and e2e to
 * assert after open (AC-17a). */
export const EXPECTED_FIXTURE_TABLE = {
  rows: 3,
  cols: 3,
  firstRow: true as const,
  bandRow: true as const,
  /** The header row's cell texts, left to right. */
  headerText: ["Region", "Q1", "Q2"],
  /** Cell (1,0) text (first body row). */
  bodyFirstCell: "North",
};

function fixtureTableFrame(widthCm: number, heightCm: number) {
  const table = createTableFrame(
    EXPECTED_FIXTURE_TABLE.rows,
    EXPECTED_FIXTURE_TABLE.cols,
    widthCm,
    heightCm,
  );
  table.bandRow = true;
  const text = [
    EXPECTED_FIXTURE_TABLE.headerText,
    [EXPECTED_FIXTURE_TABLE.bodyFirstCell, "12", "18"],
    ["South", "9", "15"],
  ];
  table.rows.forEach((row, r) => {
    row.cells.forEach((cell, c) => {
      cell.paragraphs = [{ runs: [{ text: text[r][c] }] }];
    });
  });
  return table;
}

/** A `p:graphicFrame` whose `a:tbl` is out of subset — a 2-column grid whose
 * single row carries three `a:tc`, so the reader aborts to a foreign frame. */
const OUT_OF_SUBSET_PPTX_TABLE =
  `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="99" name="Bad Table"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>` +
  `<p:xfrm><a:off x="${EMU(2)}" y="${EMU(2)}"/><a:ext cx="${EMU(6)}" cy="${EMU(2)}"/></p:xfrm>` +
  `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">` +
  `<a:tbl><a:tblGrid><a:gridCol w="${EMU(3)}"/><a:gridCol w="${EMU(3)}"/></a:tblGrid>` +
  `<a:tr h="${EMU(2)}">` +
  `<a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:t>bad</a:t></a:r></a:p></a:txBody></a:tc>` +
  `<a:tc><a:txBody><a:bodyPr/><a:p/></a:txBody></a:tc>` +
  `<a:tc><a:txBody><a:bodyPr/><a:p/></a:txBody></a:tc>` +
  `</a:tr></a:tbl></a:graphicData></a:graphic></p:graphicFrame>`;

/** A `draw:frame` whose `table:table` is out of subset — a two-column grid whose
 * row carries three cells. */
const OUT_OF_SUBSET_ODP_TABLE =
  `<draw:frame${geom(2, 2, 6, 2)}><table:table>` +
  `<table:table-column table:number-columns-repeated="2"/>` +
  `<table:table-row>` +
  `<table:table-cell><text:p>bad</text:p></table:table-cell>` +
  `<table:table-cell><text:p/></table:table-cell>` +
  `<table:table-cell><text:p/></table:table-cell>` +
  `</table:table-row></table:table></draw:frame>`;

export async function buildPptxTablesFixtureBytes(): Promise<Uint8Array> {
  const model = createBlankPresentation();
  const slide = model.slides[0];
  slide.frames = [
    fixtureTableFrame(model.slideWidthCm, model.slideHeightCm),
    {
      id: crypto.randomUUID(),
      kind: "foreign",
      label: "Table",
      xml: OUT_OF_SUBSET_PPTX_TABLE,
      xCm: 2,
      yCm: 2,
      widthCm: 6,
      heightCm: 2,
    },
  ];
  return (await writePptx(model)).bytes;
}

export async function buildOdpTablesFixtureBytes(): Promise<Uint8Array> {
  const model = createBlankPresentation();
  model.format = "odp";
  const slide = model.slides[0];
  slide.frames = [
    fixtureTableFrame(model.slideWidthCm, model.slideHeightCm),
    {
      id: crypto.randomUUID(),
      kind: "foreign",
      label: "Table",
      xml: OUT_OF_SUBSET_ODP_TABLE,
      xCm: 2,
      yCm: 2,
      widthCm: 6,
      heightCm: 2,
    },
  ];
  return (await writeOdp(model)).bytes;
}
