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
import { defaultFieldRect } from "../headerFooter";
import { createTableFrame } from "../tables";
import { insertMediaFrame, setMediaPoster } from "../frames";
import { POSTER_PNG_BASE64, SYNTHETIC_MP4_BASE64, WAV_BASE64 } from "./mediaBytes";
import { writePptx } from "../pptx/writer";
import { writeOdp } from "../odp/writer";
import { unzip } from "../../core/zip";
import { chartPartXml } from "../../core/chartPartOoxml";
import { odfChartContentXml } from "../../core/chartPartOdf";
import type { ChartSeries, ChartType, ChartVisual } from "../../core/charts";
import type { DeckHeaderFooter, PresentationModel, Slide, SlideAnimation, SlideFieldPlaceholder, SlideRun, TextFrame } from "../model";
import { DEFAULT_THEME } from "../../core/theme";

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

// --- SPEC-092 chart fixtures (AC-28/28a/28b) -------------------------------
// Each deck carries a clustered column, a line, a pie and one deliberately
// out-of-catalogue chart (a `c:bar3DChart` for pptx, a `chart:class="chart:
// surface"` for odp) to pin the preserve-don't-parse path (AC-8/9b). The decks
// are assembled by package surgery on a from-scratch blank deck — the chart
// parts written by the hoisted pure string builders (`chartPartXml` /
// `odfChartContentXml`), the container + relationship/manifest spliced in — so
// the builders stay DOM-free (no reader), the constraint the e2e-fixture build
// script runs under (Node has no DOMParser). On open every part resolves: the
// three catalogue charts paint, the fourth shows the dashed placeholder.
const CHART_URI = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const CHART_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";
const CHART_PART_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";
const ODF_CHART_MEDIA_TYPE = "application/vnd.oasis.opendocument.chart";
const decoder = new TextDecoder("utf-8");

/** The chart data both fixtures carry, so the fixture spec + e2e assert one
 * expected shape against either format. Four categories; the column chart has
 * two series, the line and pie one each; the fourth chart is out of catalogue. */
export const EXPECTED_FIXTURE_CHARTS = {
  categories: ["Q1", "Q2", "Q3", "Q4"],
  column: {
    chartType: "columnClustered" as ChartType,
    series: [
      { name: "Revenue", color: "#4e79a7", values: [10, 20, 15, 25] },
      { name: "Cost", color: "#f28e2b", values: [6, 9, 7, 11] },
    ] as ChartSeries[],
  },
  line: {
    chartType: "line" as ChartType,
    series: [{ name: "Signups", color: "#59a14f", values: [3, 7, 9, 14] }] as ChartSeries[],
  },
  pie: {
    chartType: "pie" as ChartType,
    series: [{ name: "Share", color: "#e15759", values: [40, 25, 20, 15] }] as ChartSeries[],
  },
  /** The out-of-catalogue chart's marker in each format (AC-28a). */
  unsupported: { pptxMarker: "bar3DChart", odpMarker: "chart:surface" },
} as const;

/** The four chart rectangles, a 2×2 grid on the 33.867×19.05 cm slide (the
 * out-of-catalogue chart is the fourth). */
const CHART_RECTS = [
  { xCm: 1.5, yCm: 1, widthCm: 14, heightCm: 7 },
  { xCm: 18, yCm: 1, widthCm: 14, heightCm: 7 },
  { xCm: 1.5, yCm: 10.5, widthCm: 14, heightCm: 7 },
  { xCm: 18, yCm: 10.5, widthCm: 14, heightCm: 7 },
];

function chartVisualOf(spec: { chartType: ChartType; series: ChartSeries[] }): ChartVisual {
  return {
    chartType: spec.chartType,
    categories: [...EXPECTED_FIXTURE_CHARTS.categories],
    series: spec.series.map((s) => ({ name: s.name, color: s.color, values: [...s.values] })),
    options: { legend: "right", dataLabels: false, gridlines: true },
  };
}

/** A minimal but parseable out-of-catalogue OOXML chart part (`c:bar3DChart`). */
const PPTX_UNSUPPORTED_CHART =
  `<c:chartSpace xmlns:c="${CHART_URI}" xmlns:a="${A_NS}" xmlns:r="${R_NS}">` +
  `<c:chart><c:plotArea><c:layout/><c:bar3DChart><c:barDir val="col"/></c:bar3DChart></c:plotArea></c:chart></c:chartSpace>`;

function pptxChartGraphicFrame(id: number, rId: string, rect: { xCm: number; yCm: number; widthCm: number; heightCm: number }): string {
  return (
    `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${id}" name="Chart ${id}"/>` +
    `<p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>` +
    `<p:xfrm><a:off x="${EMU(rect.xCm)}" y="${EMU(rect.yCm)}"/><a:ext cx="${EMU(rect.widthCm)}" cy="${EMU(rect.heightCm)}"/></p:xfrm>` +
    `<a:graphic><a:graphicData uri="${CHART_URI}">` +
    `<c:chart xmlns:c="${CHART_URI}" xmlns:r="${R_NS}" r:id="${rId}"/>` +
    `</a:graphicData></a:graphic></p:graphicFrame>`
  );
}

export async function buildPptxChartsFixtureBytes(): Promise<Uint8Array> {
  const base = createBlankPresentation();
  base.slides[0].frames = [];
  const parts = await unzip((await writePptx(base)).bytes);

  const partXml = [
    chartPartXml(chartVisualOf(EXPECTED_FIXTURE_CHARTS.column)),
    chartPartXml(chartVisualOf(EXPECTED_FIXTURE_CHARTS.line)),
    chartPartXml(chartVisualOf(EXPECTED_FIXTURE_CHARTS.pie)),
    PPTX_UNSUPPORTED_CHART,
  ];
  const frames: string[] = [];
  const relRows: string[] = [];
  const ctRows: string[] = [];
  partXml.forEach((xml, i) => {
    const n = i + 1;
    const partName = `ppt/charts/chart${n}.xml`;
    const rId = `rId${50 + n}`;
    parts.set(partName, encoder.encode(xml));
    frames.push(pptxChartGraphicFrame(10 + n, rId, CHART_RECTS[i]));
    relRows.push(`<Relationship Id="${rId}" Type="${CHART_REL_TYPE}" Target="../charts/chart${n}.xml"/>`);
    ctRows.push(`<Override PartName="/${partName}" ContentType="${CHART_PART_CONTENT_TYPE}"/>`);
  });

  const slide1 = decoder.decode(parts.get("ppt/slides/slide1.xml")!).replace("</p:spTree>", `${frames.join("")}</p:spTree>`);
  parts.set("ppt/slides/slide1.xml", encoder.encode(slide1));
  const relsPath = "ppt/slides/_rels/slide1.xml.rels";
  const rels = decoder.decode(parts.get(relsPath)!).replace("</Relationships>", `${relRows.join("")}</Relationships>`);
  parts.set(relsPath, encoder.encode(rels));
  const ct = decoder.decode(parts.get("[Content_Types].xml")!).replace("</Types>", `${ctRows.join("")}</Types>`);
  parts.set("[Content_Types].xml", encoder.encode(ct));

  return zip([...parts].map(([name, data]) => ({ name, data })));
}

/** The out-of-catalogue ODP chart object: a valid `office:chart` whose
 * `chart:class` is out of catalogue (`chart:surface`). */
function odpUnsupportedChartContent(): string {
  return odfChartContentXml(chartVisualOf({ chartType: "columnClustered", series: [EXPECTED_FIXTURE_CHARTS.column.series[0]] }), 14, 7).replace(
    'chart:class="chart:bar"',
    'chart:class="chart:surface"',
  );
}

export async function buildOdpChartsFixtureBytes(): Promise<Uint8Array> {
  const base = createBlankPresentation();
  base.format = "odp";
  base.slides[0].frames = [];
  const parts = await unzip((await writeOdp(base)).bytes);

  const objectContent = [
    odfChartContentXml(chartVisualOf(EXPECTED_FIXTURE_CHARTS.column), 14, 7),
    odfChartContentXml(chartVisualOf(EXPECTED_FIXTURE_CHARTS.line), 14, 7),
    odfChartContentXml(chartVisualOf(EXPECTED_FIXTURE_CHARTS.pie), 14, 7),
    odpUnsupportedChartContent(),
  ];
  const frames: string[] = [];
  const manifestRows: string[] = [];
  objectContent.forEach((xml, i) => {
    const n = i + 1;
    const dir = `Object ${n}`;
    parts.set(`${dir}/content.xml`, encoder.encode(xml));
    const r = CHART_RECTS[i];
    frames.push(
      `<draw:frame presentation:class="chart"${geom(r.xCm, r.yCm, r.widthCm, r.heightCm)}>` +
        `<draw:object xlink:href="./${dir}" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/></draw:frame>`,
    );
    manifestRows.push(
      `<manifest:file-entry manifest:full-path="${dir}/" manifest:media-type="${ODF_CHART_MEDIA_TYPE}"/>` +
        `<manifest:file-entry manifest:full-path="${dir}/content.xml" manifest:media-type="text/xml"/>`,
    );
  });

  const content = decoder.decode(parts.get("content.xml")!).replace("</draw:page>", `${frames.join("")}</draw:page>`);
  parts.set("content.xml", encoder.encode(content));
  const manifest = decoder
    .decode(parts.get("META-INF/manifest.xml")!)
    .replace("</manifest:manifest>", `${manifestRows.join("")}</manifest:manifest>`);
  parts.set("META-INF/manifest.xml", encoder.encode(manifest));

  // `mimetype` must stay first and STOREd in an ODF package.
  const ordered = [...parts].sort(([a], [b]) => (a === "mimetype" ? -1 : b === "mimetype" ? 1 : 0));
  return zip(ordered.map(([name, data]) => ({ name, data, store: name === "mimetype" })));
}

// ============================================================================
// SPEC-093 header/footer fixtures
// ============================================================================
// Built through the from-scratch writers (createBlankPresentation + configure +
// write), not hand-authored bytes, so e2e and unit assertions never drift. A
// four-slide deck: a title slide with fields suppressed (AC-97), two body
// slides showing footer/slide-number/date (one carrying a per-slide footer
// override, AC-98), and a slide with a deliberately out-of-subset date
// (`datetime13` on pptx / an Impress `D1` data style on odp) pinning the
// preserve-don't-parse path (AC-99). The four mini-deck slides above are
// untouched.

/** The deck-level header/footer the SPEC-093 fixtures configure (the pptx
 * read-back; odp cannot represent `showOnTitleSlide` so its round-trip drops
 * that one key). Auto date in the default `d MMMM yyyy` so the e2e's
 * Apply-to-all keeps the default and AC-111 can pin the rendered value. */
export const EXPECTED_HEADER_FOOTER: DeckHeaderFooter = {
  footerText: "Confidential",
  showOnTitleSlide: false,
};

/** The per-slide footer override the fixture's third slide carries (AC-98). */
export const EXPECTED_HEADER_FOOTER_OVERRIDE = "Draft copy";

function hfField(kind: SlideFieldPlaceholder["kind"], w: number, h: number, extra: Partial<SlideFieldPlaceholder> = {}): SlideFieldPlaceholder {
  const rect = defaultFieldRect(kind, w, h);
  const entry: SlideFieldPlaceholder = { kind, xCm: rect.xCm, yCm: rect.yCm, widthCm: rect.widthCm, heightCm: rect.heightCm, ...extra };
  if (rect.align !== "left" && entry.align === undefined) {
    entry.align = rect.align;
  }
  return entry;
}

function hfBody(title: string, w: number, h: number): Slide["frames"][number] {
  return {
    id: crypto.randomUUID(),
    kind: "text",
    role: "title",
    paragraphs: [{ runs: [{ text: title }] }],
    xCm: w * 0.08,
    yCm: h * 0.08,
    widthCm: w * 0.84,
    heightCm: h * 0.16,
  };
}

/** The fixture model, parameterised by format for the out-of-subset date's wire
 * form (pptx keeps a `datetime13` cached string; odp keeps a foreign `D1` data
 * style, ODF caching nothing). */
function headerFooterFixtureModel(format: "pptx" | "odp"): PresentationModel {
  const model = createBlankPresentation();
  model.format = format;
  const w = model.slideWidthCm;
  const h = model.slideHeightCm;

  // Slide 0: the title slide, fields suppressed (no field entries + the deck's
  // showOnTitleSlide=false policy). It keeps createBlank's title + subtitle
  // placeholders so it classifies as the `title` layout (isTitleSlide true).
  const title0 = model.slides[0].frames.find((f) => f.kind === "text" && f.role === "title");
  if (title0 && title0.kind === "text") {
    title0.paragraphs = [{ runs: [{ text: "Header & footer demo" }] }];
  }

  const bodyOne: Slide = {
    id: crypto.randomUUID(),
    frames: [hfBody("Agenda", w, h)],
    fields: [hfField("footer", w, h), hfField("slideNumber", w, h), hfField("date", w, h)],
  };
  const bodyTwo: Slide = {
    id: crypto.randomUUID(),
    frames: [hfBody("Details", w, h)],
    fields: [
      hfField("footer", w, h, { text: EXPECTED_HEADER_FOOTER_OVERRIDE }),
      hfField("slideNumber", w, h),
      hfField("date", w, h),
    ],
  };
  const outOfSubsetDate: SlideFieldPlaceholder =
    format === "pptx"
      ? hfField("date", w, h, { wireType: "datetime13", text: "14.5.2026" })
      : hfField("date", w, h, { wireStyleName: "D1" });
  const bodyThree: Slide = {
    id: crypto.randomUUID(),
    frames: [hfBody("Appendix", w, h)],
    fields: [hfField("footer", w, h), hfField("slideNumber", w, h), outOfSubsetDate],
  };

  model.slides.push(bodyOne, bodyTwo, bodyThree);
  const headerFooter: DeckHeaderFooter = { footerText: "Confidential", showOnTitleSlide: false };
  model.headerFooter = headerFooter;
  return model;
}

/** A fixed instant so the checked-in fixtures are byte-deterministic across
 * regenerations (both the recomputed field values and the meta/core stamp). */
const HEADER_FOOTER_FIXTURE_INSTANT = new Date(Date.UTC(2026, 4, 14, 9, 0, 0));

export async function buildPptxHeaderFooterFixtureBytes(): Promise<Uint8Array> {
  return (await writePptx(headerFooterFixtureModel("pptx"), {
    now: HEADER_FOOTER_FIXTURE_INSTANT,
    modified: HEADER_FOOTER_FIXTURE_INSTANT,
  })).bytes;
}

export async function buildOdpHeaderFooterFixtureBytes(): Promise<Uint8Array> {
  return (await writeOdp(headerFooterFixtureModel("odp"), {
    now: HEADER_FOOTER_FIXTURE_INSTANT,
    modified: HEADER_FOOTER_FIXTURE_INSTANT,
  })).bytes;
}

// ============================================================================
// SPEC-094 deck-theme fixtures
// ============================================================================
// Built through the from-scratch writers (createBlankPresentation + a themed
// model + write), never mutating the four mini-deck slides above. Slide 1
// mixes every reference kind the spec models — an accent1 fill, an
// accent1 + lumMod 60% / lumOff 40% shape, a `+mn-lt` run, a `+mj-lt` title
// (whose colour is a dk2 reference so a theme apply visibly recolours it) and
// an explicitly-coloured run. Slide 2 pins the preserve path: pptx carries a
// `p:clrMapOvr` this editor cannot parse (its refs write literals, the panel
// says "opted out"); odp carries an unrecognised extension theme block riding
// verbatim inside styles.xml.

/** The references slide 1 carries, exported so the round-trip and e2e specs
 * assert the exact same objects. */
export const SPEC094_REFS = {
  titleColor: { theme: "dk2" },
  titleFont: { theme: "major" },
  bodyFont: { theme: "minor" },
  accentFill: { theme: "accent1" },
  lighterFill: { theme: "accent1", lumMod: 0.6, lumOff: 0.4 },
} as const;

/** The literal the explicitly-coloured run keeps through every theme apply. */
export const SPEC094_LITERAL_RUN_COLOR = "#ff0000";

/** A `p:clrMapOvr` this editor cannot parse — the slide-2 opt-out (AC-23/24). */
export const SPEC094_CLR_MAP_OVR = `<p:clrMapOvr><p:extLst uri="urn:example:unparsed"/></p:clrMapOvr>`;

/** The unrecognised extension theme block the odp twin's styles.xml carries —
 * a shape this spec does not claim to know, preserved verbatim. */
export const SPEC094_ODP_EXTENSION_BLOCK =
  `<loext:theme xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0"` +
  ` loext:name="Vendor"><loext:color-table loext:name="Vendor"/></loext:theme>`;

function spec094TextFrame(
  text: string,
  box: [number, number, number, number],
  run: Omit<SlideRun, "text">,
  role?: "title",
): TextFrame {
  const frame: TextFrame = {
    id: crypto.randomUUID(),
    kind: "text",
    xCm: box[0],
    yCm: box[1],
    widthCm: box[2],
    heightCm: box[3],
    paragraphs: [{ runs: [{ text, ...run }] }],
  };
  if (role) {
    frame.role = role;
    frame.phType = "title";
  }
  return frame;
}

/** The mixed themed deck (AC-53/54), parameterised by format. */
function spec094FixtureModel(format: "pptx" | "odp"): PresentationModel {
  const model = createBlankPresentation();
  model.format = format;
  model.title = "Theme fixture";
  model.theme = {
    name: DEFAULT_THEME.name,
    colors: { ...DEFAULT_THEME.colors },
    fonts: { ...DEFAULT_THEME.fonts },
  };

  const mixed: Slide = {
    id: crypto.randomUUID(),
    frames: [
      spec094TextFrame("Theme fixture", [2, 1, 29, 3], {
        color: SPEC094_REFS.titleColor,
        fontFamily: SPEC094_REFS.titleFont,
        sizePt: 40,
      }, "title"),
      spec094TextFrame("Body follows the theme", [2, 5, 20, 2], { fontFamily: SPEC094_REFS.bodyFont }),
      {
        id: crypto.randomUUID(),
        kind: "shape",
        geometry: "rect",
        xCm: 2,
        yCm: 9,
        widthCm: 8,
        heightCm: 5,
        fill: { type: "solid", color: { ...SPEC094_REFS.accentFill } },
      },
      {
        id: crypto.randomUUID(),
        kind: "shape",
        geometry: "ellipse",
        xCm: 12,
        yCm: 9,
        widthCm: 8,
        heightCm: 5,
        fill: { type: "solid", color: { ...SPEC094_REFS.lighterFill } },
      },
      spec094TextFrame("Fixed red", [22, 9, 8, 2], { color: SPEC094_LITERAL_RUN_COLOR }),
    ],
  };

  const preservePath: Slide = {
    id: crypto.randomUUID(),
    frames: [
      spec094TextFrame("Preserve path", [2, 1, 20, 2], { color: "#004488" }),
    ],
  };
  if (format === "pptx") {
    preservePath.preserved = { "p:clrMapOvr": SPEC094_CLR_MAP_OVR };
  }

  model.slides = [mixed, preservePath];
  return model;
}

const SPEC094_FIXTURE_INSTANT = new Date(Date.UTC(2026, 8, 1, 9, 0, 0));

export async function buildSpec094PptxFixtureBytes(): Promise<Uint8Array> {
  return (await writePptx(spec094FixtureModel("pptx"), {
    now: SPEC094_FIXTURE_INSTANT,
    modified: SPEC094_FIXTURE_INSTANT,
  })).bytes;
}

/** The odp twin, with the unrecognised extension theme block spliced verbatim
 * into styles.xml's `office:styles` after the write (the preserve pin). */
export async function buildSpec094OdpFixtureBytes(): Promise<Uint8Array> {
  const bytes = (await writeOdp(spec094FixtureModel("odp"), {
    now: SPEC094_FIXTURE_INSTANT,
    modified: SPEC094_FIXTURE_INSTANT,
  })).bytes;
  const parts = await unzip(bytes);
  const styles = new TextDecoder().decode(parts.get("styles.xml")!);
  parts.set(
    "styles.xml",
    encoder.encode(styles.replace("</office:styles>", `${SPEC094_ODP_EXTENSION_BLOCK}</office:styles>`)),
  );
  const ordered = [...parts].sort(([a], [b]) => (a === "mimetype" ? -1 : b === "mimetype" ? 1 : 0));
  return zip(ordered.map(([name, data]) => ({ name, data, store: name === "mimetype" })));
}

// --- SPEC-096 object fixtures (spec096-objects) -----------------------------
// A one-slide deck for the multi-select / rotate / adjust / group work:
// three deliberately unaligned shapes, a 37°-rotated image, a roundRect with
// a non-default adjustment — all written through the real writers — plus,
// spliced verbatim into the written slide/page, a group whose a:chOff/a:chExt
// DIFFER from a:off/a:ext (pinning the child mapping) and a ROTATED verbatim
// group (pinning the confirm-to-release path; pptx only — odp models no group
// rotation). The four mini-deck slides above are never touched.

/** The three unaligned shapes' boxes in cm ([x, y, w, h], z-order) —
 * deliberately ragged: x all different, vertical gaps 0cm and 5cm so a
 * Distribute vertically visibly equalises them (e2e). */
export const SPEC096_SHAPE_BOXES: [number, number, number, number][] = [
  [2, 2, 4, 2],
  [9, 4, 3, 3],
  [14, 12, 5, 2],
];
/** The rotated image's box and rotation. */
export const SPEC096_IMAGE_BOX: [number, number, number, number] = [24, 2, 4, 3];
export const SPEC096_IMAGE_ROTATION = 37;
/** The tuned roundRect: box + its stored corner-radius fraction. */
export const SPEC096_ROUND_RECT_BOX: [number, number, number, number] = [2, 16, 6, 2.5];
export const SPEC096_ADJUSTMENT = 0.4;
/** The non-identity verbatim group's slide-space box (chExt is HALF the ext,
 * so its one child maps through a ×2 scale). */
export const SPEC096_GROUP_BOX: [number, number, number, number] = [20, 12, 4, 3];
/** The rotated verbatim group's box + rotation (pptx only). */
export const SPEC096_ROTATED_GROUP_BOX: [number, number, number, number] = [28, 12, 4, 3];
export const SPEC096_ROTATED_GROUP_DEG = 45;

const SPEC096_PNG_BASE64 =
  "iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAIAAAD91JpzAAAAEklEQVR4nGMw7jhj3HGGAUIBACiuBh3+iAU4AAAAAElFTkSuQmCC";

function spec096Model(format: "pptx" | "odp"): PresentationModel {
  const model = createBlankPresentation();
  model.format = format;
  model.title = "Objects fixture";
  const mediaId = format === "odp" ? "Pictures/image1.png" : "ppt/media/image1.png";
  const fills = ["#4472c4", "#c00000", "#00b050"];
  const slide: Slide = {
    id: crypto.randomUUID(),
    frames: [
      ...SPEC096_SHAPE_BOXES.map(([x, y, w, h], i) => ({
        id: crypto.randomUUID(),
        kind: "shape" as const,
        geometry: "rect" as const,
        xCm: x,
        yCm: y,
        widthCm: w,
        heightCm: h,
        fill: { type: "solid" as const, color: fills[i] },
      })),
      {
        id: crypto.randomUUID(),
        kind: "image" as const,
        xCm: SPEC096_IMAGE_BOX[0],
        yCm: SPEC096_IMAGE_BOX[1],
        widthCm: SPEC096_IMAGE_BOX[2],
        heightCm: SPEC096_IMAGE_BOX[3],
        rotationDeg: SPEC096_IMAGE_ROTATION,
        mediaId,
      },
      {
        id: crypto.randomUUID(),
        kind: "shape" as const,
        geometry: "roundRect" as const,
        xCm: SPEC096_ROUND_RECT_BOX[0],
        yCm: SPEC096_ROUND_RECT_BOX[1],
        widthCm: SPEC096_ROUND_RECT_BOX[2],
        heightCm: SPEC096_ROUND_RECT_BOX[3],
        adjustments: [SPEC096_ADJUSTMENT],
        fill: { type: "solid" as const, color: "#ffc000" },
      },
    ],
  };
  model.slides = [slide];
  model.media = [{ mediaId, contentType: "image/png", data: SPEC096_PNG_BASE64 }];
  return model;
}

/** The verbatim pptx groups spliced into the written slide: one whose
 * chOff/chExt are NOT the identity mapping (child in a half-scale child
 * space), and one carrying a 45° rotation on its own xfrm. */
function spec096PptxGroupsXml(): string {
  const [gx, gy, gw, gh] = SPEC096_GROUP_BOX;
  const [rx, ry, rw, rh] = SPEC096_ROTATED_GROUP_BOX;
  const nonIdentity =
    `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="90" name="Fixture Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
    `<p:grpSpPr><a:xfrm><a:off x="${EMU(gx)}" y="${EMU(gy)}"/><a:ext cx="${EMU(gw)}" cy="${EMU(gh)}"/>` +
    `<a:chOff x="0" y="0"/><a:chExt cx="${EMU(gw / 2)}" cy="${EMU(gh / 2)}"/></a:xfrm></p:grpSpPr>` +
    `<p:sp><p:nvSpPr><p:cNvPr id="91" name="GChild 1"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
    `<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${EMU(gw / 4)}" cy="${EMU(gh / 2)}"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:sp>` +
    `</p:grpSp>`;
  const rotated =
    `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="92" name="Rotated Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
    `<p:grpSpPr><a:xfrm rot="${SPEC096_ROTATED_GROUP_DEG * 60000}"><a:off x="${EMU(rx)}" y="${EMU(ry)}"/><a:ext cx="${EMU(rw)}" cy="${EMU(rh)}"/>` +
    `<a:chOff x="${EMU(rx)}" y="${EMU(ry)}"/><a:chExt cx="${EMU(rw)}" cy="${EMU(rh)}"/></a:xfrm></p:grpSpPr>` +
    `<p:sp><p:nvSpPr><p:cNvPr id="93" name="GChild 2"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
    `<p:spPr><a:xfrm><a:off x="${EMU(rx)}" y="${EMU(ry)}"/><a:ext cx="${EMU(rw / 2)}" cy="${EMU(rh)}"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:sp>` +
    `</p:grpSp>`;
  return nonIdentity + rotated;
}

/** The verbatim odp group spliced into the written page (a draw:g has no box
 * of its own — its bounds are the children's union). */
const SPEC096_ODP_GROUP =
  `<draw:g>` +
  `<draw:rect svg:x="20cm" svg:y="12cm" svg:width="2cm" svg:height="1.5cm"/>` +
  `<draw:rect svg:x="22cm" svg:y="13.5cm" svg:width="2cm" svg:height="1.5cm"/>` +
  `</draw:g>`;

const SPEC096_FIXTURE_INSTANT = new Date(Date.UTC(2026, 8, 1, 9, 0, 0));

export async function buildSpec096PptxFixtureBytes(): Promise<Uint8Array> {
  const bytes = (
    await writePptx(spec096Model("pptx"), { now: SPEC096_FIXTURE_INSTANT, modified: SPEC096_FIXTURE_INSTANT })
  ).bytes;
  const parts = await unzip(bytes);
  const slide = new TextDecoder().decode(parts.get("ppt/slides/slide1.xml")!);
  parts.set(
    "ppt/slides/slide1.xml",
    encoder.encode(slide.replace("</p:spTree>", `${spec096PptxGroupsXml()}</p:spTree>`)),
  );
  return zip([...parts].map(([name, data]) => ({ name, data })));
}

export async function buildSpec096OdpFixtureBytes(): Promise<Uint8Array> {
  const bytes = (
    await writeOdp(spec096Model("odp"), { now: SPEC096_FIXTURE_INSTANT, modified: SPEC096_FIXTURE_INSTANT })
  ).bytes;
  const parts = await unzip(bytes);
  const content = new TextDecoder().decode(parts.get("content.xml")!);
  parts.set("content.xml", encoder.encode(content.replace("</draw:page>", `${SPEC096_ODP_GROUP}</draw:page>`)));
  const ordered = [...parts].sort(([a], [b]) => (a === "mimetype" ? -1 : b === "mimetype" ? 1 : 0));
  return zip(ordered.map(([name, data]) => ({ name, data, store: name === "mimetype" })));
}

// ============================================================================
// SPEC-095 media fixtures — built through the from-scratch writers. Each deck's
// one slide carries an embedded audio clip (a playable PCM WAV, with a poster),
// an embedded video clip (a synthetic MP4 the browser cannot decode, exercising
// the unsupported degrade), a linked clip (a placeholder, never fetched) and one
// deliberately out-of-subset media element that must stay a "Media" ForeignFrame
// (pinning preserve-don't-parse).

/** What the SPEC-095 fixtures carry, for the unit spec and e2e to assert. */
export const EXPECTED_MEDIA_FRAMES = {
  /** Modelled media frames (audio + video + linked); the out-of-subset element
   * stays a foreign frame and is counted separately. */
  count: 3,
  audioContentType: "audio/wav",
  videoContentType: "video/mp4",
  /** The linked clip's caption basename. */
  linkedBasename: "remote.mp4",
};

const SPEC095_FIXTURE_INSTANT = new Date(Date.UTC(2026, 8, 1, 9, 0, 0));

const OUT_OF_SUBSET_PPTX_MEDIA =
  `<p:pic><p:nvPicPr><p:cNvPr id="99" name="Media 99"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>` +
  `<p:blipFill><a:blip/><a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
  `<p:spPr><a:xfrm><a:off x="720000" y="5400000"/><a:ext cx="1440000" cy="720000"/></a:xfrm>` +
  `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`;

const OUT_OF_SUBSET_ODP_MEDIA =
  `<draw:frame svg:x="2cm" svg:y="15cm" svg:width="4cm" svg:height="2cm">` +
  `<draw:image xlink:href="Pictures/missing.png" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/>` +
  `</draw:frame>`;

function mediaFixtureModel(format: "pptx" | "odp"): PresentationModel {
  const model = createBlankPresentation();
  model.format = format;
  const slide = model.slides[0];
  slide.frames = [];
  const audioId = insertMediaFrame(model, 0, { mediaId: "", contentType: "audio/wav", data: WAV_BASE64 }, "audio")!;
  const videoId = insertMediaFrame(model, 0, { mediaId: "", contentType: "video/mp4", data: SYNTHETIC_MP4_BASE64 }, "video")!;
  setMediaPoster(model, 0, audioId, { mediaId: "", contentType: "image/png", data: POSTER_PNG_BASE64 });
  // Distinct, non-overlapping boxes so each clip is separately clickable on the
  // stage (insertMediaFrame centres them, which would stack audio under video).
  const audio = slide.frames.find((f) => f.id === audioId)!;
  audio.xCm = 1;
  audio.yCm = 1;
  const video = slide.frames.find((f) => f.id === videoId)!;
  video.xCm = 2;
  video.yCm = 8;
  slide.frames.push({
    id: crypto.randomUUID(),
    kind: "media",
    media: "video",
    xCm: 20,
    yCm: 2,
    widthCm: 6,
    heightCm: 4,
    linkTarget: "../media/remote.mp4",
    contentType: "video/mp4",
  });
  slide.frames.push({
    id: crypto.randomUUID(),
    kind: "foreign",
    label: "Media",
    xCm: 2,
    yCm: 15,
    widthCm: 4,
    heightCm: 2,
    xml: format === "odp" ? OUT_OF_SUBSET_ODP_MEDIA : OUT_OF_SUBSET_PPTX_MEDIA,
  });
  return model;
}

export async function buildPptxMediaFixtureBytes(): Promise<Uint8Array> {
  return (
    await writePptx(mediaFixtureModel("pptx"), { now: SPEC095_FIXTURE_INSTANT, modified: SPEC095_FIXTURE_INSTANT })
  ).bytes;
}

export async function buildOdpMediaFixtureBytes(): Promise<Uint8Array> {
  return (
    await writeOdp(mediaFixtureModel("odp"), { now: SPEC095_FIXTURE_INSTANT, modified: SPEC095_FIXTURE_INSTANT })
  ).bytes;
}
