// The SPEC-055 caption fixture model, shared by the generator that writes the
// committed packages (scripts/build-spec055-fixtures.mjs) and the unit spec
// that guards them (fixtures.captions.spec.ts). It builds bytes and nothing
// else — no filesystem access, no side effects on import — so importing it can
// never rewrite a fixture the e2e suite may be reading (the same rule as
// fixtures.fields.ts).
//
// The fixtures carry three figure captions and two table captions interleaved
// plus one custom-label ("Diagram") caption, one caption above its object and
// the rest below, one caption separated from its object by a paragraph, a
// cross-reference of each of the six formats and one broken cross-reference,
// and a generated table of figures listing the figures.
import { writeDocx } from "./docx/writer";
import { writeOdt } from "./odf/writer";
import { CAPTION_STYLE_ID } from "./styles";
import type { MediaPart } from "./core/media";
import {
  createDocument,
  createFieldSpecRun,
  type Block,
  type DocumentModel,
  type Image,
  type NumberFormat,
  type Paragraph,
  type RefFormat,
  type Run,
  type Table,
  type TableOfFigures,
} from "./documentModel";

/** Pinned so a regenerated fixture's dcterms:modified / dc:date stays stable
 * (SPEC-008 round-trip determinism). */
export const FIXED_MODIFIED = new Date(Date.UTC(2026, 2, 4, 14, 30, 0));

// A 1×1 transparent PNG — enough for a valid image part without shipping a
// real screenshot; the caption, not the pixels, is what these fixtures test.
const PNG_1X1 =
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==";

const MEDIA: MediaPart = { mediaId: "m1", contentType: "image/png", data: PNG_1X1 };

function image(): Image {
  return {
    type: "image",
    mediaId: "m1",
    naturalWidthPx: 1,
    naturalHeightPx: 1,
    widthCm: 4,
    heightCm: 3,
    wrap: "inline",
  };
}

const cell = (t: string) => ({ blocks: [{ type: "paragraph", runs: [{ text: t, formatting: {} }] }] as Block[] });

function table(text: string): Table {
  return {
    type: "table",
    columnWidthsCm: [4, 4],
    rows: [{ cells: [cell(text), cell("")] }, { cells: [cell(""), cell("")] }],
  } as Table;
}

/** A Caption-styled paragraph carrying its label run, seq field, separator and
 * text, plus a hidden bookmark spanning it (its derived anchor) so a
 * cross-reference can target it. */
function caption(label: string, cached: string, text: string, anchor: string, format: NumberFormat = "decimal"): Paragraph {
  return {
    type: "paragraph",
    styleId: CAPTION_STYLE_ID,
    runs: [
      { text: `${label} `, formatting: {} },
      { text: "1", formatting: {}, field: { kind: "seq", label, format, cached } },
      { text: ": ", formatting: {} },
      { text, formatting: {} },
    ],
    bookmarks: [{ name: anchor, offset: 0, hidden: true }],
  };
}

function paragraph(text: string, runs?: Run[]): Paragraph {
  return { type: "paragraph", runs: runs ?? [{ text, formatting: {} }] };
}

function ref(target: string, show: RefFormat, cached: string): Run {
  return createFieldSpecRun(
    show === "page" ? { kind: "pageRef", target, show: "page", cached } : { kind: "ref", target, show, cached },
  );
}

const FIG1 = "_cap_Figure_1";
const FIG2 = "_cap_Figure_2";
const FIG3 = "_cap_Figure_3";
const TAB1 = "_cap_Table_1";
const TAB2 = "_cap_Table_2";
const DIA1 = "_cap_Diagram_1";

/** The shared caption fixture model. */
function captionModel(id: string, fileName: string): DocumentModel {
  const model = createDocument(id);
  const tof: TableOfFigures = {
    type: "tof",
    label: "Figure",
    options: { pageNumbers: true, dotLeader: true, hyperlinks: true },
    entries: [
      { text: "Figure 1: Revenue by quarter", page: 1, targetBookmark: FIG1 },
      { text: "Figure 2: Year-over-year growth", page: 1, targetBookmark: FIG2 },
      { text: "Figure 3: Forecast", page: 1, targetBookmark: FIG3 },
    ],
  };
  const references = paragraph("", [
    { text: "As shown in ", formatting: {} },
    ref(FIG1, "labelAndNumber", "Figure 1"),
    { text: " and ", formatting: {} },
    ref(FIG2, "full", "Figure 2: Year-over-year growth"),
    { text: ", growth is number ", formatting: {} },
    ref(FIG3, "numberOnly", "3"),
    { text: ". See also ", formatting: {} },
    ref(TAB1, "textOnly", "Q1 results"),
    { text: " on page ", formatting: {} },
    ref(TAB1, "page", "1"),
    { text: " (", formatting: {} },
    ref(TAB2, "position", "below"),
    { text: "). Broken: ", formatting: {} },
    ref("gone", "labelAndNumber", "Error! Reference source not found."),
    { text: ".", formatting: {} },
  ]);
  const blocks: Block[] = [
    { type: "paragraph", headingLevel: 1, runs: [{ text: "Quarterly Report", formatting: {} }] },
    references,
    // Figure 1: caption below its image.
    image(),
    caption("Figure", "1", "Revenue by quarter", FIG1),
    // Table 1: caption below its table.
    table("Q1 results"),
    caption("Table", "1", "Q1 results", TAB1),
    // Figure 2: caption ABOVE its image.
    caption("Figure", "2", "Year-over-year growth", FIG2),
    image(),
    // Figure 3: caption separated from its image by a paragraph (unbound).
    image(),
    paragraph("A separating sentence between the figure and its caption."),
    caption("Figure", "3", "Forecast", FIG3),
    // Table 2: caption below its table.
    table("Q2 results"),
    caption("Table", "2", "Q2 results", TAB2),
    // A custom-label caption.
    caption("Diagram", "1", "System architecture", DIA1),
    // The generated table of figures.
    tof,
  ];
  model.sections = [{ blocks }];
  model.media = [MEDIA];
  model.title = fileName;
  return model;
}

export function docxCaptionFixtureModel(): DocumentModel {
  return captionModel("word-captions", "word-captions.docx");
}

export function odtCaptionFixtureModel(): DocumentModel {
  const model = captionModel("lo-captions", "lo-captions.odt");
  model.format = "odt";
  return model;
}

/** File name → bytes for both committed fixtures, built with the pinned
 * `modified` stamp so two builds of an unchanged model agree. */
export async function buildCaptionFixtures(): Promise<{ name: string; bytes: Uint8Array }[]> {
  const docx = await writeDocx(docxCaptionFixtureModel(), { modified: FIXED_MODIFIED });
  const odt = await writeOdt(odtCaptionFixtureModel(), { modified: FIXED_MODIFIED });
  return [
    { name: "word-captions.docx", bytes: docx.bytes },
    { name: "lo-captions.odt", bytes: odt.bytes },
  ];
}
