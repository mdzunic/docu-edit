// The SPEC-077 theme fixture model (AC-51), shared by the generator that
// writes the committed packages (scripts/build-spec077-fixtures.mjs) and the
// unit spec that guards them (fixtures.theme.spec.ts). It builds bytes and
// nothing else — no filesystem access, no side effects on import — so
// importing it can never rewrite a fixture the e2e suite may be reading.
//
// The fixtures carry exactly what the AC enumerates: a theme-coloured
// heading (an accent2 run colour plus the major theme font), a theme-tinted
// table (accent1 at Lighter 80% cell shading), a theme-filled shape
// (accent4), and an explicitly-coloured run that no theme change may move.
// The document's own theme is a hand-rolled scheme distinct from every
// built-in, so a reopened fixture provably resolves against ITS theme and
// not the product default, and the gallery's "current" card is
// distinguishable in e2e.
import { writeDocx } from "./docx/writer";
import { writeOdt } from "./odf/writer";
import { parseDocx, type ParsedDocxDocument } from "./docx/reader";
import { parseOdt, type ParsedOdtDocument } from "./odf/reader";
import { createDocument } from "./documentModel";
import type { DocTheme, Block, DocumentModel, Paragraph, Table } from "./documentModel";
import type { ThemeColorRef } from "./core/theme";

/** Pinned so a regenerated fixture's dcterms:modified / dc:date does not
 * become whichever day someone re-ran the generator — see SPEC-008's
 * round-trip determinism rule (the same reason fixtures.borders.ts pins one). */
export const FIXED_MODIFIED = new Date(Date.UTC(2026, 8, 3, 9, 0, 0));

/** The fixture document's own theme: Office's structure with deliberately
 * distinctive accents, so a resolved colour can only have come from here. */
export const FIXTURE_THEME: DocTheme = {
  name: "Spec 77",
  colors: {
    dk1: "#1a1a1a",
    lt1: "#ffffff",
    dk2: "#203040",
    lt2: "#eef1f4",
    accent1: "#2f8f5b",
    accent2: "#b3541e",
    accent3: "#7a5195",
    accent4: "#1f6fb2",
    accent5: "#c2a336",
    accent6: "#6a8d3f",
    hlink: "#0563c1",
    folHlink: "#954f72",
  },
  fonts: { major: "Georgia", minor: "Arial" },
};

/** The heading run's colour: an untransformed accent reference, which both
 * formats round-trip as a reference (AC-52). */
export const HEADING_COLOR_REF: ThemeColorRef = { theme: "accent2" };

/** The table's cell shading: accent1 at tint 0.2 (keep 20% of the colour —
 * "Lighter 80%") — a TRANSFORMED reference in the tint form, which is the
 * one WordprocessingML can spell: `.docx` round-trips it via
 * w:themeFillTint, while a lumMod/lumOff ref would write its resolved
 * literal alone (D8/AC-32). `.odt` cannot carry any transform and reopens
 * the resolved literal, AC-39's first sanctioned delta. */
export const TABLE_SHADING_REF: ThemeColorRef = { theme: "accent1", tint: 0.2 };

/** The shape's solid fill: an untransformed accent reference. */
export const SHAPE_FILL_REF: ThemeColorRef = { theme: "accent4" };

/** The explicitly-coloured run — the colour a theme change must never move
 * (AC-16/AC-52). */
export const EXPLICIT_COLOR = "#c00000";

export const HEADING_TEXT = "Quarterly review";
export const EXPLICIT_TEXT = "This run is explicitly dark red and stays put.";

function paragraph(text: string, runFormatting = {}, headingLevel?: number): Paragraph {
  const block: Paragraph = { type: "paragraph", runs: [{ text, formatting: runFormatting }] };
  if (headingLevel !== undefined) {
    block.headingLevel = headingLevel;
  }
  return block;
}

function themedTable(): Table {
  const cell = (text: string, shaded: boolean) => ({
    blocks: [paragraph(text)],
    ...(shaded ? { shadingFill: TABLE_SHADING_REF } : {}),
  });
  return {
    type: "table",
    columnWidthsCm: [4, 4],
    rows: [
      { cells: [cell("Region", true), cell("Total", true)], isHeader: true },
      { cells: [cell("North", false), cell("1,240", false)] },
    ],
  };
}

/** The fixture's blocks, in the order the AC lists them. */
export function themeFixtureBlocks(): Block[] {
  return [
    // A theme-coloured heading: accent2 text in the major (heading) font,
    // both as references.
    paragraph(HEADING_TEXT, { color: HEADING_COLOR_REF, fontFamily: { theme: "major" } }, 1),
    paragraph("Body copy between the themed pieces, unformatted."),
    // A theme-tinted table.
    themedTable(),
    // A theme-filled shape.
    {
      type: "shape",
      geometry: "rect",
      widthCm: 5,
      heightCm: 3,
      wrap: "inline",
      fill: { kind: "solid", color: SHAPE_FILL_REF },
      outline: { color: "#404040", widthPt: 1, dash: "solid" },
      altText: "Themed rectangle",
    },
    // The explicitly-coloured run.
    paragraph(EXPLICIT_TEXT, { color: EXPLICIT_COLOR }),
  ];
}

export function themeFixtureModel(title: string): DocumentModel {
  const model = createDocument("spec077-theme");
  return {
    ...model,
    title,
    theme: { ...FIXTURE_THEME },
    sections: [{ blocks: themeFixtureBlocks() }],
  };
}

export const DOCX_FIXTURE_TITLE = "Word themes";
export const ODT_FIXTURE_TITLE = "Writer themes";

/** The model a same-format save starts from after an open — the shape
 * fixtures.theme.spec.ts's modelFrom and openDocument build: parsed content
 * plus the retained source package. */
function reopenedModel(parsed: ParsedDocxDocument | ParsedOdtDocument, title: string, format: "docx" | "odt"): DocumentModel {
  const model: DocumentModel = {
    id: "spec077-theme",
    title,
    sections: parsed.sections,
    defaultStyles: parsed.defaultStyles,
    format,
    opaque: parsed.opaque,
    sourceXml: parsed.sourceXml,
  };
  if (parsed.theme) {
    model.theme = parsed.theme;
  }
  return model;
}

export async function buildThemeFixtures(): Promise<{ name: string; bytes: Uint8Array }[]> {
  // Each fixture is written, reopened and written again, and the SECOND
  // write's bytes are committed: both readers materialise the document
  // defaults (font family/size) as explicit run formatting, so the first
  // write — whose model carries bare runs — is not a fixed point of
  // open → save, and AC-53's no-edit byte-identity holds only of the
  // stabilised bytes.
  const docxFirst = await writeDocx(themeFixtureModel(DOCX_FIXTURE_TITLE), { modified: FIXED_MODIFIED });
  const docx = await writeDocx(
    reopenedModel(await parseDocx(docxFirst.bytes), DOCX_FIXTURE_TITLE, "docx"),
    { modified: FIXED_MODIFIED },
  );
  const odtFirst = await writeOdt(themeFixtureModel(ODT_FIXTURE_TITLE), { modified: FIXED_MODIFIED });
  const odt = await writeOdt(
    reopenedModel(await parseOdt(odtFirst.bytes), ODT_FIXTURE_TITLE, "odt"),
    { modified: FIXED_MODIFIED },
  );
  return [
    { name: "word-theme.docx", bytes: docx.bytes },
    { name: "lo-theme.odt", bytes: odt.bytes },
  ];
}
