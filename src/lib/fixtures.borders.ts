// The SPEC-074 borders/shading fixture model (AC-33), shared by the
// generator that writes the committed packages
// (scripts/build-spec074-fixtures.mjs) and the unit spec that guards them
// (fixtures.borders.spec.ts). It builds bytes and nothing else — no
// filesystem access, no side effects on import — so importing it can never
// rewrite a fixture the e2e suite may be reading.
//
// The fixtures carry exactly what the AC enumerates: a boxed note paragraph,
// a bottom-rule heading, three consecutive identical bordered paragraphs, a
// shaded code block, and a page border (with a page background beside it,
// since both live on the section).
import { writeDocx } from "./docx/writer";
import { writeOdt } from "./odf/writer";
import { createDocument } from "./documentModel";
// BorderSide's home is core/borders (AC-1), which is where every other
// module takes it from — documentModel only consumes it.
import type { BorderSide } from "./core/borders";
import type { Block, DocumentModel, PageBorders, Paragraph, ParagraphBorders, ParagraphFormatting } from "./documentModel";

/** Pinned so a regenerated fixture's dcterms:modified / dc:date does not
 * become whichever day someone re-ran the generator — see SPEC-008's
 * round-trip determinism rule (the same reason fixtures.fields.ts pins one). */
export const FIXED_MODIFIED = new Date(Date.UTC(2026, 8, 2, 9, 0, 0));

const NOTE_RULE: BorderSide = { style: "single", widthPt: 1, color: "#7f7f7f" };
const HEADING_RULE: BorderSide = { style: "double", widthPt: 3, color: "#c00000" };
const RUN_RULE: BorderSide = { style: "single", widthPt: 0.5, color: "#000000" };
const RUN_BETWEEN: BorderSide = { style: "dotted", widthPt: 0.5, color: "#7f7f7f" };
const PAGE_RULE: BorderSide = { style: "double", widthPt: 3, color: "#1f4e79" };

/** The three consecutive paragraphs' shared borders — structurally equal on
 * every one of them, which is what makes them merge into a single bordered
 * block (AC-12). */
export const MERGED_RUN_BORDERS: ParagraphBorders = {
  top: RUN_RULE,
  bottom: RUN_RULE,
  left: RUN_RULE,
  right: RUN_RULE,
  between: RUN_BETWEEN,
  paddingPt: { top: 6, bottom: 6, left: 8, right: 8 },
};

/** The section's page border and background (AC-19..AC-24). */
export const FIXTURE_PAGE_BORDERS: PageBorders = {
  top: PAGE_RULE,
  bottom: PAGE_RULE,
  left: PAGE_RULE,
  right: PAGE_RULE,
  offsetPt: 24,
  measureFrom: "page",
};

export const FIXTURE_PAGE_BACKGROUND = "#fdf6e3";

/** The boxed note's shading (AC "a boxed note paragraph"). */
export const NOTE_SHADING = "#fff2cc";

/** The shaded code block's fill (AC "a shaded code block"). */
export const CODE_SHADING = "#f2f2f2";

/** The three merged paragraphs' text, in order — the spec's assertions and
 * the e2e suite both address them by content. */
export const MERGED_TEXTS = [
  "First of three identically bordered paragraphs.",
  "Second of three — the outer box draws once around all three.",
  "Third of three; only this one draws the bottom rule.",
];

function paragraph(text: string, formatting?: ParagraphFormatting, headingLevel?: number): Paragraph {
  const block: Paragraph = { type: "paragraph", runs: [{ text, formatting: {} }] };
  if (formatting) {
    block.formatting = formatting;
  }
  if (headingLevel !== undefined) {
    block.headingLevel = headingLevel;
  }
  return block;
}

/** The fixture's blocks, in the order the AC lists them. */
export function borderFixtureBlocks(): Block[] {
  return [
    // A bottom-rule heading — what Word's "---" AutoFormat produces, and the
    // case that opens with the rule missing without this spec.
    paragraph("Borders and shading", { borders: { bottom: HEADING_RULE, paddingPt: { bottom: 2 } } }, 1),
    paragraph("An ordinary paragraph with no borders at all, so the run below is clearly separate."),
    // A boxed note paragraph: all four edges plus a fill.
    paragraph("Note: this paragraph is boxed and shaded — the most common piece of document design there is.", {
      borders: { top: NOTE_RULE, bottom: NOTE_RULE, left: NOTE_RULE, right: NOTE_RULE },
      shadingFill: NOTE_SHADING,
    }),
    paragraph("Another unbordered paragraph, so the box above cannot merge with the run below."),
    // Three consecutive identical bordered paragraphs (AC-12).
    ...MERGED_TEXTS.map((text) => paragraph(text, { borders: MERGED_RUN_BORDERS })),
    paragraph("A closing paragraph, unbordered."),
    // A shaded code block: shading with no borders at all.
    paragraph("const answer = 42;", {
      shadingFill: CODE_SHADING,
      indentStartCm: 1,
    }),
  ];
}

export function borderFixtureModel(title: string): DocumentModel {
  const model = createDocument("spec074-borders");
  return {
    ...model,
    title,
    sections: [
      {
        blocks: borderFixtureBlocks(),
        pageBorders: FIXTURE_PAGE_BORDERS,
        pageBackground: FIXTURE_PAGE_BACKGROUND,
      },
    ],
  };
}

export async function buildBorderFixtures(): Promise<{ name: string; bytes: Uint8Array }[]> {
  const docx = await writeDocx(borderFixtureModel("Word borders"), { modified: FIXED_MODIFIED });
  const odt = await writeOdt(borderFixtureModel("Writer borders"), { modified: FIXED_MODIFIED });
  return [
    { name: "word-borders.docx", bytes: docx.bytes },
    { name: "lo-borders.odt", bytes: odt.bytes },
  ];
}
