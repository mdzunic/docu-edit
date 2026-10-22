// SPEC-050 — round-trip preservation for Markdown: AC-7 (.md -> .md,
// unedited, equalStructural), AC-8 (.docx -> .md -> .docx, equalSemantic
// against the source masked by the writer's own `dropped`) and AC-12 (a .md
// fixture exercising every supported construct, saved as .docx, reopened,
// equalSemantic against the masked source).
import { describe, expect, it } from "vitest";
import { readFileSync } from "node:fs";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";
import { parseMd } from "./reader";
import { writeMd } from "./writer";
import { parseDocx } from "../docx/reader";
import { writeDocx } from "../docx/writer";
import { assembleDocumentModel } from "../openDocument";
import { collectDropRecords, withoutDropped } from "../opaque";
import { equalSemantic, equalStructural, firstStructuralDifference, resolve } from "../equality";
import { builtInStyles } from "../styles";
import type { DocumentModel } from "../documentModel";

const fixturesDir = join(dirname(fileURLToPath(import.meta.url)), "fixtures");

function loadFixture(name: string): Uint8Array {
  return new Uint8Array(readFileSync(join(fixturesDir, name)));
}

async function openMdFixture(name: string): Promise<DocumentModel> {
  const parsed = await parseMd(loadFixture(name));
  return assembleDocumentModel("Kitchen Sink", "md", parsed);
}

/** Package-level fields (SPEC-008's opaque parts/verbatim XML, document
 * properties) are format-plumbing, not content — comparing them would mean
 * comparing freshly generated docx infrastructure (a new docProps/app.xml,
 * a new settings.xml, …) against nothing, which always differs and says
 * nothing about whether the document's *content* survived. The existing
 * .docx<->.odt round-trip tests (docx/roundTrip.spec.ts) sidestep the same
 * issue by comparing `.sections` alone; this compares the fuller model
 * (styles, media, list definitions, …) but drops the same package-plumbing
 * fields both sides would otherwise spuriously differ on. */
function round2(cm: number): number {
  return Math.round(cm * 100) / 100;
}

/** Verbatim preserved-XML slots (SPEC-008's `docxTblPr`/`docxTrPr`/
 * `docxTcPr`/`odtStyleName`) are the same kind of package-plumbing as
 * `opaque`/`sourceXml`, just scoped to one table/row/cell instead of the
 * whole document: any table a *real* .docx save round-trips picks one up
 * (the writer's own default table markup carries properties the reader
 * does not model, e.g. its look flags), regardless of whether the model
 * that produced it ever went through .docx before. A table built directly
 * from Markdown (never touched .docx) has none at all, so comparing them
 * literally would fail on package plumbing, not content — stripped here
 * for the same reason `contentOnly` strips `opaque`/`sourceXml`. */
function stripVerbatimTableXml(blocks: DocumentModel["sections"][number]["blocks"]): void {
  for (const block of blocks) {
    if (block.type === "table") {
      delete block.docxTblPr;
      delete block.odtStyleName;
      delete block.odtColumnStyleNames;
      // "fixed" is the documented default for an absent widthMode (Table's
      // own doc comment) — a real .docx reopen writes it explicitly, a
      // model that never went through .docx leaves it absent; the same
      // meaning either way.
      if (block.widthMode === "fixed") {
        delete block.widthMode;
      }
      for (const row of block.rows) {
        delete row.docxTrPr;
        delete row.odtStyleName;
        for (const cell of row.cells) {
          delete cell.docxTcPr;
          delete cell.odtStyleName;
          stripVerbatimTableXml(cell.blocks);
        }
      }
    }
  }
}

function contentOnly(model: DocumentModel): DocumentModel {
  const {
    opaque: _opaque,
    sourceXml: _sourceXml,
    properties: _properties,
    format: _format,
    title: _title,
    // Front matter is Markdown-only (never interpreted — SPEC-050's "Front-
    // matter interpretation... is carried, not read"); a .docx/.odt target
    // has nowhere to keep it and this spec does not ask for a drop report
    // in that direction (only *saving as* .md is in scope for `dropped`),
    // so it is excluded here rather than asserted lost with no mechanism
    // to mask it by.
    frontMatter: _frontMatter,
    // A document's own style *catalogue* (as opposed to what each style
    // resolves to, which `resolve()` already merges onto every paragraph/
    // run before this runs) is not portable through Markdown at all — only
    // a construct's *kind* survives (this is a heading, this is a quote),
    // never a source document's own custom definition of what that style
    // looks like. That is a real, separate loss this spec's supported
    // subset does not cover reporting (SPEC-050 depends on but does not
    // amend SPEC-019/SPEC-042's style-catalogue behaviour), so it is
    // excluded from this comparison rather than asserted lost with no
    // drop kind to mask it by.
    styles: _styles,
    ...rest
  } = structuredClone(model);
  for (const section of rest.sections) {
    stripVerbatimTableXml(section.blocks);
  }
  return {
    ...rest,
    // .docx stores every length in twips (1/1440in); converting a clean cm
    // value there and back introduces a sub-hundredth-of-a-cm rounding
    // difference that has nothing to do with content fidelity — rounded
    // away here the same way lib/tables.ts's own snapCm acknowledges the
    // same twip grid elsewhere.
    defaultStyles: {
      ...rest.defaultStyles,
      pageSize: {
        widthCm: round2(rest.defaultStyles.pageSize.widthCm),
        heightCm: round2(rest.defaultStyles.pageSize.heightCm),
      },
      margins: {
        topCm: round2(rest.defaultStyles.margins.topCm),
        rightCm: round2(rest.defaultStyles.margins.rightCm),
        bottomCm: round2(rest.defaultStyles.margins.bottomCm),
        leftCm: round2(rest.defaultStyles.margins.leftCm),
      },
    },
  } as DocumentModel;
}

/** `Image.mediaId`/`MediaPart.mediaId` are generated identifiers (a fresh
 * one on every parse — see MediaPart's own doc comment on why the key is
 * deliberately not named `id`, which would make equality silently ignore
 * it), never meant to compare equal across two independent opens of
 * unrelated bytes. What must hold is *referential* integrity — image N
 * refers to media N — so both sides are renumbered to positional
 * "media-0", "media-1", … canonical ids in document order before
 * comparing, the same way a diff tool renames variables to compare
 * structure, not spelling. */
function canonicalizeMediaIds(model: DocumentModel): DocumentModel {
  const clone = structuredClone(model);
  const canonical = new Map<string, string>();
  const nextId = (mediaId: string): string => {
    let id = canonical.get(mediaId);
    if (!id) {
      id = `media-${canonical.size}`;
      canonical.set(mediaId, id);
    }
    return id;
  };
  const visitBlocks = (blocks: DocumentModel["sections"][number]["blocks"]): void => {
    for (const block of blocks) {
      if (block.type === "image") {
        block.mediaId = nextId(block.mediaId);
      } else if (block.type === "table") {
        for (const row of block.rows) {
          for (const cell of row.cells) {
            visitBlocks(cell.blocks);
          }
        }
      }
    }
  };
  for (const section of clone.sections) {
    visitBlocks(section.blocks);
  }
  clone.media = clone.media?.map((part) => ({ ...part, mediaId: canonical.get(part.mediaId) ?? part.mediaId }));
  return clone;
}

describe("SPEC-050 — AC-7: .md -> .md structural round trip", () => {
  it("re-opens equalStructural after an unedited save (no equalBytes — the serialiser normalises syntax)", async () => {
    const opened = await openMdFixture("kitchen-sink.md");
    const written = await writeMd(opened);
    expect(written.dropped).toEqual([]);
    const reopened = await parseMd(written.bytes);
    const diff = firstStructuralDifference(opened.sections, reopened.sections);
    expect(diff).toBeNull();
    expect(equalStructural(opened.sections, reopened.sections)).toBe(true);
    expect(opened.frontMatter).toBe(reopened.frontMatter);
  });

  it("normalises *italic* and a single-asterisk source to the canonical _italic_ spelling", async () => {
    const opened = await openMdFixture("kitchen-sink.md");
    const written = await writeMd(opened);
    const text = new TextDecoder().decode(written.bytes);
    expect(text).toContain("_italic_");
    expect(text).not.toContain("*italic*");
  });
});

describe("SPEC-050 — AC-8: .docx -> .md -> .docx semantic round trip", () => {
  it("survives a docx -> md -> docx conversion modulo the writer's own dropped list", async () => {
    // Built directly as a model (rather than loaded from a third-party
    // fixture) so its styles are exactly this app's own built-ins — a real
    // Word file's own custom "Heading 1" definition would diverge from this
    // app's, which is a real, separate loss (the document's own style
    // *definitions* are not portable through Markdown) but not one this
    // test is about; see the implementation summary for that scope note.
    const sourceModel: DocumentModel = {
      id: "src",
      title: "Quarterly Report",
      defaultStyles: {
        pageSize: { widthCm: 21, heightCm: 29.7 },
        margins: { topCm: 2.54, rightCm: 2.54, bottomCm: 2.54, leftCm: 2.54 },
        fontFamily: "Calibri",
        fontSizePt: 11,
      },
      // Written explicitly (docx/writer.ts otherwise emits a minimal
      // Word-compatible Heading 1 definition — outline level and a
      // basedOn chain only, no baked bold/size of its own, since real
      // Word documents commonly rely on the application's built-in
      // template rather than a fully-specified style) so this document's
      // own style catalogue round-trips with the *same* effective
      // formatting this app's built-in Heading1 already carries — the
      // point of this test is content fidelity through Markdown, not
      // .docx's own minimal-style-definition convention.
      styles: builtInStyles(),
      sections: [
        {
          // Diverges from defaultStyles.margins.topCm above — a section
          // that merely repeats the document default is written back with
          // no override at all (SPEC-028), which would defeat the point of
          // this fixture (it needs a *real* page-setup override to survive
          // the .docx round trip and reach collectMarkdownDropRecords).
          margins: { topCm: 5, rightCm: 2.54, bottomCm: 2.54, leftCm: 2.54 },
          headers: { default: { blocks: [{ type: "paragraph", runs: [{ text: "Confidential", formatting: {} }] }] } },
          blocks: [
            {
              type: "paragraph",
              runs: [{ text: "Quarterly Report", formatting: { bold: true, fontSizePt: 20 } }],
              styleId: "Heading1",
              headingLevel: 1,
            },
            {
              type: "paragraph",
              runs: [
                { text: "This is ", formatting: {} },
                { text: "bold", formatting: { bold: true } },
                { text: " and ", formatting: {} },
                { text: "italic", formatting: { italic: true } },
                { text: " text with a ", formatting: {} },
                { text: "link", formatting: { linkHref: "https://example.com" } },
                { text: ".", formatting: {} },
              ],
            },
            { type: "paragraph", runs: [{ text: "item one", formatting: {} }], list: { type: "bullet", level: 1 } },
            { type: "paragraph", runs: [{ text: "item two", formatting: {} }], list: { type: "bullet", level: 1 } },
            {
              type: "table",
              rows: [
                {
                  isHeader: true,
                  cells: [
                    { blocks: [{ type: "paragraph", runs: [{ text: "A", formatting: {} }] }] },
                    { blocks: [{ type: "paragraph", runs: [{ text: "B", formatting: {} }] }] },
                  ],
                },
                {
                  cells: [
                    { blocks: [{ type: "paragraph", runs: [{ text: "1", formatting: {} }] }] },
                    { blocks: [{ type: "paragraph", runs: [{ text: "2", formatting: {} }] }] },
                  ],
                },
              ],
              columnWidthsCm: [8, 8],
            },
          ],
        },
      ],
    };
    const { bytes: sourceBytes } = await writeDocx(sourceModel);
    const model = assembleDocumentModel("Quarterly Report", "docx", await parseDocx(sourceBytes));

    const written = await writeMd(model);
    // A page-margin override and a header both have no Markdown
    // representation (SPEC-050's "What markdown cannot carry").
    expect(written.dropped.some((d) => d.kind === "pageSetup")).toBe(true);
    expect(written.dropped.some((d) => d.kind === "headerFooter")).toBe(true);

    const reopenedMdModel = assembleDocumentModel("Quarterly Report", "md", await parseMd(written.bytes));
    const { bytes: backToDocxBytes } = await writeDocx(reopenedMdModel);
    const reopened = assembleDocumentModel("Quarterly Report", "docx", await parseDocx(backToDocxBytes));

    const masked = canonicalizeMediaIds(withoutDropped(model, written.dropped));
    const reopenedCanonical = canonicalizeMediaIds(reopened);
    const diff = firstStructuralDifference(resolve(contentOnly(masked)), resolve(contentOnly(reopenedCanonical)));
    expect(diff).toBeNull();
    expect(equalSemantic(contentOnly(masked), contentOnly(reopenedCanonical))).toBe(true);
  });
});

describe("SPEC-050 — AC-12: a .md fixture exercising the supported subset, saved as .docx, reopened", () => {
  it("survives a md -> docx -> reopen conversion modulo the writer's own dropped list", async () => {
    const source = await openMdFixture("kitchen-sink.md");
    const dropped = collectDropRecords(source, "docx");
    // Only the raw HTML block and the reference-style link definition
    // (both preserved as opaque blocks, SPEC-008) are outside what .docx
    // can carry — everything else in the kitchen sink is in the supported
    // subset and survives.
    expect(dropped.every((d) => d.kind === "opaqueBlock")).toBe(true);

    const { bytes } = await writeDocx(source);
    expect(bytes.length).toBeGreaterThan(0);
    const reopened = assembleDocumentModel("Kitchen Sink", "docx", await parseDocx(bytes));

    const masked = canonicalizeMediaIds(withoutDropped(source, dropped));
    const reopenedCanonical = canonicalizeMediaIds(reopened);
    const diff = firstStructuralDifference(resolve(contentOnly(masked)), resolve(contentOnly(reopenedCanonical)));
    expect(diff).toBeNull();
    expect(equalSemantic(contentOnly(masked), contentOnly(reopenedCanonical))).toBe(true);
  });
});
