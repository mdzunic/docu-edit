// SPEC-077 AC-52/AC-53: guards the committed theme fixtures
// `e2e/fixtures/word-theme.docx` and `e2e/fixtures/lo-theme.odt` — a
// theme-coloured heading, a theme-tinted table, a theme-filled shape and an
// explicitly-coloured run, over a document theme distinct from every
// built-in. Each opens with its references intact, survives a theme change
// with the explicit colour unmoved, saves and reopens `equalStructural`
// over content and `model.theme` — for `.odt` with the two AC-39 deltas
// (transformed reference → resolved literal, font reference → resolved
// family) as the only sanctioned differences — and a no-edit open → save
// keeps every package part byte-identical, the theme part included (AC-33).
//
// This spec is READ-ONLY on e2e/fixtures — generation lives in
// scripts/build-spec077-fixtures.mjs, so a unit run can never rewrite a
// file the e2e suite is uploading (the same reasoning as
// fixtures.borders.spec.ts).
import { promises as fs } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { beforeAll, describe, expect, it } from "vitest";
import { parseDocx, type ParsedDocxDocument } from "./docx/reader";
import { parseOdt, type ParsedOdtDocument } from "./odf/reader";
import { writeDocx } from "./docx/writer";
import { writeOdt } from "./odf/writer";
import { unzip } from "./core/zip";
import { equalBytes } from "./core/equality";
import { equalStructural, firstStructuralDifference } from "./equality";
import { BUILT_IN_THEMES, resolveColor } from "./core/theme";
import type { DocumentModel, Paragraph, Shape, Table } from "./documentModel";
import {
  buildThemeFixtures,
  EXPLICIT_COLOR,
  EXPLICIT_TEXT,
  FIXED_MODIFIED,
  FIXTURE_THEME,
  HEADING_COLOR_REF,
  HEADING_TEXT,
  SHAPE_FILL_REF,
  TABLE_SHADING_REF,
} from "./fixtures.theme";

const FIXTURE_DIR = path.join(path.dirname(fileURLToPath(import.meta.url)), "..", "..", "e2e", "fixtures");

const REGENERATE_HINT =
  "The committed fixture no longer matches src/lib/fixtures.theme.ts. " +
  "Regenerate it with `npx vite-node scripts/build-spec077-fixtures.mjs` and commit the result.";

let docxBytes: Uint8Array;
let odtBytes: Uint8Array;

beforeAll(async () => {
  docxBytes = new Uint8Array(await fs.readFile(path.join(FIXTURE_DIR, "word-theme.docx")));
  odtBytes = new Uint8Array(await fs.readFile(path.join(FIXTURE_DIR, "lo-theme.odt")));
});

/** The model an opened document actually holds (see openDocument.ts): the
 * parsed content plus the retained source package, so a same-format save
 * takes the preservation path rather than writing from scratch. */
function modelFrom(parsed: ParsedDocxDocument | ParsedOdtDocument, format: "docx" | "odt"): DocumentModel {
  const model: DocumentModel = {
    id: "spec077-theme",
    title: "Themes",
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

function paragraphWith(model: DocumentModel, text: string): Paragraph {
  const found = model.sections[0].blocks.find(
    (block) => block.type === "paragraph" && block.runs.map((run) => run.text).join("") === text,
  );
  if (!found || found.type !== "paragraph") {
    throw new Error(`no paragraph reading "${text}"`);
  }
  return found;
}

function tableOf(model: DocumentModel): Table {
  const found = model.sections[0].blocks.find((block) => block.type === "table");
  if (!found || found.type !== "table") {
    throw new Error("no table in the fixture");
  }
  return found;
}

function shapeOf(model: DocumentModel): Shape {
  const found = model.sections[0].blocks.find((block) => block.type === "shape");
  if (!found || found.type !== "shape") {
    throw new Error("no shape in the fixture");
  }
  return found;
}

function solidFillColor(shape: Shape) {
  if (shape.fill?.kind !== "solid") {
    throw new Error("the fixture shape is not solid-filled");
  }
  return shape.fill.color;
}

describe("the committed fixtures still match the model that generates them", () => {
  it("regenerating produces byte-identical packages", async () => {
    const built = await buildThemeFixtures();
    const docx = built.find((entry) => entry.name === "word-theme.docx");
    const odt = built.find((entry) => entry.name === "lo-theme.odt");
    expect(docx, REGENERATE_HINT).toBeDefined();
    expect(odt, REGENERATE_HINT).toBeDefined();
    expect(equalBytes(docx!.bytes, docxBytes), REGENERATE_HINT).toBe(true);
    expect(equalBytes(odt!.bytes, odtBytes), REGENERATE_HINT).toBe(true);
  });
});

describe("docx fixture (AC-52)", () => {
  async function open(): Promise<DocumentModel> {
    return modelFrom(await parseDocx(docxBytes), "docx");
  }

  it("opens with every reference intact and the document's own theme", async () => {
    const model = await open();
    expect(paragraphWith(model, HEADING_TEXT).runs[0].formatting.color).toEqual(HEADING_COLOR_REF);
    expect(paragraphWith(model, HEADING_TEXT).runs[0].formatting.fontFamily).toEqual({ theme: "major" });
    expect(tableOf(model).rows[0].cells[0].shadingFill).toEqual(TABLE_SHADING_REF);
    expect(solidFillColor(shapeOf(model))).toEqual(SHAPE_FILL_REF);
    expect(paragraphWith(model, EXPLICIT_TEXT).runs[0].formatting.color).toBe(EXPLICIT_COLOR);
    expect(model.theme?.name).toBe(FIXTURE_THEME.name);
    expect(model.theme?.colors).toEqual(FIXTURE_THEME.colors);
    expect(model.theme?.fonts).toEqual(FIXTURE_THEME.fonts);
    expect(model.theme?.partName).toBeDefined();
    expect(model.theme?.dirty).toBeUndefined();
  });

  it("survives a theme change with references and the explicit colour unmoved (AC-16/18)", async () => {
    const model = await open();
    const themed: DocumentModel = { ...model, theme: { ...BUILT_IN_THEMES[1], partName: model.theme?.partName, dirty: true } };
    // D17: the only model change is model.theme — every reference and every
    // literal is exactly what it was.
    expect(equalStructural(themed.sections, model.sections)).toBe(true);
    expect(paragraphWith(themed, EXPLICIT_TEXT).runs[0].formatting.color).toBe(EXPLICIT_COLOR);
    // What DID change is resolution: the heading now resolves against Facet.
    expect(resolveColor(HEADING_COLOR_REF, themed.theme!)).not.toBe(resolveColor(HEADING_COLOR_REF, FIXTURE_THEME));
  });

  it("saves after a theme change and reopens equalStructural over content and theme", async () => {
    const model = await open();
    const themed: DocumentModel = { ...model, theme: { ...BUILT_IN_THEMES[1], partName: model.theme?.partName, dirty: true } };
    const written = await writeDocx(themed, { modified: FIXED_MODIFIED });
    const reopened = modelFrom(await parseDocx(written.bytes), "docx");
    expect(
      equalStructural(reopened.sections, themed.sections),
      firstStructuralDifference(reopened.sections, themed.sections) ?? "",
    ).toBe(true);
    expect(reopened.theme?.name).toBe(BUILT_IN_THEMES[1].name);
    expect(reopened.theme?.colors).toEqual(BUILT_IN_THEMES[1].colors);
    expect(reopened.theme?.fonts).toEqual(BUILT_IN_THEMES[1].fonts);
    expect(reopened.theme?.dirty).toBeUndefined();
  });

  it("keeps every package part byte-identical on a no-edit open → save, the theme part included (AC-53/AC-33)", async () => {
    const model = await open();
    const written = await writeDocx(model, { modified: FIXED_MODIFIED });
    const before = await unzip(docxBytes);
    const after = await unzip(written.bytes);
    expect([...after.keys()].sort()).toEqual([...before.keys()].sort());
    for (const [name, bytes] of before) {
      expect(equalBytes(after.get(name)!, bytes), `part ${name} changed on a no-edit save`).toBe(true);
    }
  });
});

describe("odt fixture (AC-52, with the two AC-39 deltas pinned)", () => {
  async function open(): Promise<DocumentModel> {
    return modelFrom(await parseOdt(odtBytes), "odt");
  }

  it("opens with the untransformed references intact and the theme recovered (AC-40)", async () => {
    const model = await open();
    // The untransformed references survive as references (AC-38).
    expect(paragraphWith(model, HEADING_TEXT).runs[0].formatting.color).toEqual(HEADING_COLOR_REF);
    expect(solidFillColor(shapeOf(model))).toEqual(SHAPE_FILL_REF);
    // Delta 1 (AC-39): the transformed table tint reopens as its resolved
    // literal — never silently anything else.
    expect(tableOf(model).rows[0].cells[0].shadingFill).toBe(resolveColor(TABLE_SHADING_REF, FIXTURE_THEME));
    // Delta 2 (AC-39): the font reference reopens as the resolved family.
    expect(paragraphWith(model, HEADING_TEXT).runs[0].formatting.fontFamily).toBe(FIXTURE_THEME.fonts.major);
    expect(paragraphWith(model, EXPLICIT_TEXT).runs[0].formatting.color).toBe(EXPLICIT_COLOR);
    expect(model.theme?.name).toBe(FIXTURE_THEME.name);
    expect(model.theme?.colors).toEqual(FIXTURE_THEME.colors);
    expect(model.theme?.fonts).toEqual(FIXTURE_THEME.fonts);
  });

  it("survives a theme change and a save → reopen with only the sanctioned deltas", async () => {
    const model = await open();
    const themed: DocumentModel = { ...model, theme: { ...BUILT_IN_THEMES[1], dirty: true } };
    expect(paragraphWith(themed, EXPLICIT_TEXT).runs[0].formatting.color).toBe(EXPLICIT_COLOR);
    const written = await writeOdt(themed, { modified: FIXED_MODIFIED });
    const reopened = modelFrom(await parseOdt(written.bytes), "odt");
    // The untransformed reference still tracks the (new) theme.
    expect(paragraphWith(reopened, HEADING_TEXT).runs[0].formatting.color).toEqual(HEADING_COLOR_REF);
    expect(solidFillColor(shapeOf(reopened))).toEqual(SHAPE_FILL_REF);
    expect(paragraphWith(reopened, EXPLICIT_TEXT).runs[0].formatting.color).toBe(EXPLICIT_COLOR);
    expect(reopened.theme?.name).toBe(BUILT_IN_THEMES[1].name);
    expect(reopened.theme?.colors).toEqual(BUILT_IN_THEMES[1].colors);
  });

  it("keeps every package part byte-identical on a no-edit open → save (AC-53)", async () => {
    const model = await open();
    const written = await writeOdt(model, { modified: FIXED_MODIFIED });
    const before = await unzip(odtBytes);
    const after = await unzip(written.bytes);
    expect([...after.keys()].sort()).toEqual([...before.keys()].sort());
    for (const [name, bytes] of before) {
      expect(equalBytes(after.get(name)!, bytes), `part ${name} changed on a no-edit save`).toBe(true);
    }
  });
});
