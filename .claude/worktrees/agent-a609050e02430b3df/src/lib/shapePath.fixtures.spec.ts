// SPEC-071 AC-4/AC-5/AC-26..AC-34 — the freeform fixtures. word-freeform.docx
// and lo-freeform.odt are hand-built by scripts/build-spec071-fixtures.mjs to
// reproduce what Word and Writer emit for hand-drawn geometry — NOT
// round-tripped through this project's own writer — so opening them exercises
// the real custom-geometry readers. Both hold the same eight shapes in the
// same document order:
//   0 open polyline   2 closed curve    4 smooth/corner mix  6 formula-driven
//   1 closed polygon  3 open scribble   5 reduced primitives 7 unsupported
// 0..5 model as a `ShapePath` and paint their real outline; 6 and 7 are what
// the narrowed placeholder scope still preserves verbatim.
import { describe, it, expect } from "vitest";
import { readFileSync } from "node:fs";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";
import { parseDocx, type ParsedDocxDocument } from "./docx/reader";
import { parseOdt, type ParsedOdtDocument } from "./odf/reader";
import { writeDocx } from "./docx/writer";
import { writeOdt } from "./odf/writer";
import { unzip } from "./core/zip";
import { firstStructuralDifference } from "./equality";
import { PATH_GRID } from "./core/shapePath";
import type { DocumentModel, Shape } from "./documentModel";

const here = dirname(fileURLToPath(import.meta.url));
const DOCX_FIXTURE = join(here, "docx", "fixtures", "word-freeform.docx");
const ODT_FIXTURE = join(here, "odf", "fixtures", "lo-freeform.odt");

function loadFixture(path: string): Uint8Array {
  return new Uint8Array(readFileSync(path));
}

/** The model an opened document actually holds (see openDocument.ts): the
 * parsed sections plus the retained source package, so a same-format save
 * takes the preservation path rather than writing from scratch. */
function modelFrom(parsed: ParsedDocxDocument | ParsedOdtDocument, format: "docx" | "odt"): DocumentModel {
  return {
    id: "freeform-fixture",
    title: "Freeform fixture",
    sections: parsed.sections,
    defaultStyles: parsed.defaultStyles,
    format,
    opaque: parsed.opaque,
    sourceXml: parsed.sourceXml,
  };
}

function shapesOf(parsed: ParsedDocxDocument | ParsedOdtDocument): Shape[] {
  return parsed.sections[0].blocks.filter((block) => block.type === "shape") as Shape[];
}

const decoder = new TextDecoder();

async function partText(bytes: Uint8Array, name: string): Promise<string> {
  const parts = await unzip(bytes);
  return decoder.decode(parts.get(name)!);
}

/** Every step op in a path, flattened across its subpaths — the model's whole
 * primitive vocabulary, so a test can assert "cubics only, no arcs left". */
function opsOf(shape: Shape): string[] {
  return shape.path!.subpaths.flatMap((subpath) => subpath.steps.map((step) => step.op));
}

// The alt text scripts/build-spec071-fixtures.mjs gives each shape, in
// document order — the fixtures' index contract, shared with e2e/SPEC-071.
const LABELS = [
  "Word freeform",
  "Closed freeform",
  "Closed curve",
  "Open scribble",
  "Smooth and corner mix",
  "Reduced primitives",
  "Formula-driven geometry",
  "Unsupported geometry",
];

/** The reader assertions both fixtures must satisfy — the two formats carry
 * the same eight shapes, so this runs once per format instead of duplicating
 * every expectation. */
function registerReaderTests(
  formatLabel: string,
  acTag: string,
  parse: () => Promise<ParsedDocxDocument | ParsedOdtDocument>,
): void {
  describe(`SPEC-071 — opening ${formatLabel} (${acTag})`, () => {
    it("reads all eight shapes in document order with their alt text", async () => {
      const shapes = shapesOf(await parse());
      expect(shapes).toHaveLength(8);
      expect(shapes.map((shape) => shape.altText)).toEqual(LABELS);
    });

    it(`parses the six path-shaped geometries into a real outline (${acTag})`, async () => {
      const shapes = shapesOf(await parse());
      for (let i = 0; i < 6; i++) {
        expect(shapes[i].path, LABELS[i]).toBeTruthy();
        expect(shapes[i].unsupportedGeometryXml, LABELS[i]).toBeUndefined();
        // `geometry` stays the bounding-box stand-in the model's convention
        // requires; `path` is what renders (AC-2).
        expect(shapes[i].geometry, LABELS[i]).toBe("rect");
      }
    });

    it("distinguishes open outlines from closed, fillable ones (AC-9)", async () => {
      const shapes = shapesOf(await parse());
      const closed = shapes.slice(0, 6).map((shape) => shape.path!.subpaths.every((sub) => sub.closed));
      expect(closed).toEqual([false, true, true, false, true, true]);
      // The two open ones carry no fill, so they paint as outline only.
      expect(shapes[0].fill).toBeUndefined();
      expect(shapes[3].fill).toBeUndefined();
      expect(shapes[1].fill).toBeTruthy();
    });

    it("carries only move/line/cubic/close — arcs and quadratics are reduced on read (AC-29)", async () => {
      const shapes = shapesOf(await parse());
      const modelled = new Set(["move", "line", "cubic", "close"]);
      for (let i = 0; i < 6; i++) {
        for (const op of opsOf(shapes[i])) {
          expect(modelled.has(op), `${LABELS[i]}: ${op}`).toBe(true);
        }
      }
      // Shape 5's source is a quadratic followed by a quarter-turn arc; both
      // arrive as cubics, the arc as more than one (it is subdivided).
      const reduced = opsOf(shapes[5]);
      expect(reduced[0]).toBe("move");
      expect(reduced.filter((op) => op === "cubic").length).toBeGreaterThan(1);
      expect(reduced.at(-1)).toBe("close");
    });

    it("keeps the reduced shape's source geometry for a lossless same-format save (AC-30)", async () => {
      const shapes = shapesOf(await parse());
      expect(shapes[5].pathSource).toBeTruthy();
      // The five shapes whose source was already move/line/cubic/close lost
      // nothing, so they carry no preserved source at all.
      for (let i = 0; i < 5; i++) {
        expect(shapes[i].pathSource, LABELS[i]).toBeUndefined();
      }
    });

    it("keeps a formula-driven geometry preserved rather than flattening it (AC-31)", async () => {
      const formula = shapesOf(await parse())[6];
      expect(formula.path).toBeUndefined();
      expect(formula.unsupportedGeometryXml).toBeTruthy();
    });

    it("keeps a genuinely unmodellable geometry on the placeholder path (AC-32)", async () => {
      const unsupported = shapesOf(await parse())[7];
      expect(unsupported.path).toBeUndefined();
      expect(unsupported.unsupportedGeometryXml).toBeTruthy();
    });

    it("never carries both a path and preserved geometry XML (AC-3)", async () => {
      for (const shape of shapesOf(await parse())) {
        expect(!!(shape.path && shape.unsupportedGeometryXml)).toBe(false);
      }
    });
  });
}

registerReaderTests("word-freeform.docx", "AC-4", async () => parseDocx(loadFixture(DOCX_FIXTURE)));
registerReaderTests("lo-freeform.odt", "AC-5", async () => parseOdt(loadFixture(ODT_FIXTURE)));

describe("SPEC-071 — the fixtures save and reopen unchanged (AC-34)", () => {
  it("word-freeform.docx opens, saves and reopens equalStructural", async () => {
    const opened = await parseDocx(loadFixture(DOCX_FIXTURE));
    const saved = await writeDocx(modelFrom(opened, "docx"));
    const reopened = await parseDocx(saved.bytes);
    const diff = firstStructuralDifference(opened.sections, reopened.sections);
    expect(diff, diff ?? "").toBeNull();
    // Still six real outlines and two preserved placeholders after the trip.
    const shapes = shapesOf(reopened);
    expect(shapes.filter((shape) => shape.path)).toHaveLength(6);
    expect(shapes.filter((shape) => shape.unsupportedGeometryXml)).toHaveLength(2);
  });

  it("lo-freeform.odt opens, saves and reopens equalStructural", async () => {
    const opened = await parseOdt(loadFixture(ODT_FIXTURE));
    const saved = await writeOdt(modelFrom(opened, "odt"));
    const reopened = await parseOdt(saved.bytes);
    const diff = firstStructuralDifference(opened.sections, reopened.sections);
    expect(diff, diff ?? "").toBeNull();
    const shapes = shapesOf(reopened);
    expect(shapes.filter((shape) => shape.path)).toHaveLength(6);
    expect(shapes.filter((shape) => shape.unsupportedGeometryXml)).toHaveLength(2);
  });

  it("writes the fixture's paths back as a:custGeom on the 21600 grid (AC-26)", async () => {
    const opened = await parseDocx(loadFixture(DOCX_FIXTURE));
    const saved = await writeDocx(modelFrom(opened, "docx"));
    const xml = await partText(saved.bytes, "word/document.xml");
    expect(xml).toContain(`<a:pathLst><a:path w="${PATH_GRID}" h="${PATH_GRID}">`);
    expect(xml).toContain("<a:moveTo>");
    expect(xml).toContain("<a:lnTo>");
    expect(xml).toContain("<a:cubicBezTo>");
    expect(xml).toContain("<a:close/>");
    // Seven custGeom elements in all: one per path shape (six, the reduced
    // one re-emitting its preserved source, which is itself a custGeom)...
    expect(xml.match(/<a:custGeom/g)).toHaveLength(7);
    // ...plus the preserved formula-driven geometry, still carrying its guide
    // reference rather than a flattened coordinate.
    expect(xml).toContain('<a:gd name="x1"');
    expect(xml).toContain('<a:pt x="x1" y="0"/>');
  });

  it("writes the fixture's paths back as draw:enhanced-path with a matching viewBox (AC-27)", async () => {
    const opened = await parseOdt(loadFixture(ODT_FIXTURE));
    const saved = await writeOdt(modelFrom(opened, "odt"));
    const xml = await partText(saved.bytes, "content.xml");
    expect(xml).toContain(`svg:viewBox="0 0 ${PATH_GRID} ${PATH_GRID}"`);
    expect(xml).toContain('draw:type="non-primitive"');
    expect(xml).toMatch(/draw:enhanced-path="M \d+ \d+ L /);
    expect(xml).toMatch(/draw:enhanced-path="[^"]*C \d+ \d+ \d+ \d+ \d+ \d+/);
    expect(xml).toMatch(/draw:enhanced-path="[^"]* Z N"/);
    // The preserved formula-driven path keeps its equation and its ?f0
    // reference instead of being written out as numbers.
    expect(xml).toContain('draw:name="f0"');
    expect(xml).toContain("M ?f0 0");
  });

  it("re-emits a reduced path's source verbatim on a same-format save, losing nothing (AC-30)", async () => {
    const opened = await parseDocx(loadFixture(DOCX_FIXTURE));
    const saved = await writeDocx(modelFrom(opened, "docx"));
    const xml = await partText(saved.bytes, "word/document.xml");
    // The arc and quadratic the model reduced are still in the saved file,
    // so re-opening it in Word finds the geometry it wrote.
    expect(xml).toContain("<a:quadBezTo>");
    expect(xml).toContain('<a:arcTo wR="10800" hR="10800" stAng="0" swAng="5400000"/>');
    expect(saved.dropped.filter((drop) => drop.kind === "shapeGeometry")).toHaveLength(0);
  });
});

describe("SPEC-071 — the fixtures carry cross-format (AC-28, AC-30)", () => {
  it("carries every modelled path from .docx to .odt and back", async () => {
    const fromDocx = await parseDocx(loadFixture(DOCX_FIXTURE));
    // A real cross-format save keeps the source format on the model — that is
    // how the writer knows the save is a conversion at all.
    const asOdt = await writeOdt(modelFrom(fromDocx, "docx"));
    const reopened = await parseOdt(asOdt.bytes);
    const shapes = shapesOf(reopened);
    // The six path shapes survive as paths; the two preserved ones cannot
    // cross formats and land as the placeholder's "rect" stand-in.
    expect(shapes.filter((shape) => shape.path)).toHaveLength(6);
    for (let i = 0; i < 6; i++) {
      expect(opsOf(shapes[i]), LABELS[i]).toEqual(opsOf(shapesOf(fromDocx)[i]));
    }
  });

  it("carries every modelled path from .odt to .docx", async () => {
    const fromOdt = await parseOdt(loadFixture(ODT_FIXTURE));
    const asDocx = await writeDocx(modelFrom(fromOdt, "odt"));
    const shapes = shapesOf(await parseDocx(asDocx.bytes));
    expect(shapes.filter((shape) => shape.path)).toHaveLength(6);
    for (let i = 0; i < 6; i++) {
      expect(opsOf(shapes[i]), LABELS[i]).toEqual(opsOf(shapesOf(fromOdt)[i]));
    }
  });

  it("reports a drop only for the geometry a cross-format save really loses (AC-28, AC-30)", async () => {
    const fromDocx = await parseDocx(loadFixture(DOCX_FIXTURE));
    const asOdt = await writeOdt(modelFrom(fromDocx, "docx"));
    // Five losslessly-modelled paths report nothing. The two preserved
    // geometries report (SPEC-051 AC-61), and so does the reduced one, whose
    // arc and quadratic are written as the converted cubics (AC-30) — three
    // drops in total, not one per shape.
    expect(asOdt.dropped.filter((drop) => drop.kind === "shapeGeometry")).toHaveLength(3);
  });
});

describe("SPEC-071 — a stored preserved geometry is not reinterpreted (AC-33)", () => {
  it("leaves a shape already preserved as unsupportedGeometryXml alone, even when its path is modellable", async () => {
    // A document stored before this spec: the reader of the day put a
    // perfectly modellable custGeom into unsupportedGeometryXml. SPEC-008
    // promises those bytes, so nothing on the load or save path may look
    // inside them and mint a `path` — only reopening the source file picks
    // up the new reader (which the fixture tests above prove it does).
    const opened = await parseDocx(loadFixture(DOCX_FIXTURE));
    const modelled = shapesOf(opened)[1];
    const custGeom =
      `<a:custGeom xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:pathLst>` +
      `<a:path w="${PATH_GRID}" h="${PATH_GRID}"><a:moveTo><a:pt x="0" y="0"/></a:moveTo>` +
      `<a:lnTo><a:pt x="${PATH_GRID}" y="${PATH_GRID}"/></a:lnTo><a:close/></a:path></a:pathLst></a:custGeom>`;
    const stored: Shape = {
      type: "shape",
      geometry: "rect",
      widthCm: modelled.widthCm,
      heightCm: modelled.heightCm,
      wrap: "inline",
      unsupportedGeometryXml: custGeom,
    };
    const model: DocumentModel = { ...modelFrom(opened, "docx"), opaque: [], sourceXml: undefined };
    model.sections = [{ ...opened.sections[0], blocks: [stored] }];

    const saved = await writeDocx(model);
    const reopened = shapesOf(await parseDocx(saved.bytes))[0];
    // The saved file carries the preserved custGeom verbatim...
    expect(await partText(saved.bytes, "word/document.xml")).toContain('<a:pt x="0" y="0"/>');
    // ...and the stored shape itself never gained a path along the way.
    expect(stored.path).toBeUndefined();
    expect(stored.unsupportedGeometryXml).toBe(custGeom);
    // Reopening the file it produced is a *source* read, so the new reader
    // does model it — the asymmetry AC-33 describes, stated rather than
    // assumed.
    expect(reopened.path).toBeTruthy();
  });
});
