// Generational round-trip stability, over every `.docx` and `.odt` fixture.
//
// The per-format round-trip specs each save a fixture ONCE and compare the
// re-opened model against the original. That shape cannot see the whole class
// of bugs where a save is not idempotent — where saving the file a document
// has already been saved as changes it again, and again, and again:
//
//   - a floating .odt picture emitted both inside its anchor paragraph and at
//     its own block slot, so every save added one more copy of it;
//   - .docx relationships and [Content_Types].xml Overrides appended beside
//     the source's own, so a document with a header went OPC-invalid on save
//     two and grew by six declarations per save thereafter.
//
// Neither showed up in a single save, and neither disturbed the model our own
// reader builds — the second one produced a package Word refuses while every
// model-equality assertion in the suite stayed green. Autosave runs at up to
// two saves a minute, so "drifts a little per save" is not a slow leak; it is
// a document that degrades while someone is working in it.
//
// So: open, save, open, save, open, save. From the first save on, nothing may
// move — not the model, not the package's parts, not its validity.
import { describe, it, expect } from "vitest";
import { readFileSync, readdirSync } from "node:fs";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";
import { readDocxBytes, readOdtBytes } from "./openDocument";
import { writeDocx } from "./docx/writer";
import { writeOdt } from "./odf/writer";
import { validateDocx } from "./docx/validate";
import { unzip } from "./core/zip";
import { equalStructural, firstStructuralDifference } from "./equality";
import type { Block, DocumentModel } from "./documentModel";

const here = dirname(fileURLToPath(import.meta.url));
const docxFixtures = join(here, "docx", "fixtures");
const odfFixtures = join(here, "odf", "fixtures");

const DOCX = readdirSync(docxFixtures).filter((name) => name.endsWith(".docx")).sort();
const ODT = readdirSync(odfFixtures).filter((name) => name.endsWith(".odt")).sort();

/** How many saves to chain. Three is the smallest number that distinguishes
 * "settles after the first save" from "moves a little every time": the
 * relationship accumulation looked identical to a one-off first-save change
 * until generation three showed it still climbing. */
const GENERATIONS = 3;

/** Every save stamps a fresh `modified` timestamp (see writeDocx's core
 * properties), which is correct and would otherwise make every generation
 * differ. Nothing else about the properties is normalised. */
function withoutModified(model: DocumentModel): DocumentModel {
  if (!model.properties?.modified) {
    return model;
  }
  const properties = { ...model.properties };
  delete properties.modified;
  return { ...model, properties };
}

/** The document's block structure as a flat list of type names, tables
 * included — the shape a duplicated or dropped block changes even when every
 * block the old assertions indexed into still looks right. */
function blockShape(model: DocumentModel): string[] {
  const shape: string[] = [];
  const walk = (blocks: readonly Block[]): void => {
    for (const block of blocks) {
      shape.push(block.type);
      if (block.type === "table") {
        for (const row of block.rows) {
          for (const cell of row.cells) {
            walk(cell.blocks);
          }
        }
      }
    }
  };
  for (const section of model.sections) {
    walk(section.blocks);
  }
  return shape;
}

interface Generation {
  model: DocumentModel;
  bytes: Uint8Array;
  /** Part name -> uncompressed byte length. Compared instead of the zip's own
   * size, which wobbles by a byte or two on its own: every save stamps a
   * fresh `modified` timestamp, and two same-length timestamps do not
   * necessarily deflate to the same number of bytes. This inventory is
   * deterministic, and says more — it names the part that grew. */
  inventory: Record<string, number>;
}

async function inventoryOf(bytes: Uint8Array): Promise<Record<string, number>> {
  const inventory: Record<string, number> = {};
  for (const [name, data] of await unzip(bytes)) {
    inventory[name] = data.length;
  }
  return inventory;
}

async function chain(
  first: Uint8Array,
  read: (bytes: Uint8Array) => Promise<DocumentModel>,
  write: (model: DocumentModel) => Promise<Uint8Array>,
): Promise<Generation[]> {
  const generations: Generation[] = [];
  let bytes = first;
  for (let generation = 0; generation < GENERATIONS; generation++) {
    const model = await read(bytes);
    bytes = await write(model);
    generations.push({ model: await read(bytes), bytes, inventory: await inventoryOf(bytes) });
  }
  return generations;
}

const readDocx = (bytes: Uint8Array) => readDocxBytes(bytes, "Fixture");
const saveDocx = async (model: DocumentModel) => (await writeDocx(model)).bytes;
const readOdt = (bytes: Uint8Array) => readOdtBytes(bytes, "Fixture");
const saveOdt = async (model: DocumentModel) => (await writeOdt(model)).bytes;

/** The assertions every chain shares: from the first save on, the model, the
 * block shape and the package's parts all hold still. */
function expectStable(generations: Generation[], label: string): void {
  const base = generations[0];
  for (let index = 1; index < generations.length; index++) {
    const current = generations[index];
    expect(blockShape(current.model), `${label}: block shape at generation ${index + 1}`).toEqual(
      blockShape(base.model),
    );
    const difference = firstStructuralDifference(
      withoutModified(base.model),
      withoutModified(current.model),
    );
    expect(
      equalStructural(withoutModified(base.model), withoutModified(current.model)),
      `${label}: model diverged at generation ${index + 1} (${difference})`,
    ).toBe(true);
    expect(current.inventory, `${label}: package contents at generation ${index + 1}`).toEqual(
      base.inventory,
    );
  }
}

describe("repeated saves are idempotent — .docx", () => {
  it.each(DOCX)("%s survives three generations unchanged", async (name) => {
    const generations = await chain(
      new Uint8Array(readFileSync(join(docxFixtures, name))),
      readDocx,
      saveDocx,
    );
    expectStable(generations, name);
  });

  it.each(DOCX)("%s stays a valid package at every generation", async (name) => {
    const generations = await chain(
      new Uint8Array(readFileSync(join(docxFixtures, name))),
      readDocx,
      saveDocx,
    );
    for (const [index, generation] of generations.entries()) {
      expect(await validateDocx(generation.bytes), `${name}: generation ${index + 1}`).toEqual([]);
    }
  });
});

describe("repeated saves are idempotent — .odt", () => {
  it.each(ODT)("%s survives three generations unchanged", async (name) => {
    const generations = await chain(
      new Uint8Array(readFileSync(join(odfFixtures, name))),
      readOdt,
      saveOdt,
    );
    expectStable(generations, name);
  });
});

// A conversion is not a one-way trip: someone who works in both applications
// saves the same document back and forth for as long as the document lives.
//
// Crossing formats IS allowed to lose things — one format cannot express
// everything the other can, and the writer says so: every item it cannot
// carry becomes a DropRecord, which the conversion report dialog shows the
// user before they commit to it (SPEC-008).
//
// What is not allowed is losing something SILENTLY. So the invariant here is
// not "nothing changes" — it is "a crossing that reports nothing dropped
// changed nothing". A writer that returns an empty `dropped` has promised the
// conversion was lossless, and this holds it to that promise.
//
// Today that is asserted at the level of block STRUCTURE. The stronger form —
// comparing the blocks' own content — does not hold yet, and the gap is worth
// naming rather than leaving for someone to rediscover. Running this file with
// `equalStructural` over `{sections, styles, media}` instead of `blockShape`
// reports, on crossings that claim to have dropped nothing:
//
//   word-shapes.docx     blocks[1].geometry            "arrow" becomes "line"
//   word-freeform.docx   blocks[6].unsupportedGeometryXml  preserved geometry lost
//   lo-shapes.odt        blocks[14].zOrder             changes
//   libreoffice-writer.odt / openoffice-era.odt
//                        blocks[0].formatting.spaceAfterPt  changes
//
// (plus `mediaId`, `sourcePart` and `styleId` churn, which are identities and
// package pointers rather than content, and would need excluding the way
// equality.ts already excludes `id`). Each is a silent cross-format loss: the
// user is shown a conversion report that says nothing was dropped. They are a
// separate class from the same-format idempotence bugs above — the conversion
// layer, not the packaging layer — and want their own review round.
describe("a cross-format save that reports no drops loses nothing", () => {
  /** One crossing: writes `model`, re-reads it, and reports what the writer
   * said it dropped along the way. */
  async function cross(
    model: DocumentModel,
    write: (model: DocumentModel) => Promise<{ bytes: Uint8Array; dropped: unknown[] }>,
    read: (bytes: Uint8Array) => Promise<DocumentModel>,
  ): Promise<{ model: DocumentModel; dropped: unknown[] }> {
    const written = await write(model);
    return { model: await read(written.bytes), dropped: written.dropped };
  }

  async function expectSilentCrossingIsLossless(
    source: DocumentModel,
    label: string,
    steps: readonly {
      write: (model: DocumentModel) => Promise<{ bytes: Uint8Array; dropped: unknown[] }>;
      read: (bytes: Uint8Array) => Promise<DocumentModel>;
    }[],
  ): Promise<void> {
    let current = source;
    for (const [index, step] of steps.entries()) {
      const next = await cross(current, step.write, step.read);
      if (next.dropped.length === 0) {
        expect(blockShape(next.model), `${label}: block shape at crossing ${index + 1}`).toEqual(
          blockShape(current),
        );
      }
      current = next.model;
    }
  }

  const toOdt = { write: (m: DocumentModel) => writeOdt(m), read: readOdt };
  const toDocx = { write: (m: DocumentModel) => writeDocx(m), read: readDocx };

  it.each(DOCX)("%s: docx -> odt -> docx -> odt", async (name) => {
    const source = await readDocx(new Uint8Array(readFileSync(join(docxFixtures, name))));
    await expectSilentCrossingIsLossless(source, name, [toOdt, toDocx, toOdt]);
  });

  it.each(ODT)("%s: odt -> docx -> odt -> docx", async (name) => {
    const source = await readOdt(new Uint8Array(readFileSync(join(odfFixtures, name))));
    await expectSilentCrossingIsLossless(source, name, [toDocx, toOdt, toDocx]);
  });
});
