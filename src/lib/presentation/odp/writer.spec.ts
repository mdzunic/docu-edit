// The odp deck-level write-back (SPEC-090): page-layout size rewrite (with its
// notes-layout exclusion and no-layout fallback), and meta.xml rebuilt from the
// deck metadata. firstSlideNumber is deliberately not serialised to odp.
import { describe, expect, it } from "vitest";
import { unzip } from "../../core/zip";
import { buildOdpFixtureBytes } from "../fixtures/build";
import { readPresentationBytes } from "../open";
import { applyDeckProperties, applyDeckSize, setFirstSlideNumber } from "../deckSetup";
import { writeOdp } from "./writer";
import type { OpaquePart } from "../../core/packageParts";
import type { PresentationModel } from "../model";

const decoder = new TextDecoder("utf-8");
const MODIFIED = new Date("2026-03-01T12:00:00Z");

async function openDeck(): Promise<PresentationModel> {
  return readPresentationBytes(await buildOdpFixtureBytes(), "odp", "Deck");
}

async function writtenParts(model: PresentationModel): Promise<Map<string, Uint8Array>> {
  const { bytes } = await writeOdp(model, { modified: MODIFIED });
  return unzip(bytes);
}

function partText(parts: Map<string, Uint8Array>, name: string): string {
  return decoder.decode(parts.get(name)!);
}

function replaceStyles(model: PresentationModel, xml: string): void {
  const part = model.opaque!.find((p: OpaquePart) => p.name === "styles.xml")!;
  part.data = Buffer.from(xml, "utf8").toString("base64");
}

describe("odp page-layout size rewrite (SPEC-090 AC-16)", () => {
  it("rewrites the master's referenced layout width/height/orientation", async () => {
    const model = await openDeck();
    applyDeckSize(model, { widthCm: 25.4, heightCm: 19.05 }, "keep");
    const styles = partText(await writtenParts(model), "styles.xml");
    expect(styles).toContain('fo:page-width="25.4cm"');
    expect(styles).toContain('fo:page-height="19.05cm"');
    expect(styles).toContain('style:print-orientation="landscape"');
  });

  it("re-reads the rewritten size (build → reopen within ±0.01 cm)", async () => {
    const model = await openDeck();
    applyDeckSize(model, { widthCm: 25.4, heightCm: 19.05 }, "keep");
    const { bytes } = await writeOdp(model, { modified: MODIFIED });
    const reopened = await readPresentationBytes(bytes, "odp", "Deck");
    expect(reopened.slideWidthCm).toBeCloseTo(25.4, 2);
    expect(reopened.slideHeightCm).toBeCloseTo(19.05, 2);
  });

  it("leaves the notes page layout untouched", async () => {
    const model = await openDeck();
    replaceStyles(
      model,
      `<?xml version="1.0" encoding="UTF-8"?>` +
        `<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"` +
        ` xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"` +
        ` xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"` +
        ` xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0">` +
        `<office:automatic-styles>` +
        `<style:page-layout style:name="PLmain"><style:page-layout-properties fo:page-width="33.867cm" fo:page-height="19.05cm"/></style:page-layout>` +
        `<style:page-layout style:name="PLnotes"><style:page-layout-properties fo:page-width="19.05cm" fo:page-height="33.867cm"/></style:page-layout>` +
        `</office:automatic-styles>` +
        `<office:master-styles>` +
        `<style:master-page style:name="Default" style:page-layout-name="PLmain">` +
        `<presentation:notes style:page-layout-name="PLnotes"/>` +
        `</style:master-page>` +
        `</office:master-styles>` +
        `</office:document-styles>`,
    );
    applyDeckSize(model, { widthCm: 25.4, heightCm: 19.05 }, "keep");
    const styles = partText(await writtenParts(model), "styles.xml");
    // PLmain resized...
    expect(styles).toContain('<style:page-layout style:name="PLmain"><style:page-layout-properties fo:page-width="25.4cm"');
    // ...PLnotes left as it was.
    expect(styles).toContain('<style:page-layout style:name="PLnotes"><style:page-layout-properties fo:page-width="19.05cm" fo:page-height="33.867cm"/>');
  });

  it("gives a deck with no page layout a fresh PMde1", async () => {
    const model = await openDeck();
    replaceStyles(
      model,
      `<?xml version="1.0" encoding="UTF-8"?>` +
        `<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"` +
        ` xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"` +
        ` xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0">` +
        `<office:automatic-styles></office:automatic-styles>` +
        `<office:master-styles><style:master-page style:name="Default"/></office:master-styles>` +
        `</office:document-styles>`,
    );
    applyDeckSize(model, { widthCm: 25.4, heightCm: 19.05 }, "keep");
    const styles = partText(await writtenParts(model), "styles.xml");
    expect(styles).toContain('<style:page-layout style:name="PMde1">');
    expect(styles).toMatch(/<style:master-page style:name="Default" style:page-layout-name="PMde1"\/>/);
  });
});

describe("odp meta.xml (SPEC-090 AC-20)", () => {
  it("rebuilds meta.xml from the deck metadata on a deck change", async () => {
    const model = await openDeck();
    applyDeckProperties(model, { ...model.properties, author: "Grace Hopper", subject: "Q3" });
    const meta = partText(await writtenParts(model), "meta.xml");
    expect(meta).toContain("Grace Hopper");
    expect(meta).toContain("Q3");
  });

  it("leaves styles.xml and meta.xml verbatim when the deck is unchanged", async () => {
    const model = await openDeck();
    const written = await writtenParts(model);
    const source = await unzip(await buildOdpFixtureBytes());
    expect(partText(written, "styles.xml")).toBe(decoder.decode(source.get("styles.xml")!));
    expect(partText(written, "meta.xml")).toBe(decoder.decode(source.get("meta.xml")!));
  });
});

describe("odp firstSlideNumber is model-only (SPEC-090 AC-17)", () => {
  it("still goes structural but writes no numbering to any part", async () => {
    const model = await openDeck();
    setFirstSlideNumber(model, 4);
    const written = await writtenParts(model);
    for (const [, data] of written) {
      expect(decoder.decode(data)).not.toContain("firstSlideNum");
    }
    // The deck-level parts are still rewritten idempotently (deckDirty).
    expect(written.has("styles.xml")).toBe(true);
  });
});
