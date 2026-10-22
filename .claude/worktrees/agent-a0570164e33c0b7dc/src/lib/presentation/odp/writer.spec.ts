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

// --- SPEC-097: a picture bullet's media part ---------------------------------

import { createBlankPresentation } from "../slides";
import { markSlideEdited } from "../text";
import type { TextFrame } from "../model";

/** A 2×2 PNG for the picture bullet. */
const BULLET_PNG_BASE64 =
  "iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAIAAAD91JpzAAAAEklEQVR4nGMw7jhj3HGGAUIBACiuBh3+iAU4AAAAAElFTkSuQmCC";

const BULLET_PART = "Pictures/image97.png";

function bulletFrame(): TextFrame {
  return {
    id: "pic",
    kind: "text",
    xCm: 2,
    yCm: 2,
    widthCm: 12,
    heightCm: 4,
    paragraphs: [
      { runs: [{ text: "Pictured" }], bullet: "disc", bulletImageId: BULLET_PART },
      { runs: [{ text: "Also pictured" }], bullet: "disc", bulletImageId: BULLET_PART },
    ],
  };
}

describe("odp picture bullets (SPEC-097)", () => {
  it("an edited deck saves the bullet's Pictures/ part and its manifest entry, and reopens with it", async () => {
    const model = await openDeck();
    model.media = [...(model.media ?? []), { mediaId: BULLET_PART, contentType: "image/png", data: BULLET_PNG_BASE64 }];
    model.slides[0].frames.push(bulletFrame());
    markSlideEdited(model.slides[0], "odp");
    const { bytes } = await writeOdp(model, { modified: MODIFIED });
    const parts = await unzip(bytes);
    expect(Buffer.from(parts.get(BULLET_PART)!).toString("base64")).toBe(BULLET_PNG_BASE64);
    expect(partText(parts, "META-INF/manifest.xml")).toContain(`<manifest:file-entry manifest:full-path="${BULLET_PART}" manifest:media-type="image/png"/>`);
    expect(partText(parts, "content.xml")).toContain(`<text:list-level-style-image text:level="1" xlink:href="${BULLET_PART}"`);

    const reopened = await readPresentationBytes(bytes, "odp", "Deck");
    const frame = reopened.slides[0].frames.find((f): f is TextFrame => f.kind === "text" && f.paragraphs[0]?.runs[0]?.text === "Pictured");
    expect(frame?.paragraphs.map((p) => p.bulletImageId)).toEqual([BULLET_PART, BULLET_PART]);
    expect(reopened.media?.find((m) => m.mediaId === BULLET_PART)?.data).toBe(BULLET_PNG_BASE64);
  });

  it("a from-scratch deck writes the bullet's part and manifest entry too", async () => {
    const model: PresentationModel = { ...createBlankPresentation(), format: "odp" };
    model.media = [{ mediaId: BULLET_PART, contentType: "image/png", data: BULLET_PNG_BASE64 }];
    model.slides[0].frames.push(bulletFrame());
    const parts = await writtenParts(model);
    expect(parts.has(BULLET_PART)).toBe(true);
    expect(partText(parts, "META-INF/manifest.xml")).toContain(`manifest:full-path="${BULLET_PART}"`);
  });

  it("a bullet image no paragraph references any more is not written", async () => {
    const model = await openDeck();
    model.media = [...(model.media ?? []), { mediaId: BULLET_PART, contentType: "image/png", data: BULLET_PNG_BASE64 }];
    markSlideEdited(model.slides[0], "odp");
    const parts = await writtenParts(model);
    expect(parts.has(BULLET_PART)).toBe(false);
    expect(partText(parts, "META-INF/manifest.xml")).not.toContain(BULLET_PART);
  });
});
