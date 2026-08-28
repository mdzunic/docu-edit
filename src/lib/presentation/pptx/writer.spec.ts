// The pptx deck-level write-back (SPEC-090): p:sldSz + firstSlideNum splices
// into presentation.xml, p:notesSz left verbatim, and docProps/core.xml rebuilt
// from the deck metadata.
import { describe, expect, it } from "vitest";
import { unzip } from "../../core/zip";
import { cmToEmu } from "../../core/units";
import { buildPptxFixtureBytes } from "../fixtures/build";
import { readPresentationBytes } from "../open";
import { applyDeckProperties, applyDeckSize, setFirstSlideNumber } from "../deckSetup";
import { writePptx } from "./writer";
import type { PresentationModel } from "../model";

const decoder = new TextDecoder("utf-8");
const MODIFIED = new Date("2026-03-01T12:00:00Z");

async function openDeck(): Promise<PresentationModel> {
  return readPresentationBytes(await buildPptxFixtureBytes(), "pptx", "Deck");
}

async function writtenParts(model: PresentationModel): Promise<Map<string, Uint8Array>> {
  const { bytes } = await writePptx(model, { modified: MODIFIED });
  return unzip(bytes);
}

function partText(parts: Map<string, Uint8Array>, name: string): string {
  return decoder.decode(parts.get(name)!);
}

describe("pptx p:sldSz splice (SPEC-090 AC-15)", () => {
  it("rewrites cx/cy to the new size in whole EMU, notesSz untouched", async () => {
    const model = await openDeck();
    applyDeckSize(model, { widthCm: 25.4, heightCm: 19.05 }, "keep");
    const pres = partText(await writtenParts(model), "ppt/presentation.xml");
    expect(pres).toContain(`<p:sldSz cx="${cmToEmu(25.4)}" cy="${cmToEmu(19.05)}"/>`);
    // notesSz is sized independently — left byte-verbatim.
    expect(pres).toContain(`<p:notesSz cx="6858000" cy="9144000"/>`);
  });

  it("leaves presentation.xml's sldSz verbatim when the deck is unchanged", async () => {
    const model = await openDeck();
    const pres = partText(await writtenParts(model), "ppt/presentation.xml");
    expect(pres).toContain(`<p:sldSz cx="12192000" cy="6858000"/>`);
  });
});

describe("pptx firstSlideNum (SPEC-090 AC-15)", () => {
  it("sets firstSlideNum on the p:presentation start tag when present", async () => {
    const model = await openDeck();
    setFirstSlideNumber(model, 7);
    const pres = partText(await writtenParts(model), "ppt/presentation.xml");
    expect(pres).toMatch(/<p:presentation\b[^>]*\bfirstSlideNum="7"[^>]*>/);
  });

  it("removes firstSlideNum when the origin returns to 1", async () => {
    const model = await openDeck();
    setFirstSlideNumber(model, 7);
    setFirstSlideNumber(model, 1);
    const pres = partText(await writtenParts(model), "ppt/presentation.xml");
    expect(pres).not.toContain("firstSlideNum");
  });
});

describe("pptx docProps/core.xml (SPEC-090 AC-20)", () => {
  it("rebuilds core.xml from the deck metadata on a deck change", async () => {
    const model = await openDeck();
    applyDeckProperties(model, { ...(model.properties ?? {}), author: "Grace Hopper", subject: "Q3" });
    const core = partText(await writtenParts(model), "docProps/core.xml");
    expect(core).toContain("Grace Hopper");
    expect(core).toContain("Q3");
  });

  it("leaves core.xml verbatim when the deck is unchanged", async () => {
    const model = await openDeck();
    const parts = await writtenParts(model);
    const source = await unzip(await buildPptxFixtureBytes());
    expect(partText(parts, "docProps/core.xml")).toBe(decoder.decode(source.get("docProps/core.xml")!));
  });
});
