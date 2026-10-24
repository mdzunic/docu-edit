// SPEC-098's wire round trip for slide-run language, shaped like
// transitionsRoundTrip.spec.ts: per-run `lang`/`noProof` survive save → reopen
// in both formats (field-level toEqual); every untouched part stays
// byte-identical (`equalBytes`, part by part); a slide whose runs carry no
// language rebuilds exactly as before; a replaced slide's part is byte-stable
// written twice; and the deck default is derived at read but never written.
// No semantic relation is claimed — lib/equality.ts's equalSemantic/resolve are
// typed against the text model.
import { describe, expect, it } from "vitest";
import { equalBytes } from "../core/equality";
import { compileQuery, type CompiledQuery } from "../core/findQuery";
import { unzip } from "../core/zip";
import { buildOdpFixtureBytes, buildPptxFixtureBytes } from "./fixtures/build";
import { findInDeck, replaceHit } from "./deckSearch";
import type { PresentationFormat, PresentationModel, TextFrame } from "./model";
import { readPresentationBytes } from "./open";
import { writeOdp } from "./odp/writer";
import { writePptx } from "./pptx/writer";
import { markSlideEdited } from "./text";

const FORMATS: { format: PresentationFormat; build: () => Promise<Uint8Array>; slidePart: string }[] = [
  { format: "pptx", build: buildPptxFixtureBytes, slidePart: "ppt/slides/slide2.xml" },
  { format: "odp", build: buildOdpFixtureBytes, slidePart: "content.xml" },
];

async function write(model: PresentationModel): Promise<Uint8Array> {
  return model.format === "pptx" ? (await writePptx(model)).bytes : (await writeOdp(model)).bytes;
}

/** Slide 2's body placeholder (the fixture's "Roadmap" list). */
function body(model: PresentationModel): TextFrame {
  const frame = model.slides[1].frames[1];
  if (frame.kind !== "text") throw new Error("fixture slide 2 has no body");
  return frame;
}

const languages = (frame: TextFrame) => frame.paragraphs.map((para) => para.runs.map(({ lang, noProof }) => ({ lang, noProof })));

describe.each(FORMATS)("$format: run language round trip", ({ format, build, slidePart }) => {
  it("per-run tags and noProof survive save → reopen, field by field", async () => {
    const model = await readPresentationBytes(await build(), format, "Deck");
    const [first, second, third, fourth] = body(model).paragraphs;
    first.runs = first.runs.map((run) => ({ ...run, lang: "de-DE" }));
    second.runs = second.runs.map((run) => ({ ...run, lang: "de-DE" }));
    third.runs = third.runs.map((run) => ({ ...run, lang: "en-GB" }));
    fourth.runs = fourth.runs.map((run) => ({ ...run, noProof: true as const }));
    const title = model.slides[1].frames[0];
    if (title.kind === "text") title.paragraphs[0].runs = title.paragraphs[0].runs.map((run) => ({ ...run, lang: "de-DE" }));
    markSlideEdited(model.slides[1], format);
    const back = await readPresentationBytes(await write(model), format, "Deck");
    // odp has no room for a real language beside zxx, so a noProof run reads
    // back noProof alone; pptx writes the default tag beside it.
    const noProofLang = format === "pptx" ? "en-US" : undefined;
    expect(languages(body(back))).toEqual([
      [{ lang: "de-DE", noProof: undefined }],
      [{ lang: "de-DE", noProof: undefined }],
      [{ lang: "en-GB", noProof: undefined }],
      [{ lang: noProofLang, noProof: true }],
    ]);
    // The deck default is the most frequent run tag, derived at read.
    expect(back.language).toBe("de-DE");
  });

  it("keeps every untouched part byte-identical, part by part", async () => {
    const source = await build();
    const model = await readPresentationBytes(source, format, "Deck");
    body(model).paragraphs[0].runs = body(model).paragraphs[0].runs.map((run) => ({ ...run, lang: "cs-CZ" }));
    markSlideEdited(model.slides[1], format);
    const before = await unzip(source);
    const after = await unzip(await write(model));
    expect([...after.keys()].toSorted()).toEqual([...before.keys()].toSorted());
    const changed = [...before.keys()].filter((name) => !equalBytes(before.get(name)!, after.get(name)!));
    expect(changed).toEqual([slidePart]);
  });

  it("rebuilds a slide whose runs carry no language exactly as before (pptx: the en-US literal; odp: no fo:language)", async () => {
    const model = await readPresentationBytes(await build(), format, "Deck");
    expect(body(model).paragraphs.flatMap((para) => para.runs).every((run) => run.lang === undefined && !run.noProof)).toBe(true);
    markSlideEdited(model.slides[1], format);
    const xml = new TextDecoder().decode((await unzip(await write(model))).get(slidePart)!);
    if (format === "pptx") {
      expect(xml.match(/<a:rPr [^>]*>/g)?.every((rPr) => rPr.startsWith(`<a:rPr lang="en-US" sz=`))).toBe(true);
      expect(xml).not.toContain("noProof");
    } else {
      expect(xml).not.toContain("fo:language");
    }
  });

  it("writes a replaced slide's part byte-stably twice", async () => {
    const model = await readPresentationBytes(await build(), format, "Deck");
    const compiled = compileQuery("Open", { matchCase: false, wholeWords: false, regex: false }) as CompiledQuery;
    const hit = findInDeck(model, compiled).find((each) => each.scope === "frame")!;
    expect(replaceHit(model, hit, "Launch")).toBe(true);
    const once = (await unzip(await write(model))).get(slidePart)!;
    const twice = (await unzip(await write(model))).get(slidePart)!;
    expect(equalBytes(once, twice)).toBe(true);
    const reopened = await readPresentationBytes(await write(model), format, "Deck");
    expect(body(reopened).paragraphs[3].runs.map((run) => run.text).join("")).toBe("Launch Berlin office");
    if (format === "pptx") {
      // pptx rebuilds the part from the re-read model to the same bytes too
      // (odp regenerates its automatic style names against the read styles).
      markSlideEdited(reopened.slides[1], format);
      expect(equalBytes(once, (await unzip(await write(reopened))).get(slidePart)!)).toBe(true);
    }
  });

  it("never serialises the deck default", async () => {
    const model = await readPresentationBytes(await build(), format, "Deck");
    model.language = "es-ES";
    markSlideEdited(model.slides[1], format);
    const parts = await unzip(await write(model));
    for (const [name, bytes] of parts) {
      const text = new TextDecoder().decode(bytes);
      expect(text.includes("es-ES") || /fo:language="es"/.test(text), name).toBe(false);
    }
    // Reopened, the default is derived from the runs again: pptx's rebuilt
    // slide declares the en-US literal on every run; odp writes none.
    expect((await readPresentationBytes(await write(model), format, "Deck")).language).toBe(format === "pptx" ? "en-US" : undefined);
  });
});
