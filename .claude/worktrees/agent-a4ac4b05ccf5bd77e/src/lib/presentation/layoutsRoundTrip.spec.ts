// SPEC-089 layout read/write round trips, driven over the in-repo mini-deck
// fixtures: repointing to an existing layout (pptx), authoring a synthesised
// layout part (pptx) / page-layout style (odp), and the byte-stability of every
// part the change did not touch.
import { describe, expect, it } from "vitest";
import { equalBytes } from "../core/equality";
import { unzip } from "../core/zip";
import { buildOdpFixtureBytes, buildPptxFixtureBytes } from "./fixtures/build";
import { readPresentationBytes } from "./open";
import { writePptx } from "./pptx/writer";
import { writeOdp } from "./odp/writer";
import { applySlideLayout, synthesiseStandardLayouts, standardLayoutKey, layoutFor } from "./layouts";
import type { PresentationModel } from "./model";

async function pptxModel(): Promise<PresentationModel> {
  return readPresentationBytes(await buildPptxFixtureBytes(), "pptx", "Deck");
}
async function odpModel(): Promise<PresentationModel> {
  return readPresentationBytes(await buildOdpFixtureBytes(), "odp", "Deck");
}

describe("AC-13/18: pptx repoint to an existing layout", () => {
  it("re-points only the changed slide's .rels and reproduces the layoutKey", async () => {
    const source = await buildPptxFixtureBytes();
    const model = await readPresentationBytes(source, "pptx", "Deck");
    // slide1 references slideLayout1 (title); switch it to slideLayout2 (obj).
    const layout2 = "ppt/slideLayouts/slideLayout2.xml";
    expect(model.slides[0].layoutKey).toBe("ppt/slideLayouts/slideLayout1.xml");
    const report = applySlideLayout(model, 0, layout2);
    expect(report.applied).toBe(true);

    const bytes = (await writePptx(model)).bytes;
    const before = await unzip(source);
    const after = await unzip(bytes);

    // Only slide1's own part + its .rels changed; slide2/3/4 rels are verbatim.
    expect(equalBytes(before.get("ppt/slides/_rels/slide2.xml.rels")!, after.get("ppt/slides/_rels/slide2.xml.rels")!)).toBe(true);
    expect(equalBytes(before.get("ppt/slides/_rels/slide4.xml.rels")!, after.get("ppt/slides/_rels/slide4.xml.rels")!)).toBe(true);
    const relsAfter = new TextDecoder().decode(after.get("ppt/slides/_rels/slide1.xml.rels")!);
    expect(relsAfter).toContain("slideLayout2.xml");
    expect(relsAfter).not.toContain("slideLayout1.xml");

    const back = await readPresentationBytes(bytes, "pptx", "Deck");
    expect(back.slides[0].layoutKey).toBe(layout2);
    // Title geometry snapped to layout2's title slot (2,1,29,3 cm).
    const title = back.slides[0].frames.find((f) => f.kind === "text" && f.role === "title")!;
    expect(title.xCm).toBeCloseTo(2, 1);
    expect(title.yCm).toBeCloseTo(1, 1);
  });

  it("a no-edit download stays byte-identical part-for-part", async () => {
    const source = await buildPptxFixtureBytes();
    const model = await readPresentationBytes(source, "pptx", "Deck");
    const bytes = (await writePptx(model)).bytes;
    const before = await unzip(source);
    const after = await unzip(bytes);
    expect([...after.keys()].toSorted()).toEqual([...before.keys()].toSorted());
    for (const [name, data] of before) {
      expect(equalBytes(data, after.get(name)!), `part ${name}`).toBe(true);
    }
  });
});

describe("AC-14/18: pptx authors a missing layout part", () => {
  it("authors slideLayout with .rels + Content-Types + master splice, reproduced on reopen", async () => {
    const model = await pptxModel();
    // Offer the synthesised standard set; the deck has no section-header layout.
    model.layouts = [...(model.layouts ?? []), ...synthesiseStandardLayouts(model.slideWidthCm, model.slideHeightCm)];
    applySlideLayout(model, 0, standardLayoutKey("sectionHeader"));
    const bytes = (await writePptx(model)).bytes;
    const after = await unzip(bytes);
    // A new slideLayout part past the two existing ones, with its own .rels.
    expect(after.has("ppt/slideLayouts/slideLayout3.xml")).toBe(true);
    expect(after.has("ppt/slideLayouts/_rels/slideLayout3.xml.rels")).toBe(true);
    const ct = new TextDecoder().decode(after.get("[Content_Types].xml")!);
    expect(ct).toContain("/ppt/slideLayouts/slideLayout3.xml");
    const master = new TextDecoder().decode(after.get("ppt/slideMasters/slideMaster1.xml")!);
    expect((master.match(/<p:sldLayoutId /g) ?? []).length).toBe(3);

    const back = await readPresentationBytes(bytes, "pptx", "Deck");
    expect(back.slides[0].layoutKey).toBe("ppt/slideLayouts/slideLayout3.xml");
    const entry = layoutFor(back, back.slides[0]);
    expect(entry?.kind).toBe("sectionHeader");
  });
});

describe("AC-16/18: odp authors a page-layout style", () => {
  it("splices a style:presentation-page-layout into styles.xml and stamps the page", async () => {
    const source = await buildOdpFixtureBytes();
    const model = await odpModel();
    model.layouts = [...(model.layouts ?? []), ...synthesiseStandardLayouts(model.slideWidthCm, model.slideHeightCm)];
    applySlideLayout(model, 0, standardLayoutKey("twoContent"));
    const bytes = (await writeOdp(model)).bytes;
    const before = await unzip(source);
    const after = await unzip(bytes);

    const styles = new TextDecoder().decode(after.get("styles.xml")!);
    expect(styles).toContain('style:presentation-page-layout style:name="de-AL1"');
    const content = new TextDecoder().decode(after.get("content.xml")!);
    expect(content).toContain('presentation:presentation-page-layout-name="de-AL1"');
    // A part the change did not touch stays byte-verbatim.
    expect(equalBytes(before.get("meta.xml")!, after.get("meta.xml")!)).toBe(true);

    const back = await readPresentationBytes(bytes, "odp", "Deck");
    const entry = layoutFor(back, back.slides[0]);
    expect(entry?.kind).toBe("twoContent");
  });

  it("a no-edit odp download stays byte-identical part-for-part", async () => {
    const source = await buildOdpFixtureBytes();
    const model = await readPresentationBytes(source, "odp", "Deck");
    const bytes = (await writeOdp(model)).bytes;
    const before = await unzip(source);
    const after = await unzip(bytes);
    for (const [name, data] of before) {
      expect(equalBytes(data, after.get(name)!), `part ${name}`).toBe(true);
    }
  });
});
