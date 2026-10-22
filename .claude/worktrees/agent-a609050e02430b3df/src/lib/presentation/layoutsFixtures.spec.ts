// SPEC-089 fixture coverage: the pptx catalogue (six standard kinds + the exotic
// "other", plus a slide with no layout reference) and the odp catalogue decoded
// from LAYOUT_GOLDEN — the golden asserted both ways (reader decode + the
// styles.xml splice).
import { describe, expect, it } from "vitest";
import { buildSpec089OdpFixtureBytes, buildSpec089PptxFixtureBytes } from "./fixtures/build";
import { parsePptx } from "./pptx/reader";
import { parseOdp } from "./odp/reader";
import { authorPageLayoutStyle } from "./odp/writer";
import {
  GOLDEN_LAYOUTS,
  GOLDEN_TITLE_MARKUP,
  GOLDEN_TWO_MARKUP,
} from "./odp/layoutGolden";

describe("AC-2/5/19: pptx layout catalogue", () => {
  it("derives every standard kind plus the exotic 'other'", async () => {
    const parsed = await parsePptx(await buildSpec089PptxFixtureBytes());
    const kinds = parsed.layouts.map((l) => l.kind);
    expect(kinds).toEqual([
      "title",
      "titleAndBody",
      "sectionHeader",
      "twoContent",
      "titleOnly",
      "blank",
      "other", // the type="cust" exotic layout
    ]);
    // Keyed by absolute part path, named by p:cSld/@name.
    expect(parsed.layouts[0].key).toBe("ppt/slideLayouts/slideLayout1.xml");
    expect(parsed.layouts[3].name).toBe("Two Content");
    // Two-content slots carry verbatim phType/phIdx.
    const two = parsed.layouts[3];
    expect(two.slots.map((s) => [s.role, s.phType, s.phIdx])).toEqual([
      ["title", "title", undefined],
      ["body", "body", 1],
      ["body", "body", 2],
    ]);
  });

  it("assigns layoutKey per slide and leaves a no-reference slide unset", async () => {
    const parsed = await parsePptx(await buildSpec089PptxFixtureBytes());
    expect(parsed.slides[0].layoutKey).toBe("ppt/slideLayouts/slideLayout1.xml");
    expect(parsed.slides[1].layoutKey).toBeUndefined(); // slide 2 has no slideLayout rel
  });
});

describe("AC-3/4/19: odp catalogue from LAYOUT_GOLDEN, golden both ways", () => {
  it("reader decodes LAYOUT_GOLDEN to GOLDEN_LAYOUTS", async () => {
    const parsed = await parseOdp(await buildSpec089OdpFixtureBytes());
    // The three referenced page-layouts decode exactly to the golden; the
    // fourth page (no reference) adds the synthetic `Default::` entry.
    const golden = parsed.layouts.filter((l) => l.key !== "Default::");
    expect(golden).toEqual(GOLDEN_LAYOUTS);
    expect(parsed.slides[0].layoutKey).toBe("Default::AL-title");
    expect(parsed.slides[3].layoutKey).toBe("Default::");
  });

  it("the styles.xml splice reproduces the golden markup", () => {
    expect(authorPageLayoutStyle("AL-title", GOLDEN_LAYOUTS[0])).toBe(GOLDEN_TITLE_MARKUP);
    expect(authorPageLayoutStyle("AL-two", GOLDEN_LAYOUTS[1])).toBe(GOLDEN_TWO_MARKUP);
  });
});
