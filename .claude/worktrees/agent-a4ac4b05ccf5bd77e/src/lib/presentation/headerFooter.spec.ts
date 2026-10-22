// SPEC-093: the deck field resolver + apply logic. Framework-free, so tested
// without mounting anything.
import { describe, expect, it } from "vitest";
import type { PresentationModel, Slide, SlideLayout } from "./model";
import {
  DEFAULT_FIELD_BAND,
  applyHeaderFooter,
  defaultFieldRect,
  fieldText,
  headerFooterDraft,
  setFieldVisibility,
  slideNumberFor,
  type HeaderFooterValues,
} from "./headerFooter";

const INSTANT = new Date(2026, 4, 14, 9, 30, 0); // Thu 14 May 2026

function slide(overrides: Partial<Slide> = {}): Slide {
  return { id: crypto.randomUUID(), frames: [], ...overrides };
}

function model(overrides: Partial<PresentationModel> = {}): PresentationModel {
  return {
    kind: "presentation",
    id: "deck",
    title: "Deck",
    format: "pptx",
    slides: [slide(), slide(), slide()],
    slideWidthCm: 33.867,
    slideHeightCm: 19.05,
    ...overrides,
  };
}

const ALL_ON: HeaderFooterValues = {
  footer: true,
  footerText: "Confidential",
  date: true,
  dateSource: "auto",
  dateFormat: "d MMMM yyyy",
  dateFixedText: "",
  slideNumber: true,
  showOnTitleSlide: true,
};

describe("slideNumberFor", () => {
  it("counts from 1 when firstSlideNumber is absent", () => {
    const m = model();
    expect(slideNumberFor(m, 0)).toBe(1);
    expect(slideNumberFor(m, 2)).toBe(3);
  });

  it("counts from firstSlideNumber when present", () => {
    const m = model({ firstSlideNumber: 5 });
    expect(slideNumberFor(m, 0)).toBe(5);
    expect(slideNumberFor(m, 2)).toBe(7);
  });
});

describe("fieldText per kind", () => {
  it("renders the slide number live from index", () => {
    const m = model({ firstSlideNumber: 10 });
    expect(fieldText(m, 0, "slideNumber", INSTANT)).toBe("10");
    expect(fieldText(m, 1, "slideNumber", INSTANT)).toBe("11");
  });

  it("renders the deck footer, with a per-slide text override winning", () => {
    const m = model({ headerFooter: { footerText: "Deck footer" } });
    m.slides[1].fields = [{ kind: "footer", xCm: 0, yCm: 0, widthCm: 1, heightCm: 1, text: "Just this slide" }];
    expect(fieldText(m, 0, "footer", INSTANT)).toBe("Deck footer");
    expect(fieldText(m, 1, "footer", INSTANT)).toBe("Just this slide");
  });

  it("renders an auto date from the deck format at the given instant", () => {
    expect(fieldText(model(), 0, "date", INSTANT)).toBe("14 May 2026");
    const m = model({ headerFooter: { dateFormat: "yyyy-MM-dd" } });
    expect(fieldText(m, 0, "date", INSTANT)).toBe("2026-05-14");
  });

  it("renders a fixed date from the deck text", () => {
    const m = model({ headerFooter: { dateSource: "fixed", dateFixedText: "Spring 2026" } });
    expect(fieldText(m, 0, "date", INSTANT)).toBe("Spring 2026");
  });

  it("renders an out-of-subset date from its cache and never recomputes", () => {
    const m = model();
    m.slides[0].fields = [
      { kind: "date", xCm: 0, yCm: 0, widthCm: 1, heightCm: 1, wireType: "datetime13", text: "14.5.26" },
    ];
    expect(fieldText(m, 0, "date", INSTANT)).toBe("14.5.26");
    // A different instant must not change it.
    expect(fieldText(m, 0, "date", new Date(2030, 0, 1))).toBe("14.5.26");
  });

  it("falls back to the default pattern for an odp out-of-subset date (no cache)", () => {
    const m = model({ format: "odp" });
    m.slides[0].fields = [{ kind: "date", xCm: 0, yCm: 0, widthCm: 1, heightCm: 1, wireStyleName: "D1" }];
    expect(fieldText(m, 0, "date", INSTANT)).toBe("14 May 2026");
  });
});

describe("title-slide suppression", () => {
  it("returns '' for every kind on index 0 when showOnTitleSlide is false and no layouts", () => {
    const m = model({ headerFooter: { footerText: "Confidential", showOnTitleSlide: false } });
    expect(fieldText(m, 0, "footer", INSTANT)).toBe("");
    expect(fieldText(m, 0, "slideNumber", INSTANT)).toBe("");
    expect(fieldText(m, 0, "date", INSTANT)).toBe("");
    // Non-title slides still render.
    expect(fieldText(m, 1, "slideNumber", INSTANT)).toBe("2");
  });

  it("shows fields on index 0 when showOnTitleSlide is not false", () => {
    const m = model({ headerFooter: { footerText: "Confidential" } });
    expect(fieldText(m, 0, "footer", INSTANT)).toBe("Confidential");
  });

  it("uses the SPEC-089 layout type when a layout model resolves", () => {
    const titleLayout: SlideLayout = { key: "L-title", kind: "title", name: "Title", slots: [] };
    const bodyLayout: SlideLayout = { key: "L-body", kind: "titleAndBody", name: "Body", slots: [] };
    const m = model({
      layouts: [titleLayout, bodyLayout],
      headerFooter: { footerText: "Confidential", showOnTitleSlide: false },
    });
    // Index 1 carries the title layout → suppressed; index 0 carries body → shown.
    m.slides[0].layoutKey = "L-body";
    m.slides[1].layoutKey = "L-title";
    expect(fieldText(m, 0, "footer", INSTANT)).toBe("Confidential");
    expect(fieldText(m, 1, "footer", INSTANT)).toBe("");
  });
});

describe("DEFAULT_FIELD_BAND geometry", () => {
  it("lays out three equal boxes with gaps in the bottom band", () => {
    const w = 33.867;
    const h = 19.05;
    const date = defaultFieldRect("date", w, h);
    const footer = defaultFieldRect("footer", w, h);
    const number = defaultFieldRect("slideNumber", w, h);
    const bandWidth = w - 2 * DEFAULT_FIELD_BAND.sideInsetCm;
    const boxWidth = (bandWidth - 2 * DEFAULT_FIELD_BAND.gapCm) / 3;
    expect(date.widthCm).toBeCloseTo(boxWidth, 6);
    expect(date.xCm).toBeCloseTo(DEFAULT_FIELD_BAND.sideInsetCm, 6);
    expect(footer.xCm).toBeCloseTo(DEFAULT_FIELD_BAND.sideInsetCm + boxWidth + DEFAULT_FIELD_BAND.gapCm, 6);
    expect(number.xCm).toBeCloseTo(DEFAULT_FIELD_BAND.sideInsetCm + 2 * (boxWidth + DEFAULT_FIELD_BAND.gapCm), 6);
    // Right box ends flush with the right inset.
    expect(number.xCm + number.widthCm).toBeCloseTo(w - DEFAULT_FIELD_BAND.sideInsetCm, 6);
    // Vertical placement.
    expect(date.yCm).toBeCloseTo(h - DEFAULT_FIELD_BAND.bottomInsetCm, 6);
    expect(date.heightCm).toBe(DEFAULT_FIELD_BAND.heightCm);
    // Alignment order date/footer/number = left/center/right.
    expect(date.align).toBe("left");
    expect(footer.align).toBe("center");
    expect(number.align).toBe("right");
  });
});

describe("setFieldVisibility", () => {
  it("adds a default-band entry, is idempotent, and removes cleanly", () => {
    const m = model();
    const s = m.slides[0];
    setFieldVisibility(m, s, "footer", true);
    expect(s.fields?.map((f) => f.kind)).toEqual(["footer"]);
    const geom = s.fields![0];
    expect(geom.align).toBe("center");
    setFieldVisibility(m, s, "footer", true); // idempotent
    expect(s.fields?.length).toBe(1);
    setFieldVisibility(m, s, "date", true);
    expect(s.fields?.map((f) => f.kind).toSorted()).toEqual(["date", "footer"]);
    setFieldVisibility(m, s, "footer", false);
    expect(s.fields?.map((f) => f.kind)).toEqual(["date"]);
    setFieldVisibility(m, s, "date", false);
    expect(s.fields).toBeUndefined(); // array dropped when it empties
  });
});

describe("applyHeaderFooter scope rules and dirtying", () => {
  it("Apply writes deck values + this slide's visibility only, and dirties", () => {
    const m = model();
    applyHeaderFooter(m, ALL_ON, { mode: "slide", index: 1 });
    expect(m.headerFooter).toEqual({ footerText: "Confidential" });
    expect(m.headerFooterDirty).toBe(true);
    expect(m.slides[0].fields).toBeUndefined();
    expect(m.slides[1].fields?.map((f) => f.kind).toSorted()).toEqual(["date", "footer", "slideNumber"]);
    expect(m.slides[2].fields).toBeUndefined();
    expect(m.slides[1].dirty).toBe(true); // markSlideEdited
    expect(m.slides[0].dirty).toBeUndefined();
  });

  it("Apply to all writes visibility to every slide and drops per-slide text overrides", () => {
    const m = model();
    m.slides[0].fields = [{ kind: "footer", xCm: 0, yCm: 0, widthCm: 1, heightCm: 1, text: "old override" }];
    applyHeaderFooter(m, ALL_ON, { mode: "all" });
    for (const s of m.slides) {
      expect(s.fields?.map((f) => f.kind).toSorted()).toEqual(["date", "footer", "slideNumber"]);
      expect(s.dirty).toBe(true);
      expect(s.fields?.find((f) => f.kind === "footer")?.text).toBeUndefined();
    }
  });

  it("stores only non-default deck values and clears headerFooter when all default", () => {
    const m = model();
    applyHeaderFooter(
      m,
      { footer: false, footerText: "", date: false, dateSource: "auto", dateFormat: "d MMMM yyyy", dateFixedText: "", slideNumber: false, showOnTitleSlide: true },
      { mode: "all" },
    );
    expect(m.headerFooter).toBeUndefined();
  });

  it("odp apply deletes sourceXml on every touched slide (markSlideEdited)", () => {
    const m = model({ format: "odp" });
    for (const s of m.slides) {
      s.sourceXml = "<draw:page/>";
    }
    applyHeaderFooter(m, ALL_ON, { mode: "all" });
    for (const s of m.slides) {
      expect(s.sourceXml).toBeUndefined();
    }
  });

  it("round-trips a draft through the model", () => {
    const m = model();
    applyHeaderFooter(m, { ...ALL_ON, showOnTitleSlide: false, dateSource: "fixed", dateFixedText: "Q2" }, { mode: "slide", index: 0 });
    const draft = headerFooterDraft(m, m.slides[0]);
    expect(draft.footer).toBe(true);
    expect(draft.footerText).toBe("Confidential");
    expect(draft.date).toBe(true);
    expect(draft.dateSource).toBe("fixed");
    expect(draft.dateFixedText).toBe("Q2");
    expect(draft.slideNumber).toBe(true);
    expect(draft.showOnTitleSlide).toBe(false);
  });
});
