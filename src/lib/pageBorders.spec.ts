import { describe, expect, it } from "vitest";
import {
  DEFAULT_PAGE_BORDER_OFFSET_PT,
  pageBackgroundFromDraft,
  pageBorderDraftFrom,
  pageBordersFromDraft,
} from "./pageBorders";
import type { PageBorders } from "./documentModel";

const LINE = { style: "double" as const, widthPt: 3, color: "#c00000" };

describe("pageBorderDraftFrom (SPEC-074 AC-19..AC-23)", () => {
  it("opens empty for a section with no border", () => {
    const draft = pageBorderDraftFrom(undefined, undefined);
    expect(draft.top).toBe(false);
    expect(draft.bottom).toBe(false);
    expect(draft.left).toBe(false);
    expect(draft.right).toBe(false);
    expect(draft.offsetPt).toBe(DEFAULT_PAGE_BORDER_OFFSET_PT);
    expect(draft.measureFrom).toBe("page");
    expect(draft.appliesTo).toBe("allPages");
    expect(draft.backgroundOn).toBe(false);
  });

  it("reads which edges are bordered (AC-20)", () => {
    const draft = pageBorderDraftFrom({ top: LINE, bottom: LINE }, undefined);
    expect([draft.top, draft.bottom, draft.left, draft.right]).toEqual([true, true, false, false]);
  });

  it("takes the shared line from the first bordered edge (AC-19)", () => {
    const draft = pageBorderDraftFrom({ right: LINE }, undefined);
    expect(draft.style).toBe("double");
    expect(draft.widthPt).toBe(3);
    expect(draft.color).toBe("#c00000");
  });

  it("shows the automatic colour as the black it paints as", () => {
    const draft = pageBorderDraftFrom({ top: { style: "single", widthPt: 1, color: "auto" } }, undefined);
    expect(draft.color).toBe("#000000");
  });

  it("reads the offset and its reference edge (AC-21)", () => {
    const draft = pageBorderDraftFrom({ top: LINE, offsetPt: 12, measureFrom: "text" }, undefined);
    expect(draft.offsetPt).toBe(12);
    expect(draft.measureFrom).toBe("text");
  });

  it("reads the page background, distinguishing absent from a colour (AC-23)", () => {
    expect(pageBorderDraftFrom(undefined, undefined).backgroundOn).toBe(false);
    const set = pageBorderDraftFrom(undefined, "#ffffff");
    expect(set.backgroundOn).toBe(true);
    expect(set.background).toBe("#ffffff");
  });

  it("reads an empty string as no background, not as a colour (AC-23)", () => {
    // "" is not a colour: were it read as one, the switch would come up on
    // and OK would commit "" as the section's fill — which no format can
    // write (see odf/writer's fo:background-color).
    const draft = pageBorderDraftFrom(undefined, "");
    expect(draft.backgroundOn).toBe(false);
    expect(pageBackgroundFromDraft(draft)).toBeNull();
  });
});

describe("pageBordersFromDraft", () => {
  it("is null when no edge is chosen", () => {
    expect(pageBordersFromDraft(pageBorderDraftFrom(undefined, undefined))).toBeNull();
  });

  it("draws the shared line on every chosen edge (AC-19, AC-20)", () => {
    const draft = { ...pageBorderDraftFrom(undefined, undefined), top: true, left: true, style: "dashed" as const, widthPt: 2, color: "#0070c0" };
    const borders = pageBordersFromDraft(draft) as PageBorders;
    expect(borders.top).toEqual({ style: "dashed", widthPt: 2, color: "#0070c0" });
    expect(borders.left).toEqual({ style: "dashed", widthPt: 2, color: "#0070c0" });
    expect(borders.bottom).toBeUndefined();
    expect(borders.right).toBeUndefined();
  });

  it("carries the offset and its reference edge (AC-21)", () => {
    const draft = { ...pageBorderDraftFrom(undefined, undefined), top: true, offsetPt: 6, measureFrom: "text" as const };
    const borders = pageBordersFromDraft(draft) as PageBorders;
    expect(borders.offsetPt).toBe(6);
    expect(borders.measureFrom).toBe("text");
  });

  it("leaves appliesTo absent for the all-pages default (AC-22)", () => {
    const all = { ...pageBorderDraftFrom(undefined, undefined), top: true };
    expect((pageBordersFromDraft(all) as PageBorders).appliesTo).toBeUndefined();
    const first = { ...all, appliesTo: "firstPage" as const };
    expect((pageBordersFromDraft(first) as PageBorders).appliesTo).toBe("firstPage");
  });

  it("round-trips a draft back through pageBorderDraftFrom", () => {
    const borders: PageBorders = { top: LINE, bottom: LINE, offsetPt: 10, measureFrom: "text", appliesTo: "exceptFirstPage" };
    const back = pageBordersFromDraft(pageBorderDraftFrom(borders, "#eeffee"));
    expect(back).toEqual(borders);
  });
});

describe("pageBackgroundFromDraft (AC-23)", () => {
  it("is null when the background is off, the colour when it is on", () => {
    const off = pageBorderDraftFrom(undefined, undefined);
    expect(pageBackgroundFromDraft(off)).toBeNull();
    expect(pageBackgroundFromDraft({ ...off, backgroundOn: true, background: "#fdf6e3" })).toBe("#fdf6e3");
  });
});
