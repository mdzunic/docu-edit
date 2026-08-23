// SPEC-027 — the Page setup dialog's OK command: applying paper size,
// orientation, margins and the measurement-unit preference, snapped onto
// the twip grid (AC-10) with any stale preserved `w:sectPr` dropped so the
// writer regenerates pgSz/pgMar from the new values instead of reproducing
// the source document's original ones.
import { describe, it, expect } from "vitest";
import { createDocument, type DocumentModel, type Paragraph } from "./documentModel";
import type { CaretPosition } from "./textEditing";
import {
  caretSelection,
  setHeaderFooterDistance,
  setPageMargins,
  setPageSetup,
  type ModelSelection,
} from "./commands";
import { cmToTwips, twipsToCm } from "./core/units";

// The same twip-grid snap setPageSetup applies internally (commands.ts's
// snapCm) — 29.7cm (A4's nominal height) is itself not exactly representable
// in twips, so even a freshly-entered "round" value comes back very slightly
// off; expectations below snap through the same grid instead of asserting
// the idealized decimal.
function snap(cm: number): number {
  return twipsToCm(cmToTwips(cm));
}

function paragraph(text = ""): Paragraph {
  return { type: "paragraph", runs: [{ text, formatting: {} }] };
}

function doc(): DocumentModel {
  const model = createDocument("page-setup-test");
  model.sections = [{ blocks: [paragraph("Hello")] }];
  return model;
}

function caret(): CaretPosition {
  return { section: 0, block: 0, offset: 0 };
}

function selection(): ModelSelection {
  return caretSelection(caret());
}

describe("setPageSetup (SPEC-027 AC-1..7)", () => {
  it("sets the page size and margins on defaultStyles", () => {
    const result = setPageSetup(doc(), selection(), {
      pageSize: { widthCm: 29.7, heightCm: 21 },
      margins: { topCm: 3, rightCm: 2, bottomCm: 3, leftCm: 2 },
    });
    expect(result.model.defaultStyles.pageSize).toEqual({ widthCm: snap(29.7), heightCm: snap(21) });
    expect(result.model.defaultStyles.margins).toEqual({
      topCm: snap(3),
      rightCm: snap(2),
      bottomCm: snap(3),
      leftCm: snap(2),
    });
  });

  it("sets the measurement-unit preference when given (AC-7)", () => {
    const result = setPageSetup(doc(), selection(), {
      pageSize: { widthCm: 21, heightCm: 29.7 },
      margins: { topCm: 2.54, rightCm: 2.54, bottomCm: 2.54, leftCm: 2.54 },
      measurementUnit: "in",
    });
    expect(result.model.measurementUnit).toBe("in");
  });

  it("leaves the measurement-unit preference untouched when omitted", () => {
    const model = doc();
    model.measurementUnit = "mm";
    const result = setPageSetup(model, selection(), {
      pageSize: { widthCm: 21, heightCm: 29.7 },
      margins: { topCm: 2.54, rightCm: 2.54, bottomCm: 2.54, leftCm: 2.54 },
    });
    expect(result.model.measurementUnit).toBe("mm");
  });

  it("snaps page size and margins onto the twip grid (AC-10 — no drift)", () => {
    // 2.5cm is not exactly representable in twips; snapping must move it to
    // the value .docx's reader would compute back after a save.
    const result = setPageSetup(doc(), selection(), {
      pageSize: { widthCm: 21, heightCm: 29.7 },
      margins: { topCm: 2.5, rightCm: 2.5, bottomCm: 2.5, leftCm: 2.5 },
    });
    const snapped = result.model.defaultStyles.margins.topCm;
    // Applying setPageSetup again with the already-snapped value must be a
    // no-op — the defining property of "on the grid".
    const again = setPageSetup(result.model, selection(), {
      pageSize: result.model.defaultStyles.pageSize,
      margins: result.model.defaultStyles.margins,
    });
    expect(again.model.defaultStyles.margins.topCm).toBe(snapped);
  });

  it("drops a stale preserved docxSectPr so the writer regenerates it (AC-10)", () => {
    const model = doc();
    model.format = "docx";
    model.sourceXml = { docxSectPr: "<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/></w:sectPr>" };
    const result = setPageSetup(model, selection(), {
      pageSize: { widthCm: 29.7, heightCm: 21 },
      margins: { topCm: 2, rightCm: 2, bottomCm: 2, leftCm: 2 },
    });
    expect(result.model.sourceXml?.docxSectPr).toBeUndefined();
  });

  it("leaves sourceXml untouched when there is no preserved sectPr", () => {
    const model = doc();
    const result = setPageSetup(model, selection(), {
      pageSize: { widthCm: 21, heightCm: 29.7 },
      margins: { topCm: 2.54, rightCm: 2.54, bottomCm: 2.54, leftCm: 2.54 },
    });
    expect(result.model.sourceXml).toBeUndefined();
  });

  it("preserves the input selection (no paragraph content touched)", () => {
    const result = setPageSetup(doc(), selection(), {
      pageSize: { widthCm: 21, heightCm: 29.7 },
      margins: { topCm: 2.54, rightCm: 2.54, bottomCm: 2.54, leftCm: 2.54 },
    });
    expect(result.model.sections).toEqual(doc().sections);
  });
});

describe("setPageMargins also drops a stale preserved docxSectPr (SPEC-027 AC-10)", () => {
  it("clears docxSectPr on a ruler-drag margin change", () => {
    const model = doc();
    model.format = "docx";
    model.sourceXml = { docxSectPr: "<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/></w:sectPr>" };
    const result = setPageMargins(model, selection(), { leftCm: 3 });
    expect(result.model.sourceXml?.docxSectPr).toBeUndefined();
  });
});

describe("setPageMargins top/bottom (SPEC-064 AC-2a — the vertical ruler)", () => {
  it("sets the top margin from a vertical-ruler drag, snapped to the grid", () => {
    const result = setPageMargins(doc(), selection(), { topCm: 3 });
    expect(result.model.defaultStyles.margins.topCm).toBe(snap(3));
    // The other three margins are untouched.
    expect(result.model.defaultStyles.margins.bottomCm).toBe(doc().defaultStyles.margins.bottomCm);
    expect(result.model.defaultStyles.margins.leftCm).toBe(doc().defaultStyles.margins.leftCm);
  });

  it("sets the bottom margin and clamps a negative drag to zero", () => {
    const result = setPageMargins(doc(), selection(), { bottomCm: -5 });
    expect(result.model.defaultStyles.margins.bottomCm).toBe(0);
  });

  it("edits the caret section's own margins when it carries them (SPEC-028)", () => {
    const model = doc();
    model.sections[0].margins = { topCm: 2, rightCm: 2, bottomCm: 2, leftCm: 2 };
    const result = setPageMargins(model, selection(), { topCm: 4 });
    expect(result.model.sections[0].margins?.topCm).toBe(snap(4));
    // The document defaults stay put.
    expect(result.model.defaultStyles.margins.topCm).toBe(model.defaultStyles.margins.topCm);
  });
});

describe("setHeaderFooterDistance (SPEC-064 AC-3a/AC-3b)", () => {
  it("writes the caret section's header distance from the top edge", () => {
    const result = setHeaderFooterDistance(doc(), selection(), { headerDistanceCm: 1.5 });
    expect(result.model.sections[0].headerDistanceCm).toBe(snap(1.5));
    expect(result.model.sections[0].footerDistanceCm).toBeUndefined();
  });

  it("writes the caret section's footer distance from the bottom edge", () => {
    const result = setHeaderFooterDistance(doc(), selection(), { footerDistanceCm: 0.8 });
    expect(result.model.sections[0].footerDistanceCm).toBe(snap(0.8));
  });

  it("clamps a negative distance to zero", () => {
    const result = setHeaderFooterDistance(doc(), selection(), { headerDistanceCm: -2 });
    expect(result.model.sections[0].headerDistanceCm).toBe(0);
  });

  it("leaves the page margins and content untouched", () => {
    const result = setHeaderFooterDistance(doc(), selection(), { headerDistanceCm: 1.5 });
    expect(result.model.defaultStyles.margins).toEqual(doc().defaultStyles.margins);
    expect(result.model.sections[0].blocks).toEqual(doc().sections[0].blocks);
  });
});
