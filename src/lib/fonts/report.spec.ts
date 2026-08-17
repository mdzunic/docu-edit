import { describe, expect, it } from "vitest";
import type { DocumentModel } from "../documentModel";
import { createDefaultStyles } from "../documentModel";
import { documentFontReport, fontStatusOf } from "./report";

function docWith(families: string[], defaultFamily = "Calibri"): DocumentModel {
  const defaultStyles = { ...createDefaultStyles(), fontFamily: defaultFamily };
  return {
    id: "d",
    title: "t",
    defaultStyles,
    sections: [
      {
        blocks: [
          {
            type: "paragraph",
            formatting: {},
            runs: families.map((family) => ({ text: "x", formatting: { fontFamily: family } })),
          },
        ],
      },
    ],
  } as unknown as DocumentModel;
}

const nothingInstalled = { isInstalled: () => false };

describe("fontStatusOf", () => {
  it("marks a font with a bundled twin as substituted, naming the twin (AC-8)", () => {
    expect(fontStatusOf("Calibri", nothingInstalled)).toEqual({
      family: "Calibri",
      status: "substituted",
      substitute: "Carlito",
    });
    expect(fontStatusOf("Cambria", nothingInstalled).substitute).toBe("Caladea");
  });

  it("marks an installed font as available", () => {
    expect(fontStatusOf("Calibri", { isInstalled: (f) => f === "Calibri" }).status).toBe("available");
  });

  it("marks a bundled family as available even when not installed", () => {
    expect(fontStatusOf("Carlito", nothingInstalled).status).toBe("available");
  });

  it("marks an embedded family as available (AC-11/13)", () => {
    const ctx = { isInstalled: () => false, embeddedFamilies: new Set(["myembedded"]) };
    expect(fontStatusOf("MyEmbedded", ctx).status).toBe("available");
  });

  it("marks a font with no twin and not installed as missing (AC-10)", () => {
    expect(fontStatusOf("Papyrus", nothingInstalled)).toEqual({ family: "Papyrus", status: "missing" });
  });
});

describe("documentFontReport", () => {
  it("lists every family the document uses, sorted, with status (AC-7)", () => {
    const model = docWith(["Papyrus", "Arial"], "Calibri");
    const report = documentFontReport(model, nothingInstalled);
    const families = report.map((e) => e.family);
    // The document default (Calibri) plus both run families, alphabetical.
    expect(families).toEqual(["Arial", "Calibri", "Papyrus"]);
    expect(report.find((e) => e.family === "Arial")?.status).toBe("substituted");
    expect(report.find((e) => e.family === "Papyrus")?.status).toBe("missing");
  });
});
