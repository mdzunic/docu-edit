import { describe, expect, it } from "vitest";
import { importStyles, planStyleImport, type StyleImportResolution } from "./customStyles";
import { createDocument, type DocumentModel, type NamedStyle } from "./documentModel";

// SPEC-042 AC-19..24 — the import merge plan and its application.

function fixture(styles: NamedStyle[]): DocumentModel {
  const model = createDocument("import-fixture");
  model.styles = styles;
  return model;
}

describe("planStyleImport (AC-21, AC-23)", () => {
  it("splits source styles into conflicting (same name) and non-conflicting", () => {
    const model = fixture([{ id: "Shared", name: "Shared Name" }]);
    const source: NamedStyle[] = [
      { id: "S1", name: "Shared Name" },
      { id: "S2", name: "Unique Name" },
    ];
    const plan = planStyleImport(model, source);
    expect(plan.nonConflicting.map((s) => s.id)).toEqual(["S2"]);
    expect(plan.conflicts).toHaveLength(1);
    expect(plan.conflicts[0].existing.id).toBe("Shared");
  });

  it("matches names case-insensitively and trimmed", () => {
    const model = fixture([{ id: "Shared", name: "  Shared Name " }]);
    const plan = planStyleImport(model, [{ id: "S1", name: "shared name" }]);
    expect(plan.conflicts).toHaveLength(1);
  });
});

describe("importStyles (AC-22, AC-23, AC-24)", () => {
  it("adds non-conflicting styles under fresh ids and summarises them as added", () => {
    const model = fixture([]);
    const source: NamedStyle[] = [{ id: "S1", name: "Imported Unique Style" }];
    const { model: next, summary } = importStyles(model, source, []);
    expect(summary.added).toEqual(["Imported Unique Style"]);
    expect(next.styles?.some((s) => s.name === "Imported Unique Style")).toBe(true);
  });

  it("replaces the existing style in place, keeping its id, when resolution is replace", () => {
    const model = fixture([{ id: "Existing", name: "Shared Name", runFormatting: { bold: true } }]);
    const source: NamedStyle[] = [{ id: "S1", name: "Shared Name", runFormatting: { italic: true } }];
    const resolutions: StyleImportResolution[] = [{ sourceId: "S1", resolution: "replace" }];
    const { model: next, summary } = importStyles(model, source, resolutions);
    expect(summary.replaced).toEqual(["Shared Name"]);
    const replaced = next.styles?.find((s) => s.id === "Existing");
    expect(replaced?.runFormatting).toEqual({ italic: true });
  });

  it("keeps the existing style unchanged and reports it skipped when resolution is keep", () => {
    const model = fixture([{ id: "Existing", name: "Shared Name", runFormatting: { bold: true } }]);
    const source: NamedStyle[] = [{ id: "S1", name: "Shared Name", runFormatting: { italic: true } }];
    const resolutions: StyleImportResolution[] = [{ sourceId: "S1", resolution: "keep" }];
    const { model: next, summary } = importStyles(model, source, resolutions);
    expect(summary.skipped).toEqual(["Shared Name"]);
    expect(next.styles?.find((s) => s.id === "Existing")?.runFormatting).toEqual({ bold: true });
  });

  it("skips a conflicting style entirely when resolution is skip", () => {
    const model = fixture([{ id: "Existing", name: "Shared Name" }]);
    const source: NamedStyle[] = [{ id: "S1", name: "Shared Name" }];
    const resolutions: StyleImportResolution[] = [{ sourceId: "S1", resolution: "skip" }];
    const { summary } = importStyles(model, source, resolutions);
    expect(summary.skipped).toEqual(["Shared Name"]);
  });

  it("remaps a based-on reference between two imported styles to their fresh ids", () => {
    const model = fixture([]);
    const source: NamedStyle[] = [
      { id: "S1", name: "Base" },
      { id: "S2", name: "Derived", basedOn: "S1" },
    ];
    const { model: next } = importStyles(model, source, []);
    const base = next.styles?.find((s) => s.name === "Base");
    const derived = next.styles?.find((s) => s.name === "Derived");
    expect(derived?.basedOn).toBe(base?.id);
  });

  it("never mints two styles with the same id", () => {
    const model = fixture([{ id: "Imported", name: "Existing With That Id" }]);
    const source: NamedStyle[] = [{ id: "Imported", name: "Different Name" }];
    const { model: next } = importStyles(model, source, []);
    const ids = next.styles?.map((s) => s.id) ?? [];
    expect(new Set(ids).size).toBe(ids.length);
  });
});
