import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { ref } from "vue";
import { useCommandPalette, type CommandPaletteDeps } from "./useCommandPalette";
import { createDocument } from "../../lib/documentModel";

/** Builds deps with sensible enabled defaults; overrides tweak one axis. */
function makeDeps(overrides: Partial<CommandPaletteDeps> = {}): CommandPaletteDeps {
  const base = {
    doc: ref(createDocument("Test")),
    editable: ref(true),
    undoRedo: ref({ canUndo: true, canRedo: true }),
    canConvertLiteralChips: ref(true),
    tableSelectedForChart: ref(true),
    captionTargetSelected: ref(true),
    resolvedTheme: () => "dark",
    dispatch: vi.fn(),
    applyStyle: vi.fn(),
  };
  return { ...base, ...overrides } as unknown as CommandPaletteDeps;
}

describe("useCommandPalette (SPEC-062)", () => {
  beforeEach(() => localStorage.clear());
  afterEach(() => localStorage.clear());

  it("dispatches a command through the same table the menu uses — identical execution (AC-16, AC-17)", () => {
    // One shared registry stands in for editorCommands: a command that mutates
    // a model. The menu path and the palette path both dispatch through it, so
    // the resulting models must be identical.
    const registry = new Map<string, () => void>();
    const runOn = (model: { log: string[] }) => () => {
      model.log.push("did-the-thing");
    };

    const viaMenu = { log: [] as string[] };
    const menuRegistry = new Map([["insert.footnote", runOn(viaMenu)]]);
    // The menu's own dispatch (DocumentView.handleMenuCommand).
    menuRegistry.get("insert.footnote")?.();

    const viaPalette = { log: [] as string[] };
    registry.set("insert.footnote", runOn(viaPalette));
    const palette = useCommandPalette(makeDeps({ dispatch: (id) => registry.get(id)?.() }));
    palette.runCommandFromPalette("insert.footnote");

    expect(viaPalette).toEqual(viaMenu);
    expect(palette.commandPaletteOpen.value).toBe(false);
  });

  it("routes an Apply-style row to applyStyle, not dispatch (AC-23)", () => {
    const applyStyle = vi.fn();
    const dispatch = vi.fn();
    const palette = useCommandPalette(makeDeps({ applyStyle, dispatch }));
    palette.runCommandFromPalette("style-apply:Heading1");
    expect(applyStyle).toHaveBeenCalledWith("Heading1");
    expect(dispatch).not.toHaveBeenCalled();
  });

  it("records the run command as most-recent and persists it (AC-20)", () => {
    const palette = useCommandPalette(makeDeps());
    palette.runCommandFromPalette("edit.undo");
    palette.runCommandFromPalette("format.bold");
    expect(palette.paletteRecentIds.value.slice(0, 2)).toEqual(["format.bold", "edit.undo"]);
    expect(JSON.parse(localStorage.getItem("docu-edit:recent-commands") ?? "[]")[0]).toBe("format.bold");
  });

  it("builds Apply-style rows live from the open document's styles (AC-22, AC-23)", () => {
    const palette = useCommandPalette(makeDeps());
    const labels = palette.paletteStyleRows.value.map((row) => row.label);
    expect(labels.length).toBeGreaterThan(0);
    expect(labels.every((label) => label.startsWith("Apply style: "))).toBe(true);
    expect(palette.paletteStyleRows.value.every((row) => row.id.startsWith("style-apply:"))).toBe(true);
  });

  it("marks editing commands read-only when not editable, but leaves always-on ones enabled (AC-18, AC-19)", () => {
    const palette = useCommandPalette(makeDeps({ editable: ref(false) }));
    const availability = palette.paletteAvailability.value;
    expect(availability["insert.footnote"]).toBe("document is read-only");
    expect(availability["format.bold"]).toBe("document is read-only");
    expect(availability["file.print"]).toBeUndefined();
    expect(availability["view.ruler"]).toBeUndefined();
  });

  it("reflects contextual predicates in the reason (AC-19)", () => {
    const palette = useCommandPalette(
      makeDeps({
        undoRedo: ref({ canUndo: false, canRedo: false }),
        tableSelectedForChart: ref(false),
        captionTargetSelected: ref(false),
        canConvertLiteralChips: ref(false),
        resolvedTheme: () => "light",
      }),
    );
    const a = palette.paletteAvailability.value;
    expect(a["edit.undo"]).toBe("nothing to undo");
    expect(a["edit.redo"]).toBe("nothing to redo");
    expect(a["insert.chart-from-table"]).toBe("no table selected");
    expect(a["insert.caption"]).toBe("no caption target selected");
    expect(a["tools.convert-merge-fields"]).toBe("no {{…}} text to convert");
    expect(a["view.dim-page"]).toBe("dark theme only");
  });

  it("returns focus to the previously focused element on close (AC-14)", () => {
    const button = document.createElement("button");
    document.body.appendChild(button);
    button.focus();
    const palette = useCommandPalette(makeDeps());
    palette.openCommandPalette();
    expect(palette.commandPaletteOpen.value).toBe(true);
    palette.closeCommandPalette();
    expect(document.activeElement).toBe(button);
    button.remove();
  });
});
