import { describe, it, expect, vi } from "vitest";
import { ref, shallowRef } from "vue";
import type { PresentationModel } from "../../lib/presentation/model";
import { createBlankPresentation } from "../../lib/presentation/slides";
import { useDeckRedesign } from "./useDeckRedesign";
import { usePresentationMenu } from "./usePresentationMenu";

// The menu is a second route to commands the button strips already carry, so
// what matters is that every id lands on the composable the strip calls — not
// that the command works, which its own spec covers. A routing table is worth
// testing precisely because it is easy to wire an id to the wrong handler and
// never notice: both paths look right in the UI until someone uses the menu.

function stubs() {
  const calls: string[] = [];
  const spy = (name: string) => vi.fn(() => calls.push(name));
  const panel = (name: string, open = false) => ({ open: ref(open), toggle: spy(`${name}.toggle`), show: spy(`${name}.show`) });
  const deps = {
    undo: { canUndo: ref(false), canRedo: ref(false), undo: spy("undo"), redo: spy("redo") },
    deckActions: {
      canPresent: ref(true),
      otherFormatLabel: ref("OpenDocument (.odp)"),
      download: spy("download"),
      downloadCopy: spy("downloadCopy"),
      openHtmlExport: spy("exportHtml"),
      present: spy("present"),
      close: spy("close"),
    },
    deckSetup: { open: spy("deckSetup") },
    notes: panel("notes"),
    transitions: panel("transitions"),
    animations: panel("animations"),
    layouts: panel("layouts"),
    tables: panel("tables"),
    media: panel("media"),
    headerFooter: panel("headerFooter"),
    themePanel: panel("themePanel"),
    charts: { pickChartType: vi.fn((kind: string) => calls.push(`chart:${kind}`)) },
    preferences: { open: spy("preferences") },
    stageEdit: {
      arrange: {
        selectedCount: ref(0),
        canUngroup: ref(false),
        group: spy("group"),
        ungroup: spy("ungroup"),
        openPosition: spy("position"),
        align: vi.fn((v: string) => calls.push(`align:${v}`)),
      },
      editing: {
        toggleBold: spy("bold"),
        toggleItalic: spy("italic"),
        toggleUnderline: spy("underline"),
        toggleBulletCmd: spy("bullet"),
        indent: vi.fn((n: number) => calls.push(`indent:${n}`)),
        setAlign: vi.fn((v: string) => calls.push(`setAlign:${v}`)),
      },
      insertion: {
        armTextBox: spy("textBox"),
        armShape: vi.fn((kind: string) => calls.push(`shape:${kind}`)),
      },
      manipulation: { arrange: vi.fn((v: string) => calls.push(`arrange:${v}`)) },
    },
    addSlide: spy("addSlide"),
  };
  return { deps, calls };
}

describe("usePresentationMenu routing", () => {
  const cases: Array<[string, string]> = [
    ["file.export", "download"],
    ["file.export-copy", "downloadCopy"],
    ["file.export-html", "exportHtml"],
    ["file.deck-setup", "deckSetup"],
    ["app.preferences", "preferences"],
    ["file.close", "close"],
    ["edit.undo", "undo"],
    ["edit.redo", "redo"],
    ["view.design", "themePanel.toggle"],
    ["view.transition", "transitions.toggle"],
    ["view.animate", "animations.toggle"],
    ["view.layout", "layouts.toggle"],
    ["view.table", "tables.toggle"],
    ["view.media", "media.toggle"],
    ["view.header-footer", "headerFooter.show"],
    ["view.notes", "notes.toggle"],
    ["view.present", "present"],
    ["insert.slide", "addSlide"],
    ["format.group", "group"],
    ["format.ungroup", "ungroup"],
    ["format.position", "position"],
  ];

  it.each(cases)("routes %s to %s", (command, expected) => {
    const { deps, calls } = stubs();
    usePresentationMenu(deps).run(command);
    expect(calls).toEqual([expected]);
  });

  // `arrange:` is the one parameterised id the bar still emits — the shape,
  // chart and alignment forms went with the Insert and Format menus, which the
  // ribbon's two strips replaced (they call the composables directly).
  it("routes the parameterised id with its argument", () => {
    const { deps, calls } = stubs();
    const menu = usePresentationMenu(deps);
    menu.run("arrange:front");
    expect(calls).toEqual(["arrange:front"]);
  });

  // A suffix form the bar no longer emits must not quietly still work — that
  // is how a deleted menu leaves a live back door behind it.
  it("ignores the parameterised ids the ribbon took over", () => {
    const { deps, calls } = stubs();
    const menu = usePresentationMenu(deps);
    menu.run("insert.shape:ellipse");
    menu.run("insert.chart:pie");
    menu.run("format.align:center");
    menu.run("align:left");
    expect(calls).toEqual([]);
  });

  it("does nothing for an id it does not know", () => {
    const { deps, calls } = stubs();
    const menu = usePresentationMenu(deps);
    menu.run("nonsense");
    menu.run("nonsense:with-argument");
    expect(calls).toEqual([]);
  });
});

describe("usePresentationMenu state", () => {
  it("reports what is checked and what is available", () => {
    const { deps } = stubs();
    const { state } = usePresentationMenu(deps);
    expect(state.value.canUndo).toBe(false);
    expect(state.value.canPresent).toBe(true);
    expect(state.value.otherFormatLabel).toBe("OpenDocument (.odp)");
    expect(state.value.designOpen).toBe(false);
    expect(state.value.hasSelection).toBe(false);
    expect(state.value.canGroup).toBe(false);
  });

  it("tracks the panels and the selection as they change", () => {
    const { deps } = stubs();
    const { state } = usePresentationMenu(deps);

    deps.themePanel.open.value = true;
    deps.undo.canUndo.value = true;
    expect(state.value.designOpen).toBe(true);
    expect(state.value.canUndo).toBe(true);

    // Group needs more than one frame; a single selection is not enough.
    deps.stageEdit.arrange.selectedCount.value = 1;
    expect(state.value.hasSelection).toBe(true);
    expect(state.value.canGroup).toBe(false);

    deps.stageEdit.arrange.selectedCount.value = 2;
    expect(state.value.canGroup).toBe(true);
  });
});

// SPEC-115: Slide › Redesign… is rendered by the deck-designer experiment's
// gate; the menu only routes its id to the view's redesign controller and
// reports whether there is anything to arrange.
describe("Slide › Redesign… (SPEC-115)", () => {
  it("routes slide.redesign to the redesign controller, which a deck with no slides refuses", () => {
    const { deps } = stubs();
    const model = shallowRef<PresentationModel | null>(createBlankPresentation());
    const redesign = useDeckRedesign({ model, currentIndex: ref(0), record: vi.fn(), commit: vi.fn() });
    const menu = usePresentationMenu({ ...deps, redesign });
    expect(menu.state.value.canRedesign).toBe(true);
    menu.run("slide.redesign");
    expect(redesign.open.value).toBe(true);
    redesign.close();

    model.value = { ...createBlankPresentation(), slides: [] };
    expect(menu.state.value.canRedesign).toBe(false);
    menu.run("slide.redesign");
    expect(redesign.open.value).toBe(false);
  });

  it("is unavailable when the view wires no redesign controller", () => {
    const { deps } = stubs();
    const menu = usePresentationMenu(deps);
    expect(menu.state.value.canRedesign).toBe(false);
    expect(() => menu.run("slide.redesign")).not.toThrow();
  });
});
