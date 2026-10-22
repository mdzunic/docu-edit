// SPEC-135 AC "plain text everywhere": every annotation text — a note's, a
// thread message's, an author's — renders as plain text, never markup. The
// security test feeds `<img src=x onerror=…>` through each surface that shows
// one — the hover popover (note and thread), the note box, the comments pane
// (threads and an .ods workbook's notes) and the indicator's accessible
// description — and asserts no element is created from it, the literal text
// is what shows, and no v-html appears in any of those components.
import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { mount } from "@vue/test-utils";
import { afterEach, describe, expect, it, vi } from "vitest";
import { computed, effectScope, nextTick, ref, shallowRef, triggerRef, type EffectScope } from "vue";
import { NOTE_HOVER_DELAY_MS } from "../../lib/spreadsheet/annotations";
import type { Sheet, SpreadsheetFormat, SpreadsheetModel } from "../../lib/spreadsheet/model";
import { createBlankWorkbook } from "../../lib/spreadsheet/sheets";
import { NOTE_DESCRIPTION_ID, useCellNotes } from "../../views/spreadsheet/useCellNotes";
import { useSheetComments } from "../../views/spreadsheet/useSheetComments";
import { useSheetSidePanel } from "../../views/spreadsheet/useSheetSidePanel";
import { useSpreadsheetUndo } from "../../views/spreadsheet/useSpreadsheetUndo";
import SheetCommentsPanel from "./SheetCommentsPanel.vue";
import SheetNoteLayer from "./SheetNoteLayer.vue";
import SheetNotePopover from "./SheetNotePopover.vue";

const HOSTILE = '<img src=x onerror="window.__spec135Pwned = true">';
const AUTHOR = "<b>Mallory</b>";

const scopes: EffectScope[] = [];
afterEach(() => {
  scopes.splice(0).forEach((scope) => scope.stop());
  vi.useRealTimers();
  localStorage.clear();
});

function workbook(format: SpreadsheetFormat): SpreadsheetModel {
  const model = createBlankWorkbook();
  model.format = format;
  const sheet: Sheet = {
    ...model.sheets[0]!,
    id: "s1",
    name: "Summary",
    notes: [{ noteId: "n-B2", ref: { row: 1, col: 1 }, author: AUTHOR, text: HOSTILE, visible: true }],
    threads:
      format === "xlsx"
        ? [{ threadId: "{T}", ref: { row: 1, col: 3 }, resolved: false, messages: [{ messageId: "{T}", personId: "p", author: AUTHOR, text: HOSTILE, createdAt: "2026-03-01T09:00:00Z" }] }]
        : undefined,
  };
  model.sheets = [sheet];
  return model;
}

function controllers(format: SpreadsheetFormat = "xlsx") {
  const wb = shallowRef<SpreadsheetModel | null>(workbook(format));
  const activeSheet = computed(() => wb.value?.sheets[0] ?? null);
  const active = ref({ row: 1, col: 1 });
  const selection = computed(() => ({ active: active.value, ranges: [{ top: active.value.row, left: active.value.col, bottom: active.value.row, right: active.value.col }], activeRangeIndex: 0 }));
  const scope = effectScope();
  scopes.push(scope);
  return {
    active,
    ...scope.run(() => {
      const undo = useSpreadsheetUndo({ model: wb, recalc: (_keys, done) => done([]), onApplied: () => triggerRef(wb), activateSheet: () => {} });
      const sidePanel = useSheetSidePanel();
      return {
        notes: useCellNotes({ book: { workbook: wb, activeSheet }, undo, selection: { selection } }),
        comments: useSheetComments({ book: { workbook: wb, activeSheet, activeIndex: ref(0) }, undo, selection: { selection }, sidePanel }),
      };
    })!,
  };
}

/** The surface shows the hostile text literally and built nothing from it. */
function expectInert(root: Element, text = HOSTILE): void {
  expect(root.querySelector("img")).toBeNull();
  expect(root.querySelector("b")).toBeNull();
  expect(root.textContent).toContain(text);
  expect(Reflect.get(window, "__spec135Pwned")).toBeUndefined();
}

const anchor = () => ({ cell: { left: 64, top: 20, width: 64, height: 20 }, above: 20 });

describe("annotation texts are plain text (SPEC-135 security)", () => {
  it("the hover popover shows a note's and a thread's hostile text and author literally", () => {
    const wrapper = mount(SheetNotePopover, {
      props: { note: { author: AUTHOR, text: HOSTILE }, thread: { author: AUTHOR, text: HOSTILE, replies: 0, resolved: false }, address: "B2" },
    });
    expectInert(wrapper.element);
    expect(wrapper.get('[data-testid="sheet-note-popover-text"]').text()).toBe(HOSTILE);
    expect(wrapper.get('[data-testid="sheet-thread-popover-text"]').text()).toBe(HOSTILE);
    expect(wrapper.get('[data-testid="sheet-note-popover-author"]').text()).toBe(AUTHOR);
  });

  it("the note box, the layer's popover and the indicator's accessible description show it literally", async () => {
    vi.useFakeTimers();
    const { notes, active } = controllers();
    const wrapper = mount(SheetNoteLayer, { props: { controller: notes, anchor, bounds: { width: 800, height: 600 } } });
    const box = wrapper.get('[data-testid="sheet-note-box"]');
    expectInert(box.element);
    const description = wrapper.get(`#${NOTE_DESCRIPTION_ID}`);
    expectInert(description.element);
    expect(description.text()).toBe(`Note by ${AUTHOR}: ${HOSTILE}`);
    active.value = { row: 1, col: 3 };
    notes.hoverCell({ row: 1, col: 3 });
    vi.advanceTimersByTime(NOTE_HOVER_DELAY_MS);
    await nextTick();
    expectInert(wrapper.get('[data-testid="sheet-thread-popover"]').element);
    expectInert(wrapper.get(`#${NOTE_DESCRIPTION_ID}`).element);
    wrapper.unmount();
  });

  it("the note editor holds the hostile text as the textarea's value, never as markup", async () => {
    const { notes } = controllers();
    notes.openEditor({ row: 1, col: 1 });
    const wrapper = mount(SheetNoteLayer, { props: { controller: notes, anchor, bounds: { width: 800, height: 600 } }, attachTo: document.body });
    await nextTick();
    const editor = wrapper.get('[data-testid="sheet-note-editor"]');
    expect((editor.element as HTMLTextAreaElement).value).toBe(HOSTILE);
    expect(wrapper.element.querySelector("img")).toBeNull();
    wrapper.unmount();
  });

  it("the comments pane shows a thread's and an .ods note's hostile text and author literally", () => {
    const xlsx = controllers("xlsx");
    const threads = mount(SheetCommentsPanel, { props: { controller: xlsx.comments } });
    expectInert(threads.get('[data-testid="sheet-comment-thread-{T}"]').element);
    expect(threads.get('[data-testid="sheet-comment-text"]').text()).toBe(HOSTILE);
    expect(threads.get('[data-testid="sheet-comment-author"]').text()).toBe(AUTHOR);
    const ods = controllers("ods");
    const notes = mount(SheetCommentsPanel, { props: { controller: ods.comments } });
    expectInert(notes.get('[data-testid="sheet-comment-note-n-B2"]').element);
  });

  it("none of the annotation components uses v-html", () => {
    const here = dirname(fileURLToPath(import.meta.url));
    for (const name of ["SheetNotePopover.vue", "SheetNoteLayer.vue", "SheetCommentsPanel.vue", "SheetCellLayer.vue"]) {
      expect(readFileSync(join(here, name), "utf8"), name).not.toMatch(/v-html|innerHTML/);
    }
  });
});
