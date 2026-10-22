// SPEC-053 AC-10 field shading vs. SPEC-047 AC-16 (the open phase's longest
// main-thread block): pagedViews must not depend on the raw caret mirror.
// It did — renderEntryFor read `caretPosition.value` for every body paragraph —
// so the session-restore caret placement right after the first layout rebuilt
// every realized page's entries and re-rendered every DocPage inside the same
// task as the first layout (a second full page pass, measured at ~+25ms of the
// open's longest block on the SPEC-047 fixture), and every keystroke in plain
// text paid the same full-page pass again. The caret may only reach the page
// render narrowed to "inside a body paragraph that holds a shadeable field".
import { computed, nextTick, ref, shallowRef } from "vue";
import { describe, expect, it } from "vitest";
import { usePaginationView, type PaginationViewDeps } from "./usePaginationView";
import {
  createDocument,
  createFieldSpecRun,
  createChipRun,
  sectionMargins,
  sectionPageSize,
  type DocumentModel,
  type Paragraph,
  type Run,
} from "../../lib/documentModel";
import type { PageLayout } from "../../lib/pagination";
import type { CaretPosition } from "../../lib/textEditing";

function text(t: string): Run {
  return { text: t, formatting: {} };
}

/** Paragraph 0: plain text. Paragraph 1: text + a date field + text.
 * Paragraph 2: text + a mail-merge chip (a field that never shades). */
function model(): DocumentModel {
  const doc = createDocument("d");
  const blocks: Paragraph[] = [
    { type: "paragraph", runs: [text("plain paragraph")] },
    {
      type: "paragraph",
      runs: [text("dated "), createFieldSpecRun({ kind: "date", format: "d MMMM yyyy", cached: "1 January 2020" }), text(" end")],
    },
    { type: "paragraph", runs: [text("chip "), createChipRun({ kind: "merge", name: "First" })] },
  ];
  doc.sections[0].blocks = blocks;
  return doc;
}

function harness() {
  const doc = shallowRef<DocumentModel | null>(model());
  const caretPosition = ref<CaretPosition | null>(null);
  const fieldShadingOn = ref(false);
  const page: PageLayout = {
    sectionIndex: 0,
    pageSize: sectionPageSize(doc.value!, 0),
    margins: sectionMargins(doc.value!, 0),
    items: [0, 1, 2].map((blockIndex) => ({ sectionIndex: 0, blockIndex })),
  };
  const deps = {
    doc,
    editable: ref(true),
    pagination: {
      pages: shallowRef([page]),
      noteAreas: shallowRef({ footnotes: [], endnotes: [] }),
      isMeasured: () => false,
      lineOffsetsFor: () => null,
      tabLayoutFor: () => null,
    },
    noteLabelsMap: ref(new Map()),
    noteTooltipsMap: ref(new Map()),
    pageNumbers: ref([1]),
    marksOn: ref(false),
    virtualizing: ref(false),
    viewportPageRange: ref({ start: 0, end: 0 }),
    caretPage: ref(1),
    anchorPage: ref(0),
    findMatchPage: ref(-1),
    activeHeaderFooter: ref(null),
    activeNote: ref(null),
    cellSelection: ref(null),
    imageSelection: ref(null),
    shapeSelection: ref(null),
    shapeMultiSelection: ref([]),
    editingShapeText: ref(false),
    chartSelection: ref(null),
    collapsedHeadings: ref(new Set<string>()),
    remount: ref({ key: "", count: 0, epoch: 0 }),
    realizedSticky: new WeakSet<object>(),
    hfRemountEpoch: ref(0),
    noteRemountEpoch: ref(0),
    tocPageNumberFor: () => null,
    trackChangesDisplayMode: ref("allMarkup"),
    fieldShadingOn,
    caretPosition,
  };
  // Only the fields pagedViews reads are supplied (the *Deps convention).
  const { pagedViews } = usePaginationView(deps as unknown as PaginationViewDeps);
  // A dependent computed, so "recomputed" is observable as a new identity.
  const views = computed(() => pagedViews.value);
  return { views, caretPosition, fieldShadingOn, doc };
}

function shaded(views: ReturnType<typeof harness>["views"], blockIndex: number): boolean {
  const entry = views.value[0].entries[blockIndex];
  return entry.runs?.some((run) => run.atomClass?.includes("doc-field-shaded")) ?? false;
}

describe("usePaginationView — the page render only tracks the caret through a field paragraph", () => {
  it("a caret move within plain text does not rebuild the page views", async () => {
    const { views, caretPosition } = harness();
    const before = views.value;
    caretPosition.value = { section: 0, block: 0, offset: 3 };
    await nextTick();
    expect(views.value).toBe(before);
    caretPosition.value = { section: 0, block: 0, offset: 9 };
    await nextTick();
    expect(views.value).toBe(before);
  });

  it("a caret move onto a merge chip (a field that never shades) does not rebuild them either", async () => {
    const { views, caretPosition } = harness();
    const before = views.value;
    caretPosition.value = { section: 0, block: 2, offset: 5 };
    await nextTick();
    expect(views.value).toBe(before);
  });

  it("the caret landing on a field run still shades it, and leaving it unshades (AC-10)", async () => {
    const { views, caretPosition } = harness();
    expect(shaded(views, 1)).toBe(false);
    caretPosition.value = { section: 0, block: 1, offset: 6 };
    await nextTick();
    expect(shaded(views, 1)).toBe(true);
    caretPosition.value = { section: 0, block: 0, offset: 0 };
    await nextTick();
    expect(shaded(views, 1)).toBe(false);
  });

  it("a caret inside a table cell never reaches the body render", async () => {
    const { views, caretPosition } = harness();
    const before = views.value;
    caretPosition.value = { section: 0, block: 1, offset: 6, cell: { row: 0, col: 0, block: 0 } };
    await nextTick();
    expect(views.value).toBe(before);
  });

  it("View → Field shading still repaints every field without a caret (AC-11)", async () => {
    const { views, fieldShadingOn } = harness();
    expect(shaded(views, 1)).toBe(false);
    fieldShadingOn.value = true;
    await nextTick();
    expect(shaded(views, 1)).toBe(true);
  });
});
