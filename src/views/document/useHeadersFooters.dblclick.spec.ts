// SPEC-029 regression: double-clicking a word in a header/footer must leave
// the word selected. Both dblclick paths went straight to
// enterHeaderFooterEditing, which unconditionally parked the caret at
// { block: 0, offset: 0 } on the next tick — so the browser's own word
// selection was wiped and the caret jumped to the start of the first
// paragraph, on the click that enters the band *and* on every later
// double-click inside the band that was already live.
import { nextTick, ref, shallowRef } from "vue";
import { beforeEach, describe, expect, it, vi } from "vitest";
import { useHeadersFooters } from "./useHeadersFooters";
import {
  createDocument,
  sectionMargins,
  sectionPageSize,
  type DocumentModel,
} from "../../lib/documentModel";
import type { PageLayout } from "../../lib/pagination";

const HEADER_TEXT = "Draft report header";

function model(): DocumentModel {
  const doc = createDocument("d");
  doc.sections[0].blocks = [{ type: "paragraph", runs: [{ text: "body", formatting: {} }] }];
  doc.sections[0].headers = {
    default: { blocks: [{ type: "paragraph", runs: [{ text: HEADER_TEXT, formatting: {} }] }] },
  };
  return doc;
}

/** The band as DocPage renders it: one paragraph carrying the synthetic
 * single-section model's coordinates. `data-hf-editable` is what hfRootEl
 * looks for, and the composable's nextTick work runs against it. */
function harness() {
  const doc = shallowRef<DocumentModel | null>(model());
  const root = document.createElement("div");
  const band = document.createElement("div");
  band.setAttribute("data-hf-editable", "1");
  const paragraph = document.createElement("p");
  paragraph.className = "doc-paragraph";
  paragraph.setAttribute("data-section", "0");
  paragraph.setAttribute("data-block", "0");
  paragraph.textContent = HEADER_TEXT;
  band.appendChild(paragraph);
  root.appendChild(band);
  document.body.appendChild(root);
  const page: PageLayout = {
    sectionIndex: 0,
    pageSize: sectionPageSize(doc.value!, 0),
    margins: sectionMargins(doc.value!, 0),
    items: [{ sectionIndex: 0, blockIndex: 0 }],
  };
  const deps = {
    doc,
    contentRoot: ref<Element | null>(root),
    editable: ref(true),
    documentId: ref("d"),
    currentSelection: () => null,
    applyEdit: vi.fn(),
    getPaginationPages: () => [page],
    getPageNumbers: () => [1],
    getCaretSectionIndex: () => 0,
    pageOfCaret: () => 0,
    clearCellSelection: vi.fn(),
    getClearImageSelection: () => vi.fn(),
    getExitNoteEditing: () => vi.fn(),
    getUpdateFieldInSurface: () => vi.fn(),
    getCaret: () => ({ section: 0, block: 0, offset: 0 }),
    markDirty: vi.fn(),
    refreshUndoRedoState: vi.fn(),
  };
  return {
    band,
    paragraph,
    headers: useHeadersFooters(deps as unknown as Parameters<typeof useHeadersFooters>[0]),
  };
}

/** What the browser's own double-click leaves behind: the word selected. */
function selectWord(paragraph: HTMLElement, word: string): void {
  const node = paragraph.firstChild!;
  const start = HEADER_TEXT.indexOf(word);
  const range = document.createRange();
  range.setStart(node, start);
  range.setEnd(node, start + word.length);
  const selection = document.getSelection()!;
  selection.removeAllRanges();
  selection.addRange(range);
}

function selectedText(): string {
  return document.getSelection()?.toString() ?? "";
}

describe("double-click inside a header/footer band (SPEC-029 AC-1)", () => {
  beforeEach(() => {
    document.body.innerHTML = "";
  });

  it("keeps the word the entering double-click selected", async () => {
    const { band, paragraph, headers } = harness();
    selectWord(paragraph, "report");
    headers.handleDblclickHeader(0, { currentTarget: band } as unknown as MouseEvent);
    await nextTick();
    expect(headers.activeHeaderFooter.value?.part).toBe("header");
    expect(selectedText()).toBe("report");
  });

  it("keeps the word a second double-click selected in the live band", async () => {
    const { band, paragraph, headers } = harness();
    headers.handleDblclickHeader(0, { currentTarget: band } as unknown as MouseEvent);
    await nextTick();
    selectWord(paragraph, "header");
    headers.handleDblclickHeader(0, { currentTarget: band } as unknown as MouseEvent);
    await nextTick();
    expect(selectedText()).toBe("header");
  });

  it("records the selected word as the band's own selection", async () => {
    const { band, paragraph, headers } = harness();
    headers.handleDblclickHeader(0, { currentTarget: band } as unknown as MouseEvent);
    await nextTick();
    selectWord(paragraph, "Draft");
    headers.handleDblclickHeader(0, { currentTarget: band } as unknown as MouseEvent);
    await nextTick();
    expect(headers.currentHfSelection()).toEqual({
      anchor: { section: 0, block: 0, offset: 0 },
      focus: { section: 0, block: 0, offset: 5 },
    });
  });

  it("still parks the caret at the band start when entering with no selection", async () => {
    const { band, headers } = harness();
    document.getSelection()?.removeAllRanges();
    headers.handleDblclickHeader(0, { currentTarget: band } as unknown as MouseEvent);
    await nextTick();
    expect(selectedText()).toBe("");
    expect(headers.currentHfSelection()).toEqual({
      anchor: { section: 0, block: 0, offset: 0 },
      focus: { section: 0, block: 0, offset: 0 },
    });
  });
});
