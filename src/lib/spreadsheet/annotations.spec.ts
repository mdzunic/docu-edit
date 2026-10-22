// SPEC-135's annotation commands and rules (annotations.ts): every command,
// copy-on-write, dirty marking per format, the open rule's scheme table,
// HYPERLINK() detection and its follow-time target, remapAnnotations, the
// structural, move and sort adjusters, the rename/delete remap of place
// locations and the cross-format copy.
import { describe, expect, it } from "vitest";
import { replyParagraphText } from "../core/comments";
import type { CellRef, GridRange } from "./address";
import {
  addNote,
  addThread,
  canonicalPlaceLocation,
  crossAnnotations,
  deleteNote,
  deleteThreadMessage,
  editNote,
  editThreadMessage,
  ensurePerson,
  externalHref,
  hyperlinkAt,
  hyperlinkFormulaCall,
  hyperlinkFormulaTarget,
  markAnnotationsEdited,
  noteAt,
  openHyperlinkTarget,
  remapAnnotations,
  removeHyperlink,
  replyToThread,
  setHyperlink,
  setNoteVisible,
  setThreadResolved,
  threadAt,
} from "./annotations";
import type { Cell, CellHyperlink, CellNote, CellThread, Row, Sheet, SpreadsheetFormat, SpreadsheetModel } from "./model";
import { moveRange } from "./move";
import { deleteSheet, renameSheet, stripRetention } from "./sheets";
import { sortRange } from "./sort";
import { applyStructuralEdit } from "./structure";

function sheet(id: string, name: string, fields: Partial<Sheet> = {}): Sheet {
  return {
    conditionalFormats: [],
    validations: [],
    id,
    name,
    visibility: "visible",
    kind: "grid",
    rows: [],
    columns: [],
    defaultColWidthPx: 64,
    defaultRowHeightPt: 15,
    merges: [],
    partName: `xl/worksheets/${id}.xml`,
    ...fields,
  };
}

function model(sheets: Sheet[], format: SpreadsheetFormat = "xlsx"): SpreadsheetModel {
  return {
    names: [],
    dxfs: [],
    kind: "spreadsheet",
    id: "m",
    title: "T",
    format,
    readerVersion: 15,
    sheets,
    activeSheetIndex: 0,
    dateSystem: "1900",
    opaque: [],
    styles: [{}],
    defaultStyleId: 0,
  };
}

const ref = (row: number, col: number): CellRef => ({ row, col });
const range = (top: number, left: number, bottom: number, right: number): GridRange => ({ top, left, bottom, right });
const value = (col: number, v: Cell["value"], extra: Partial<Cell> = {}): Cell => ({ col, value: v, ...extra });
const row = (index: number, cells: Cell[]): Row => ({ index, cells });
const sheetOf = (m: SpreadsheetModel, id = "s1"): Sheet => m.sheets.find((each) => each.id === id)!;
const note = (noteId: string, at: CellRef, extra: Partial<CellNote> = {}): CellNote => ({ noteId, ref: at, author: "Ada", text: "Check", ...extra });
const link = (linkId: string, at: GridRange, target: CellHyperlink["target"] = { kind: "url", href: "https://example.com" }): CellHyperlink => ({ linkId, range: at, target });
const thread = (threadId: string, at: CellRef, resolved = false): CellThread => ({
  threadId,
  ref: at,
  resolved,
  messages: [
    { messageId: threadId, author: "Grace", text: "Why?", createdAt: "2026-10-01T09:00:00.000Z", personId: "{P1}", source: "<threadedComment/>", mentions: "<mentions/>" },
    { messageId: `${threadId}-r`, author: "Ada", text: "Because.", createdAt: "2026-10-01T10:00:00.000Z", personId: "{P2}", mentions: "<mentions/>" },
  ],
});

describe("note commands", () => {
  it("adds a note copy-on-write, sorted by (row, col), and marks an .xlsx sheet annotationsDirty — never dirty", () => {
    const first = note("n-C3", ref(2, 2));
    const m = model([sheet("s1", "Sheet1", { notes: [first] })]);
    const before = sheetOf(m);
    const added = addNote(m, "s1", ref(0, 1), "Revenue source", "You")!;
    expect(added).toMatchObject({ noteId: "n-B1", ref: ref(0, 1), author: "You", text: "Revenue source" });
    const after = sheetOf(m);
    expect(after).not.toBe(before);
    expect(after.notes).not.toBe(before.notes);
    expect(before.notes).toEqual([first]);
    expect(after.notes!.map((each) => each.noteId)).toEqual(["n-B1", "n-C3"]);
    expect(after.notes![1]).toBe(first);
    expect(after.annotationsDirty).toBe(true);
    expect(after.dirty).toBeUndefined();
  });

  it("marks an .ods sheet dirty (its annotations live in the table XML)", () => {
    const m = model([sheet("s1", "Sheet1")], "ods");
    addNote(m, "s1", ref(0, 0), "Hello", "You");
    expect(sheetOf(m).dirty).toBe(true);
    expect(sheetOf(m).annotationsDirty).toBeUndefined();
  });

  it("keeps one note per cell: adding to a noted cell edits it, and a blank note adds nothing", () => {
    const m = model([sheet("s1", "Sheet1", { notes: [note("n-A1", ref(0, 0))] })]);
    addNote(m, "s1", ref(0, 0), "Second thoughts", "You");
    expect(sheetOf(m).notes).toHaveLength(1);
    expect(noteAt(sheetOf(m), ref(0, 0))!.text).toBe("Second thoughts");
    const untouched = model([sheet("s1", "Sheet1")]);
    const sheets = untouched.sheets;
    expect(addNote(untouched, "s1", ref(1, 1), "   ", "You")).toBeNull();
    expect(untouched.sheets).toBe(sheets);
  });

  it("mints a free id when a moved note already holds the cell's read-time id", () => {
    const m = model([sheet("s1", "Sheet1", { notes: [note("n-A1", ref(1, 0))] })]);
    expect(addNote(m, "s1", ref(0, 0), "New", "You")!.noteId).toBe("n-A1-2");
  });

  it("an edit drops the note's source and rich runs; an emptied note is deleted", () => {
    const rich = note("n-A1", ref(0, 0), { runs: [{ text: "Ada:", font: { bold: true } }, { text: " check" }], text: "Ada: check", source: { xlsxComment: "<comment/>", vmlShape: "<v:shape/>" }, visible: true, box: { widthPx: 200, heightPx: 100 } });
    const m = model([sheet("s1", "Sheet1", { notes: [rich] })]);
    const edited = editNote(m, "s1", "n-A1", "Plain now")!;
    expect(edited).toEqual({ noteId: "n-A1", ref: ref(0, 0), author: "Ada", text: "Plain now", visible: true, box: { widthPx: 200, heightPx: 100 } });
    expect(editNote(m, "s1", "n-A1", "  \n ")).toBeNull();
    expect(sheetOf(m).notes).toBeUndefined();
  });

  it("deletes a note and shows / hides one, regenerating its shape but keeping its comment text", () => {
    const m = model([sheet("s1", "Sheet1", { notes: [note("n-A1", ref(0, 0), { source: { xlsxComment: "<comment/>", vmlShape: "<v:shape/>" } }), note("n-B1", ref(0, 1))] })]);
    expect(setNoteVisible(m, "s1", "n-A1", true)).toBe(true);
    expect(noteAt(sheetOf(m), ref(0, 0))).toMatchObject({ visible: true, source: { xlsxComment: "<comment/>" } });
    expect(noteAt(sheetOf(m), ref(0, 0))!.source!.vmlShape).toBeUndefined();
    expect(setNoteVisible(m, "s1", "n-A1", true)).toBe(false);
    expect(setNoteVisible(m, "s1", "n-A1", false)).toBe(true);
    expect(noteAt(sheetOf(m), ref(0, 0))!.visible).toBeUndefined();
    expect(deleteNote(m, "s1", "n-B1")).toBe(true);
    expect(deleteNote(m, "s1", "n-B1")).toBe(false);
    expect(sheetOf(m).notes!.map((each) => each.noteId)).toEqual(["n-A1"]);
  });
});

describe("thread commands", () => {
  it("starts a thread with a GUID id and a new person, replacing — never mutating — model.persons", () => {
    const m = model([sheet("s1", "Sheet1")]);
    const persons = [{ personId: "{P1}", displayName: "Grace" }];
    m.persons = persons;
    const started = addThread(m, "s1", ref(1, 1), "Where from?", "You")!;
    expect(started.threadId).toMatch(/^\{[0-9A-F-]{36}\}$/);
    expect(started.messages[0]).toMatchObject({ messageId: started.threadId, author: "You", text: "Where from?" });
    expect(Date.parse(started.messages[0]!.createdAt)).not.toBeNaN();
    expect(persons).toHaveLength(1);
    expect(m.persons).not.toBe(persons);
    expect(m.persons![1]).toMatchObject({ displayName: "You", userId: "You", providerId: "None", personId: started.messages[0]!.personId });
    expect(threadAt(sheetOf(m), ref(1, 1))).toEqual(started);
    expect(sheetOf(m).annotationsDirty).toBe(true);
  });

  it("finds a known author's person instead of adding one", () => {
    const m = model([sheet("s1", "Sheet1")]);
    m.persons = [{ personId: "{P1}", displayName: "Grace" }];
    const persons = m.persons;
    expect(ensurePerson(m, "Grace").personId).toBe("{P1}");
    expect(m.persons).toBe(persons);
  });

  it("replies, edits (losing source and mentions), resolves and reopens", () => {
    const m = model([sheet("s1", "Sheet1", { threads: [thread("{T1}", ref(5, 1))] })]);
    const replied = replyToThread(m, "s1", "{T1}", "Thanks!", "Grace")!;
    expect(replied.messages.map((each) => each.text)).toEqual(["Why?", "Because.", "Thanks!"]);
    const edited = editThreadMessage(m, "s1", "{T1}", "{T1}", "Why, though?")!;
    expect(edited.messages[0]).toEqual({ messageId: "{T1}", author: "Grace", text: "Why, though?", createdAt: "2026-10-01T09:00:00.000Z", personId: "{P1}" });
    expect(edited.messages[1]!.mentions).toBe("<mentions/>");
    expect(setThreadResolved(m, "s1", "{T1}", true)).toBe(true);
    expect(threadAt(sheetOf(m), ref(5, 1))!.resolved).toBe(true);
    expect(setThreadResolved(m, "s1", "{T1}", true)).toBe(false);
    expect(setThreadResolved(m, "s1", "{T1}", false)).toBe(true);
    expect(threadAt(sheetOf(m), ref(5, 1))!.resolved).toBe(false);
  });

  it("resolving drops only the root message's source (its done flag changes)", () => {
    const m = model([sheet("s1", "Sheet1", { threads: [thread("{T1}", ref(0, 0))] })]);
    setThreadResolved(m, "s1", "{T1}", true);
    const [root, reply] = threadAt(sheetOf(m), ref(0, 0))!.messages;
    expect(root!.source).toBeUndefined();
    expect(root!.mentions).toBe("<mentions/>");
    expect(reply).toEqual(thread("{T1}", ref(0, 0)).messages[1]);
  });

  it("deleting a reply keeps the thread; deleting the opening message deletes it", () => {
    const m = model([sheet("s1", "Sheet1", { threads: [thread("{T1}", ref(0, 0))] })]);
    expect(deleteThreadMessage(m, "s1", "{T1}", "{T1}-r")!.messages).toHaveLength(1);
    expect(deleteThreadMessage(m, "s1", "{T1}", "{T1}")).toBeNull();
    expect(sheetOf(m).threads).toBeUndefined();
  });

  it("adding to a cell that holds a thread replies to it", () => {
    const m = model([sheet("s1", "Sheet1", { threads: [thread("{T1}", ref(0, 0))] })]);
    addThread(m, "s1", ref(0, 0), "More", "You");
    expect(sheetOf(m).threads).toHaveLength(1);
    expect(sheetOf(m).threads![0]!.messages).toHaveLength(3);
  });
});

describe("hyperlink commands", () => {
  it("sets one link over a range, replacing every link it intersects, and keeps the id of the same range", () => {
    const m = model([sheet("s1", "Sheet1", { hyperlinks: [link("{L1}", range(0, 0, 0, 0)), link("h-C3", range(2, 2, 3, 3)), link("h-F1", range(0, 5, 0, 5))] })]);
    const set = setHyperlink(m, "s1", range(0, 0, 2, 2), { kind: "place", location: "Data!A10" }, "Go");
    expect(set).toEqual({ linkId: "h-A1", range: range(0, 0, 2, 2), target: { kind: "place", location: "Data!A10" }, tooltip: "Go" });
    expect(sheetOf(m).hyperlinks!.map((each) => each.linkId)).toEqual(["h-A1", "h-F1"]);
    const again = setHyperlink(m, "s1", range(0, 5, 0, 5), { kind: "url", href: "https://example.org" });
    expect(again.linkId).toBe("h-F1");
    expect(hyperlinkAt(sheetOf(m), ref(1, 1))!.linkId).toBe("h-A1");
    expect(sheetOf(m).annotationsDirty).toBe(true);
  });

  it("removes a link", () => {
    const m = model([sheet("s1", "Sheet1", { hyperlinks: [link("h-A1", range(0, 0, 0, 0))] })]);
    expect(removeHyperlink(m, "s1", "h-A1")).toBe(true);
    expect(sheetOf(m).hyperlinks).toBeUndefined();
    expect(removeHyperlink(m, "s1", "h-A1")).toBe(false);
  });
});

describe("openHyperlinkTarget — which targets a sheet opens", () => {
  const cases: [string, ReturnType<typeof openHyperlinkTarget>][] = [
    ["https://example.com/contract", "external"],
    ["HTTP://EXAMPLE.COM", "external"],
    ["mailto:team@example.com", "external"],
    ["javascript:alert(1)", "refused"],
    ["  java\tscript:alert(1)", "refused"],
    ["data:text/html,<b>x</b>", "refused"],
    ["vbscript:msgbox", "refused"],
    ["file:///C:/secret.txt", "refused"],
    ["ftp://example.com/file", "refused"],
    ["Book2.xlsx", "refused"],
    ["../docs/report.pdf", "refused"],
    ["[Book2.xlsx]Sheet1!A1", "refused"],
  ];
  for (const [href, verdict] of cases) {
    it(`${JSON.stringify(href)} → ${verdict}`, () => {
      expect(openHyperlinkTarget({ kind: "url", href })).toBe(verdict);
    });
  }

  it("a place is the view's to resolve", () => {
    expect(openHyperlinkTarget({ kind: "place", location: "#REF!" })).toBe("place");
  });

  it("an external href carries its subaddress", () => {
    expect(externalHref({ kind: "url", href: "https://example.com/doc", subaddress: "part2" })).toBe("https://example.com/doc#part2");
  });
});

describe("place locations", () => {
  it("are canonical reference text", () => {
    expect(canonicalPlaceLocation("Data!a10")).toBe("Data!A10");
    expect(canonicalPlaceLocation("'My Sheet'!$A$1:B2")).toBe("'My Sheet'!$A$1:B2");
    expect(canonicalPlaceLocation("MyName")).toBe("MyName");
    expect(canonicalPlaceLocation("#REF!")).toBe("#REF!");
  });
});

describe("HYPERLINK() cells", () => {
  it("detects only a top-level HYPERLINK call", () => {
    expect(hyperlinkFormulaCall({ formula: 'HYPERLINK("https://example.com","Site")' }, ref(0, 0))?.name).toBe("HYPERLINK");
    expect(hyperlinkFormulaCall({ formula: 'IF(A1,HYPERLINK("https://example.com"),"")' }, ref(0, 0))).toBeNull();
    expect(hyperlinkFormulaCall({ formula: "SUM(A1:A3)" }, ref(0, 0))).toBeNull();
    expect(hyperlinkFormulaCall({ formula: "HYPERLINK(", formulaUnparsed: true }, ref(0, 0))).toBeNull();
    expect(hyperlinkFormulaCall({}, ref(0, 0))).toBeNull();
  });

  it("computes the target at follow time from the first argument, coerced to text", () => {
    const cell = value(1, "Site", { formula: 'HYPERLINK(A1&"/contract","Site")' });
    const m = model([sheet("s1", "Sheet1", { rows: [row(0, [value(0, "https://example.com"), cell])] })]);
    expect(hyperlinkFormulaTarget(m, "s1", ref(0, 1), cell)).toEqual({ kind: "url", href: "https://example.com/contract" });
    const place = value(0, "Go", { formula: 'HYPERLINK("#Data!A10","Go")' });
    expect(hyperlinkFormulaTarget(m, "s1", ref(3, 0), place)).toEqual({ kind: "place", location: "Data!A10" });
    const blank = value(0, "", { formula: 'HYPERLINK(Z9,"Nothing")' });
    expect(hyperlinkFormulaTarget(m, "s1", ref(4, 0), blank)).toBeNull();
  });

  it("is never stored as a hyperlink", () => {
    const cell = value(0, "Site", { formula: 'HYPERLINK("https://example.com","Site")' });
    const m = model([sheet("s1", "Sheet1", { rows: [row(0, [cell])] })]);
    hyperlinkFormulaTarget(m, "s1", ref(0, 0), cell);
    expect(sheetOf(m).hyperlinks).toBeUndefined();
  });
});

describe("remapAnnotations", () => {
  it("applies an arbitrary mapping, deleting on null and keeping a moved note's source", () => {
    const moved = note("n-A1", ref(0, 0), { source: { xlsxComment: "<comment/>", vmlShape: "<v:shape/>" } });
    const m = model([sheet("s1", "Sheet1", { notes: [moved, note("n-B1", ref(0, 1))], threads: [thread("{T}", ref(1, 0))], hyperlinks: [link("h-A3", range(2, 0, 2, 0))] })]);
    remapAnnotations(m, "s1", (at) => (at.col === 1 ? null : { row: at.row + 10, col: at.col }));
    const s = sheetOf(m);
    expect(s.notes).toEqual([{ ...moved, ref: ref(10, 0) }]);
    expect(s.threads![0]!.ref).toEqual(ref(11, 0));
    expect(s.hyperlinks![0]!.range).toEqual(range(12, 0, 12, 0));
    expect(s.annotationsDirty).toBe(true);
  });

  it("leaves a sheet with nothing to move untouched by identity", () => {
    const m = model([sheet("s1", "Sheet1", { notes: [note("n-A1", ref(0, 0))] })]);
    const before = m.sheets[0];
    remapAnnotations(m, "s1", (at) => at);
    expect(m.sheets[0]).toBe(before);
  });

  it("moves a range link only onto a range of the same size", () => {
    const m = model([sheet("s1", "Sheet1", { hyperlinks: [link("h-A1", range(0, 0, 1, 1))] })]);
    remapAnnotations(m, "s1", (at) => (at.row === 0 ? { row: 5, col: at.col } : at));
    expect(sheetOf(m).hyperlinks![0]!.range).toEqual(range(0, 0, 1, 1));
  });
});

describe("annotations follow their cells", () => {
  it("a row inserted above moves notes, threads and links; a place location in another sheet follows", () => {
    const m = model([
      sheet("s1", "Summary", { notes: [note("n-B2", ref(1, 1))], threads: [thread("{T}", ref(5, 1))], hyperlinks: [link("h-E5", range(4, 4, 5, 5))] }),
      sheet("s2", "Data", { hyperlinks: [link("h-A1", range(0, 0, 0, 0), { kind: "place", location: "Summary!B2" })] }),
    ]);
    expect(applyStructuralEdit(m, { kind: "insertRows", sheetId: "s1", at: 1, count: 1 }).ok).toBe(true);
    expect(sheetOf(m, "s1").notes![0]!.ref).toEqual(ref(2, 1));
    expect(sheetOf(m, "s1").threads![0]!.ref).toEqual(ref(6, 1));
    expect(sheetOf(m, "s1").hyperlinks![0]!.range).toEqual(range(5, 4, 6, 5));
    expect(sheetOf(m, "s2").hyperlinks![0]!.target).toEqual({ kind: "place", location: "Summary!B3" });
    expect(sheetOf(m, "s2").annotationsDirty).toBe(true);
  });

  it("a column deleted through an annotation removes it; a range link straddling it shrinks", () => {
    const m = model([sheet("s1", "Sheet1", { notes: [note("n-B1", ref(0, 1)), note("n-D1", ref(0, 3))], hyperlinks: [link("h-A3", range(2, 0, 2, 2)), link("h-B5", range(4, 1, 4, 1))] })]);
    applyStructuralEdit(m, { kind: "deleteCols", sheetId: "s1", at: 1, count: 1 });
    const s = sheetOf(m);
    expect(s.notes!.map((each) => [each.noteId, each.ref])).toEqual([["n-D1", ref(0, 2)]]);
    expect(s.hyperlinks!.map((each) => [each.linkId, each.range])).toEqual([["h-A3", range(2, 0, 2, 1)]]);
  });

  it("a place whose cells were deleted keeps #REF!", () => {
    const m = model([sheet("s1", "Summary"), sheet("s2", "Data", { hyperlinks: [link("h-A1", range(0, 0, 0, 0), { kind: "place", location: "Summary!B2" })] })]);
    applyStructuralEdit(m, { kind: "deleteRows", sheetId: "s1", at: 1, count: 1 });
    expect(sheetOf(m, "s2").hyperlinks![0]!.target).toEqual({ kind: "place", location: "Summary!#REF!" });
  });

  it("cut-paste carries a note, a thread and a link with their cells, onto another sheet too", () => {
    const m = model([
      sheet("s1", "Sheet1", { rows: [row(0, [value(0, "x")])], notes: [note("n-A1", ref(0, 0))], threads: [thread("{T}", ref(1, 1))], hyperlinks: [link("h-A1", range(0, 0, 1, 1))] }),
      sheet("s2", "Sheet2"),
    ]);
    expect(moveRange(m, { fromSheetId: "s1", from: range(0, 0, 1, 1), toSheetId: "s2", toTop: 4, toLeft: 2 }).ok).toBe(true);
    expect(sheetOf(m, "s1").notes).toBeUndefined();
    expect(sheetOf(m, "s1").threads).toBeUndefined();
    expect(sheetOf(m, "s1").hyperlinks).toBeUndefined();
    expect(sheetOf(m, "s2").notes![0]!.ref).toEqual(ref(4, 2));
    expect(sheetOf(m, "s2").threads![0]!.ref).toEqual(ref(5, 3));
    expect(sheetOf(m, "s2").hyperlinks![0]!.range).toEqual(range(4, 2, 5, 3));
  });

  it("a destination annotation is overwritten; a link only partly overlapping stays", () => {
    const m = model([sheet("s1", "Sheet1", { notes: [note("n-A1", ref(0, 0)), note("n-K11", ref(10, 10))], hyperlinks: [link("h-J10", range(9, 9, 12, 12)), link("h-K12", range(11, 10, 11, 10))] })]);
    moveRange(m, { fromSheetId: "s1", from: range(0, 0, 1, 1), toSheetId: "s1", toTop: 10, toLeft: 10 });
    const s = sheetOf(m);
    expect(s.notes!.map((each) => [each.noteId, each.ref])).toEqual([["n-A1", ref(10, 10)]]);
    expect(s.hyperlinks!.map((each) => each.linkId)).toEqual(["h-J10"]);
  });

  it("a place pointing into the moved block follows it", () => {
    const m = model([sheet("s1", "Sheet1", { rows: [row(0, [value(0, "x")])] }), sheet("s2", "Sheet2", { hyperlinks: [link("h-A1", range(0, 0, 0, 0), { kind: "place", location: "Sheet1!A1" })] })]);
    moveRange(m, { fromSheetId: "s1", from: range(0, 0, 0, 0), toSheetId: "s1", toTop: 9, toLeft: 0 });
    expect(sheetOf(m, "s2").hyperlinks![0]!.target).toEqual({ kind: "place", location: "Sheet1!A10" });
  });

  it("a sort carries notes, threads and links with their records", () => {
    const m = model([
      sheet("s1", "Sheet1", {
        rows: [row(0, [value(0, 3)]), row(1, [value(0, 1)]), row(2, [value(0, 2)])],
        notes: [note("n-A1", ref(0, 0))],
        threads: [thread("{T}", ref(1, 0))],
        hyperlinks: [link("h-A3", range(2, 0, 2, 0))],
      }),
    ]);
    const sorted = sortRange(m, "s1", range(0, 0, 2, 0), { keys: [{ col: 0, descending: false }], hasHeader: false, caseSensitive: false, localeTag: "en-US" });
    expect(sorted.ok).toBe(true);
    const s = sheetOf(m);
    expect(s.notes![0]!.ref).toEqual(ref(2, 0));
    expect(s.threads![0]!.ref).toEqual(ref(0, 0));
    expect(s.hyperlinks![0]!.range).toEqual(range(1, 0, 1, 0));
  });
});

describe("sheet rename and delete", () => {
  const places = (): SpreadsheetModel =>
    model([
      sheet("s1", "Summary", { hyperlinks: [link("h-A1", range(0, 0, 0, 0), { kind: "place", location: "Data!A10" })] }),
      sheet("s2", "Data", { notes: [note("n-A1", ref(0, 0))] }),
    ]);

  it("a rename follows every place location", () => {
    const m = places();
    renameSheet(m, 1, "Figures");
    expect(sheetOf(m, "s1").hyperlinks![0]!.target).toEqual({ kind: "place", location: "Figures!A10" });
  });

  it("a delete turns its places into #REF! and takes the sheet's own annotations with it", () => {
    const m = places();
    deleteSheet(m, 1);
    expect(sheetOf(m, "s1").hyperlinks![0]!.target).toEqual({ kind: "place", location: "#REF!" });
    expect(m.sheets.flatMap((each) => each.notes ?? [])).toEqual([]);
  });
});

describe("crossing formats", () => {
  const source = (): Pick<Sheet, "notes" | "threads" | "hyperlinks"> => ({
    notes: [note("n-A1", ref(0, 0), { runs: [{ text: "Check", font: { bold: true } }], source: { xlsxComment: "<comment/>" }, visible: true })],
    threads: [thread("{T}", ref(1, 0), true)],
    hyperlinks: [{ ...link("{L}", range(2, 0, 2, 0)), source: { relId: "rId3" }, tooltip: "Go" }],
  });

  it("to .ods a thread becomes a note — the opening message then one replyParagraphText paragraph per reply, resolved dropped", () => {
    const crossed = crossAnnotations(source(), "ods");
    const [opening, reply] = thread("{T}", ref(1, 0)).messages;
    expect(crossed.threads).toBeUndefined();
    expect(crossed.notes).toEqual([
      { noteId: "n-A1", ref: ref(0, 0), author: "Ada", text: "Check", visible: true },
      { noteId: "n-A2", ref: ref(1, 0), author: "Grace", createdAt: opening!.createdAt, text: `Why?\n${replyParagraphText(reply!)}` },
    ]);
    expect(crossed.hyperlinks).toEqual([{ ...link("{L}", range(2, 0, 2, 0)), tooltip: "Go" }]);
  });

  it("to .xlsx notes, threads and links cross field for field without their sources", () => {
    const crossed = crossAnnotations(source(), "xlsx");
    expect(crossed.notes![0]).toEqual({ noteId: "n-A1", ref: ref(0, 0), author: "Ada", text: "Check", visible: true, runs: [{ text: "Check", font: { bold: true } }] });
    expect(crossed.threads![0]!.messages[0]!.source).toBeUndefined();
    expect(crossed.threads![0]!.resolved).toBe(true);
  });

  it("a thread on a noted cell joins that note", () => {
    const crossed = crossAnnotations({ notes: [note("n-A2", ref(1, 0))], threads: [thread("{T}", ref(1, 0))] }, "ods");
    expect(crossed.notes).toHaveLength(1);
    expect(crossed.notes![0]!.text.split("\n")[0]).toBe("Check");
  });

  it("stripRetention drops every annotation source and annotationsDirty, and the persons for an .ods copy", () => {
    const m = model([sheet("s1", "Sheet1", { ...source(), annotationsDirty: true })]);
    m.persons = [{ personId: "{P1}", displayName: "Grace", source: "<person/>" }];
    const ods = stripRetention({ ...m, format: "ods" });
    expect(ods.persons).toBeUndefined();
    expect(ods.sheets[0]!.annotationsDirty).toBeUndefined();
    expect(JSON.stringify(ods.sheets[0])).not.toContain("source");
    const xlsx = stripRetention(m);
    expect(xlsx.persons).toEqual([{ personId: "{P1}", displayName: "Grace" }]);
    expect(xlsx.sheets[0]!.threads).toHaveLength(1);
  });
});

describe("markAnnotationsEdited", () => {
  it("installs the sheet once marked, reusing an already-marked sheet", () => {
    const m = model([sheet("s1", "Sheet1", { annotationsDirty: true })]);
    const marked = markAnnotationsEdited(m, m.sheets[0]!);
    expect(marked).toBe(m.sheets[0]);
  });
});
