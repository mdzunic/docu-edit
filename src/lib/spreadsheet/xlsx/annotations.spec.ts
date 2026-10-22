// SPEC-135's xlsx wire mapping (xlsx/annotations.ts, threadPlaceholder.ts,
// vmlNoteGolden.ts): reading every field of spec135-annotations.xlsx, the
// out-of-subset preserve path, the writer rules — a no-edit download and a
// forced rebuild byte-identical part for part, placeholders regenerated, the
// VML part spliced with the checkbox shape kept byte for byte, parts minted
// and dropped, relationships appended after the highest id and an
// unreferenced hyperlink relationship left — the goldens, the round trip
// (add/edit/reply/resolve/link → save → reopen field for field, write twice
// byte-stable), the sanctioned deltas, a duplicated, a deleted and a
// cross-format sheet, and the human-captured Excel 365 workbook's cases
// (skipped until it is committed).
import { existsSync, readFileSync } from "node:fs";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { equalBytes } from "../../core/equality";
import { parseXml } from "../../core/xml";
import { unzip } from "../../core/zip";
import { zip } from "../../core/zipWriter";
import { equalStructural, firstStructuralDifference } from "../../equality";
import { addNote, addThread, deleteNote, deleteThreadMessage, editNote, editThreadMessage, removeHyperlink, replyToThread, setHyperlink, setNoteVisible, setThreadResolved } from "../annotations";
import { setCellEntry } from "../cells";
import { buildSpec121LinksXlsxBytes, buildSpec135AnnotationsOdsBytes, buildSpec135AnnotationsXlsxBytes, SPEC135_ANNOTATIONS_XLSX } from "../fixtures/build";
import { packageIntegrity } from "../fixtures/packageIntegrity";
import type { Sheet, SpreadsheetModel } from "../model";
import { readSpreadsheetBytes } from "../open";
import { deleteSheet, duplicateSheet, stripRetention } from "../sheets";
import { applyStructuralEdit } from "../structure";
import { readXlsxPersons, readXlsxSheetAnnotations, spliceWorksheetAnnotations } from "./annotations";
import { placeholderAuthor, placeholderCommentXml, placeholderText, THREAD_PLACEHOLDER_PREAMBLE } from "./threadPlaceholder";
import { defaultNoteAnchor, vmlDrawingXml, vmlNoteShapeFor } from "./vmlNoteGolden";
import { writeXlsx } from "./writer";

const F = SPEC135_ANNOTATIONS_XLSX;
const decoder = new TextDecoder();
const encoder = new TextEncoder();
const EXCEL365 = "e2e/fixtures/spec135-excel365-threads.xlsx";
const EXCEL365_PATH = join(__dirname, "../../../..", EXCEL365);
const EXCEL365_SKIP = "fixture not yet captured: spec135-excel365-threads.xlsx";
const SUMMARY_RELS = F.sheets.Summary.rels;

const open = (bytes: Uint8Array, name: string = F.file): Promise<SpreadsheetModel> => readSpreadsheetBytes(bytes, "xlsx", name);
const textOf = (parts: Map<string, Uint8Array>, name: string): string => decoder.decode(parts.get(name));
const summaryOf = (model: SpreadsheetModel): Sheet => model.sheets[0]!;

async function fixture(): Promise<{ bytes: Uint8Array; parts: Map<string, Uint8Array>; model: SpreadsheetModel }> {
  const bytes = await buildSpec135AnnotationsXlsxBytes();
  return { bytes, parts: await unzip(bytes), model: await open(bytes) };
}

/** A sheet's annotations field for field — ids and retention left out. */
function fields(sheet: Sheet): unknown {
  return {
    notes: (sheet.notes ?? []).map(({ noteId: _id, source: _source, ...note }) => note),
    threads: (sheet.threads ?? []).map(({ threadId: _id, messages, ...thread }) => ({
      ...thread,
      messages: messages.map(({ messageId: _message, personId: _person, source: _source, ...message }) => message),
    })),
    hyperlinks: (sheet.hyperlinks ?? []).map(({ linkId: _id, source: _source, ...link }) => link),
  };
}

/** The names of the parts whose bytes differ (or exist on one side only). */
function changedParts(a: Map<string, Uint8Array>, b: Map<string, Uint8Array>): string[] {
  const names = new Set([...a.keys(), ...b.keys()]);
  return [...names].filter((name) => !a.has(name) || !b.has(name) || !equalBytes(a.get(name)!, b.get(name)!)).toSorted();
}

/** The `v:shape` elements of a VML text, in order. */
const vmlShapes = (vml: string): string[] => [...vml.matchAll(/<v:shape [\s\S]*?<\/v:shape>/g)].map((match) => match[0]);
const checkboxShape = (vml: string): string => vmlShapes(vml).find((shape) => shape.includes(`ObjectType="Checkbox"`))!;

/** The model with `annotationsDirty` forced on its first sheet. */
function forceRebuild(model: SpreadsheetModel): void {
  model.sheets = model.sheets.map((sheet, index) => (index === 0 ? { ...sheet, annotationsDirty: true } : sheet));
}

// --- reading -----------------------------------------------------------------------------------

describe("reading spec135-annotations.xlsx", () => {
  it("reads both notes: author, rich runs (the bold author first), text, visibility, box and their source elements", async () => {
    const { model, parts } = await fixture();
    const notes = summaryOf(model).notes!;
    expect(notes.map((note) => note.noteId)).toEqual(["n-B2", "n-B4"]);
    const [b2, b4] = notes;
    expect(b2).toMatchObject({ ref: { row: 1, col: 1 }, author: F.notes.B2.author, text: F.notes.B2.text });
    expect(b2!.visible).toBeUndefined();
    expect(b2!.box).toBeUndefined();
    expect(b2!.runs!.map((run) => [run.text, run.font?.bold === true])).toEqual(F.notes.B2.runs.map((run) => [run.text, run.bold]));
    expect(b2!.runs![0]!.font).toMatchObject({ family: "Tahoma", sizePt: 9 });
    expect(b4).toMatchObject({ ref: { row: 3, col: 1 }, author: F.notes.B4.author, text: F.notes.B4.text, visible: true, box: F.notes.B4.box });
    expect(b4!.runs).toBeUndefined();
    const comments = textOf(parts, F.comments.partName);
    const vml = textOf(parts, F.vml.partName);
    for (const note of notes) {
      expect(comments).toContain(note.source!.xlsxComment!);
      expect(note.source!.xlsxComment!.startsWith(`<comment ref="${note.noteId.slice(2)}"`)).toBe(true);
      expect(vmlShapes(vml)).toContain(note.source!.vmlShape!);
    }
  });

  it("reads both threads with their messages, authors through the persons, dates and resolved state — and no placeholder as a note", async () => {
    const { model } = await fixture();
    expect((model.persons ?? []).map(({ source: _source, ...person }) => person)).toEqual(F.persons);
    const threads = summaryOf(model).threads!;
    for (const [ref, thread] of [
      ["D2", threads[0]!],
      ["E2", threads[1]!],
    ] as const) {
      const T = F.threads[ref];
      expect(thread.threadId).toBe(T.threadId);
      expect(thread.resolved).toBe(T.resolved);
      expect(thread.messages.map((message) => [message.messageId, message.author, message.text, message.createdAt, message.personId])).toEqual(
        T.messages.map((message) => [message.id, message.author, message.text, message.dT, message.personId]),
      );
      expect(thread.messages.every((message) => message.source?.startsWith("<threadedComment ") && message.mentions === undefined)).toBe(true);
    }
    expect(summaryOf(model).notes!.some((note) => note.author.startsWith("tc="))).toBe(false);
  });

  it("reads every hyperlink: the range, url targets through their relationships, the place, tooltip, display and xr:uid ids", async () => {
    const { model } = await fixture();
    expect(summaryOf(model).hyperlinks).toEqual([
      { linkId: F.hyperlinks[0].uid, range: { top: 5, left: 0, bottom: 5, right: 1 }, target: { kind: "url", href: F.hyperlinks[0].target }, source: { relId: "rId1" } },
      {
        linkId: F.hyperlinks[1].uid,
        range: { top: 6, left: 0, bottom: 6, right: 0 },
        target: { kind: "url", href: F.hyperlinks[1].target },
        tooltip: F.hyperlinks[1].tooltip,
        source: { relId: "rId2" },
      },
      { linkId: F.hyperlinks[2].uid, range: { top: 6, left: 1, bottom: 6, right: 1 }, target: { kind: "place", location: "Data!A1" }, display: "See data" },
      { linkId: F.hyperlinks[3].uid, range: { top: 7, left: 0, bottom: 7, right: 0 }, target: { kind: "url", href: "mailto:team@example.com" }, source: { relId: "rId3" } },
      { linkId: F.hyperlinks[4].uid, range: { top: 7, left: 1, bottom: 7, right: 1 }, target: { kind: "url", href: "javascript:alert(1)" }, source: { relId: "rId4" } },
    ]);
    expect(model.sheets[1]!.notes).toBeUndefined();
    expect(model.sheets[1]!.hyperlinks).toBeUndefined();
  });

  it("sets nothing dirty, and the reader-upgrade entry points read the same from the parts", async () => {
    const { model, parts } = await fixture();
    expect(model.sheets.some((sheet) => sheet.dirty || sheet.annotationsDirty)).toBe(false);
    const persons = readXlsxPersons(parts);
    expect(persons).toEqual(model.persons);
    const read = readXlsxSheetAnnotations(parts, F.sheets.Summary.partName, textOf(parts, F.sheets.Summary.partName));
    const summary = summaryOf(model);
    expect(read).toEqual({ notes: summary.notes, threads: summary.threads, hyperlinks: summary.hyperlinks });
  });

  it("an `h-${A1}` id for a link without xr:uid, `n-${A1}` for every note, the same on every read", async () => {
    const { parts } = await fixture();
    const sheet = textOf(parts, F.sheets.Summary.partName).replace(/ xr:uid="\{9C0D[^"]*\}"/g, "");
    const read = readXlsxSheetAnnotations(parts, F.sheets.Summary.partName, sheet);
    expect(read.hyperlinks!.map((link) => link.linkId)).toEqual(["h-A6", "h-A7", "h-B7", "h-A8", "h-B8"]);
  });
});

// --- the subset ----------------------------------------------------------------------------------

/** The fixture with one part's text replaced. */
async function withPart(name: string, edit: (xml: string) => string): Promise<Uint8Array> {
  const parts = await unzip(await buildSpec135AnnotationsXlsxBytes());
  parts.set(name, encoder.encode(edit(textOf(parts, name))));
  return zip([...parts].map(([partName, data]) => ({ name: partName, data })));
}

describe("the out-of-subset rule: annotations never half-parse", () => {
  const cases: [string, string, (xml: string) => string][] = [
    ["two comments on one cell", F.comments.partName, (xml) => xml.replace(`<comment ref="B4"`, `<comment ref="B2"`)],
    ["an unresolvable author", F.comments.partName, (xml) => xml.replace(`<comment ref="B4" authorId="0"`, `<comment ref="B4" authorId="9"`)],
    ["unparseable VML", F.vml.partName, (xml) => xml.replace("</x:ClientData>", "</x:ClientDatum>")],
    ["a thread author no person names", F.threadedComments.partName, (xml) => xml.replace(`personId="${F.persons[0]!.personId}"`, `personId="{00000000-0000-4000-8000-000000000000}"`)],
  ];
  for (const [label, part, edit] of cases) {
    it(`${label}: the sheet has no notes or threads, and a save after a cell edit and a link edit keeps every comment part byte-verbatim`, async () => {
      const bytes = await withPart(part, edit);
      const source = await unzip(bytes);
      const model = await open(bytes);
      const summary = summaryOf(model);
      expect(summary.notes).toBeUndefined();
      expect(summary.threads).toBeUndefined();
      expect(summary.hyperlinks).toHaveLength(5);
      setCellEntry(model, 0, { row: 0, col: 0 }, { value: "Metric (edited)" });
      setHyperlink(model, summary.id, { top: 6, left: 2, bottom: 6, right: 2 }, { kind: "place", location: "Data!A10" });
      const written = await unzip(await writeXlsx(model));
      for (const name of [F.comments.partName, F.vml.partName, F.threadedComments.partName, F.person.partName]) {
        expect(equalBytes(written.get(name)!, source.get(name)!), name).toBe(true);
      }
      expect(textOf(written, F.sheets.Summary.partName)).toContain(`<legacyDrawing r:id="${F.vml.rId}"/>`);
      expect(textOf(written, F.sheets.Summary.partName)).toContain(`<hyperlink ref="C7" location="Data!A10"/>`);
    });
  }

  it("a hyperlink naming no relationship leaves the links unread and <hyperlinks> verbatim on save", async () => {
    const bytes = await withPart(SUMMARY_RELS, (xml) => xml.replace(/<Relationship Id="rId4"[^>]*\/>/, ""));
    const model = await open(bytes);
    expect(summaryOf(model).hyperlinks).toBeUndefined();
    expect(summaryOf(model).notes).toHaveLength(2);
    addNote(model, summaryOf(model).id, { row: 3, col: 3 }, "x", "You");
    const before = textOf(await unzip(bytes), F.sheets.Summary.partName);
    const hyperlinks = /<hyperlinks>[\s\S]*<\/hyperlinks>/.exec(before)![0];
    expect(textOf(await unzip(await writeXlsx(model)), F.sheets.Summary.partName)).toContain(hyperlinks);
  });
});

// --- writing -----------------------------------------------------------------------------------

describe("writing: what does not change", () => {
  it("a no-edit download is byte-identical part for part", async () => {
    const { bytes, model } = await fixture();
    expect(changedParts(await unzip(bytes), await unzip(await writeXlsx(model)))).toEqual([]);
  });

  it("a rebuild with nothing changed re-emits every annotation part, the worksheet and its .rels byte for byte (placeholders regenerated identically)", async () => {
    const { parts, model } = await fixture();
    forceRebuild(model);
    const written = await unzip(await writeXlsx(model));
    for (const name of [F.comments.partName, F.vml.partName, F.threadedComments.partName, F.person.partName, SUMMARY_RELS, F.sheets.Summary.partName, "[Content_Types].xml"]) {
      expect(equalBytes(written.get(name)!, parts.get(name)!), name).toBe(true);
    }
  });

  it("a dirty sheet (a cell edited) keeps every annotation part byte for byte", async () => {
    const { parts, model } = await fixture();
    setCellEntry(model, 0, { row: 2, col: 0 }, { value: "Margin" });
    const written = await unzip(await writeXlsx(model));
    for (const name of [F.comments.partName, F.vml.partName, F.threadedComments.partName, F.person.partName, SUMMARY_RELS]) {
      expect(equalBytes(written.get(name)!, parts.get(name)!), name).toBe(true);
    }
    expect(/<hyperlinks>[\s\S]*<\/hyperlinks>/.exec(textOf(written, F.sheets.Summary.partName))![0]).toBe(/<hyperlinks>[\s\S]*<\/hyperlinks>/.exec(textOf(parts, F.sheets.Summary.partName))![0]);
  });
});

describe("writing: the writer rules", () => {
  it("annotationsDirty alone splices only <hyperlinks> (and <legacyDrawing>) into the retained worksheet: sheetData and every other element verbatim", async () => {
    const { parts, model } = await fixture();
    setHyperlink(model, summaryOf(model).id, { top: 6, left: 2, bottom: 6, right: 2 }, { kind: "place", location: "Data!A10" }, "Go to the data");
    const written = textOf(await unzip(await writeXlsx(model)), F.sheets.Summary.partName);
    const source = textOf(parts, F.sheets.Summary.partName);
    const added = `<hyperlink ref="C7" location="Data!A10" tooltip="Go to the data"/>`;
    expect(written).toBe(source.replace(`<hyperlink ref="A8"`, `${added}<hyperlink ref="A8"`));
  });

  it("a new note: the comments part rebuilt with the author appended, the VML spliced (the checkbox byte for byte), a golden shape after the last note shape", async () => {
    const { parts, model } = await fixture();
    addNote(model, summaryOf(model).id, { row: 3, col: 3 }, "Check the costs", "You");
    const written = await unzip(await writeXlsx(model));
    const comments = textOf(written, F.comments.partName);
    expect(comments).toContain(`<authors>${F.authors.map((author) => `<author>${author}</author>`).join("")}<author>You</author></authors>`);
    expect(comments).toContain(`<comment ref="D4" authorId="3" shapeId="0"><text><t>Check the costs</t></text></comment></commentList>`);
    const before = textOf(parts, F.vml.partName);
    const vml = textOf(written, F.vml.partName);
    expect(checkboxShape(vml)).toBe(checkboxShape(before));
    expect(vmlShapes(vml).slice(0, 4)).toEqual(vmlShapes(before).slice(0, 4));
    expect(vmlShapes(vml)[4]).toBe(vmlNoteShapeFor({ row: 3, col: 3 }, { idmap: 1, n: 6, visible: false, metrics: undefined }).replace(/margin-top:[\d.]+pt/, vmlShapes(vml)[4]!.match(/margin-top:[\d.]+pt/)![0]));
    expect(vml).toBe(before.replace(" </v:shape><v:shapetype id=\"_x0000_t201\"", ` </v:shape>${vmlShapes(vml)[4]}<v:shapetype id="_x0000_t201"`));
    expect(await packageIntegrity(await writeXlsx(model))).toEqual([]);
  });

  it("a reply and a resolve: the threaded part keeps every unchanged message verbatim, the placeholder is regenerated from the thread", async () => {
    const { parts, model } = await fixture();
    const thread = summaryOf(model).threads![0]!;
    replyToThread(model, summaryOf(model).id, thread.threadId, "Noted <with> & care", "You");
    setThreadResolved(model, summaryOf(model).id, thread.threadId, true);
    const written = await unzip(await writeXlsx(model));
    const threaded = textOf(written, F.threadedComments.partName);
    for (const message of thread.messages.slice(1)) {
      expect(threaded).toContain(message.source!);
    }
    expect(threaded).toContain(thread.messages[0]!.source!.replace(`id="${thread.threadId}"`, `id="${thread.threadId}" done="1"`));
    expect(threaded).toMatch(/<threadedComment ref="D2" dT="[^"]+Z" personId="\{[0-9A-F-]{36}\}" id="\{[0-9A-F-]{36}\}" parentId="\{7A8B9C0D-1E2F-4A3B-8C4D-5E6F7A8B9C01\}"><text>Noted &lt;with&gt; &amp; care<\/text><\/threadedComment>/);
    const reopened = await open(await writeXlsx(model));
    const regenerated = placeholderCommentXml(summaryOf(reopened).threads![0]!, "D2", 1);
    expect(textOf(written, F.comments.partName)).toContain(regenerated);
    expect(regenerated).toContain(`\nReply:\n    Noted &lt;with&gt; &amp; care</t>`);
    // The persons part gains the new author; the source persons verbatim.
    expect(textOf(written, F.person.partName)).toContain(textOf(parts, F.person.partName).slice(0, textOf(parts, F.person.partName).indexOf("</personList>")));
    expect(textOf(written, F.person.partName)).toMatch(/<person displayName="You" id="\{[0-9A-F-]{36}\}" userId="You" providerId="None"\/><\/personList>$/);
  });

  it("a new external link appends its relationship after the highest rId; an unreferenced hyperlink relationship stays", async () => {
    const { parts, model } = await fixture();
    const sheet = summaryOf(model);
    setHyperlink(model, sheet.id, { top: 8, left: 0, bottom: 8, right: 0 }, { kind: "url", href: "https://new.example/?a=1&b=2" });
    removeHyperlink(model, sheet.id, F.hyperlinks[3].uid);
    const written = await unzip(await writeXlsx(model));
    const rels = textOf(written, SUMMARY_RELS);
    expect(rels).toBe(textOf(parts, SUMMARY_RELS).replace("</Relationships>", `<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://new.example/?a=1&amp;b=2" TargetMode="External"/></Relationships>`));
    const worksheet = textOf(written, F.sheets.Summary.partName);
    expect(worksheet).not.toContain(`<hyperlink ref="A8"`);
    expect(worksheet).toContain(`<hyperlink ref="A9" r:id="rId9"/></hyperlinks>`);
    expect(await packageIntegrity(await writeXlsx(model))).toEqual([]);
  });

  it("the last note and thread gone: the comments and threaded parts leave with their relationships and Overrides, the VML keeps the checkbox and <legacyDrawing> stays", async () => {
    const { parts, model } = await fixture();
    const sheet = summaryOf(model);
    for (const note of sheet.notes!) {
      deleteNote(model, sheet.id, note.noteId);
    }
    for (const thread of sheet.threads!) {
      deleteThreadMessage(model, sheet.id, thread.threadId, thread.threadId);
    }
    const bytes = await writeXlsx(model);
    const written = await unzip(bytes);
    expect(written.has(F.comments.partName)).toBe(false);
    expect(written.has(F.threadedComments.partName)).toBe(false);
    expect(textOf(written, "[Content_Types].xml")).not.toContain("comments1.xml");
    expect(textOf(written, "[Content_Types].xml")).not.toContain("threadedComment1.xml");
    expect(textOf(written, SUMMARY_RELS)).not.toMatch(/comments1\.xml|threadedComment1\.xml/);
    const vml = textOf(written, F.vml.partName);
    expect(vmlShapes(vml)).toEqual([checkboxShape(textOf(parts, F.vml.partName))]);
    expect(textOf(written, F.sheets.Summary.partName)).toContain(`<legacyDrawing r:id="${F.vml.rId}"/>`);
    expect(written.has(F.person.partName)).toBe(true);
    expect(await packageIntegrity(bytes)).toEqual([]);
  });

  it("a sheet's first note mints the comments and VML parts at the next free names, their relationships, Overrides and <legacyDrawing> in CT_Worksheet order", async () => {
    const { model } = await fixture();
    const data = model.sheets[1]!;
    addNote(model, data.id, { row: 9, col: 0 }, "Target row", "You");
    setHyperlink(model, data.id, { top: 0, left: 0, bottom: 0, right: 0 }, { kind: "url", href: "https://example.com/data" });
    const bytes = await writeXlsx(model);
    const written = await unzip(bytes);
    expect(textOf(written, "xl/comments2.xml")).toContain(`<comment ref="A10" authorId="0" shapeId="0"><text><t>Target row</t></text></comment>`);
    expect(textOf(written, "xl/drawings/vmlDrawing2.vml")).toBe(vmlDrawingXml(2, [vmlNoteShapeFor({ row: 9, col: 0 }, { idmap: 2, n: 1, visible: false, metrics: undefined })]).replace(/margin-top:[\d.]+pt/, textOf(written, "xl/drawings/vmlDrawing2.vml").match(/margin-top:[\d.]+pt/)![0]));
    expect(textOf(written, "[Content_Types].xml")).toContain(`<Override PartName="/xl/comments2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/>`);
    const sheet = textOf(written, F.sheets.Data.partName);
    expect(sheet).toMatch(/<\/sheetData><hyperlinks><hyperlink ref="A1" r:id="rId1"\/><\/hyperlinks><pageMargins [^>]*\/><legacyDrawing r:id="rId3"\/><\/worksheet>$/);
    expect(textOf(written, "xl/worksheets/_rels/sheet2.xml.rels")).toContain(`Target="../drawings/vmlDrawing2.vml"`);
    expect(await packageIntegrity(bytes)).toEqual([]);
    const reopened = await open(bytes);
    expect(fields(reopened.sheets[1]!)).toEqual(fields(model.sheets[1]!));
  });
});

describe("writing: the worksheet splice and dangerous schemes", () => {
  it("a javascript: link is kept in the model and the file — its relationship and <hyperlink> survive an annotation edit", async () => {
    const { model } = await fixture();
    expect(summaryOf(model).hyperlinks![4]!.target).toEqual({ kind: "url", href: "javascript:alert(1)" });
    addNote(model, summaryOf(model).id, { row: 3, col: 3 }, "x", "You");
    const written = await unzip(await writeXlsx(model));
    expect(textOf(written, SUMMARY_RELS)).toContain(`<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="javascript:alert(1)" TargetMode="External"/>`);
    expect(textOf(written, F.sheets.Summary.partName)).toContain(`<hyperlink ref="B8" r:id="rId4" xr:uid="${F.hyperlinks[4].uid}"/>`);
  });

  it("splices into a prefixed worksheet with the root's prefixes, declaring r: and xr: locally when the root lacks them", () => {
    const xml =
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><x:worksheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">` +
      `<x:sheetData/><x:pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/><x:tableParts count="0"/></x:worksheet>`;
    const link = { linkId: "{9C0D1E2F-3A4B-4C5D-8E6F-7A8B9C0D1E39}", range: { top: 0, left: 0, bottom: 1, right: 0 }, target: { kind: "url" as const, href: "https://x.example/" } };
    const out = spliceWorksheetAnnotations(xml, [{ link, relId: "rId7" }], "rId8");
    expect(out).toBe(
      xml.replace(
        "<x:pageMargins",
        `<x:hyperlinks xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision">` +
          `<x:hyperlink ref="A1:A2" r:id="rId7" xr:uid="${link.linkId}"/></x:hyperlinks><x:pageMargins`,
      ).replace("<x:tableParts", `<x:legacyDrawing xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId8"/><x:tableParts`),
    );
    expect(spliceWorksheetAnnotations(out, [], null)).toBe(xml);
  });
});

// --- the round trip ------------------------------------------------------------------------------

describe("the round trip", () => {
  it("add/edit/reply/resolve/link → save → reopen reproduces every modelled field; writing the reopened model is byte-stable and reproduces the save", async () => {
    const { model } = await fixture();
    const id = summaryOf(model).id;
    addNote(model, id, { row: 3, col: 3 }, "Line one\nLine two", "You");
    editNote(model, id, "n-B2", "Rewritten");
    setNoteVisible(model, id, "n-B4", false);
    replyToThread(model, id, F.threads.D2.threadId, "Me too", "You");
    setThreadResolved(model, id, F.threads.E2.threadId, false);
    addThread(model, id, { row: 0, col: 1 }, "A new thread", "Pat Doe");
    setHyperlink(model, id, { top: 6, left: 2, bottom: 6, right: 2 }, { kind: "place", location: "Data!A10" }, "Data");
    setHyperlink(model, id, { top: 9, left: 0, bottom: 10, right: 1 }, { kind: "url", href: "https://example.com/x", subaddress: "part" });
    const bytes = await writeXlsx(model);
    expect(await packageIntegrity(bytes)).toEqual([]);
    const reopened = await open(bytes);
    const difference = firstStructuralDifference(fields(summaryOf(reopened)), fields(summaryOf(model)));
    expect(difference).toBeNull();
    expect(equalStructural(fields(summaryOf(reopened)), fields(summaryOf(model)))).toBe(true);
    expect(summaryOf(reopened).notes!.find((note) => note.noteId === "n-B4")!.box).toEqual(F.notes.B4.box);
    // Twice, with every annotation forced through the writer.
    forceRebuild(reopened);
    const first = await writeXlsx(reopened);
    const second = await writeXlsx(reopened);
    expect(equalBytes(first, second)).toBe(true);
    expect(changedParts(await unzip(first), await unzip(bytes))).toEqual([]);
  });

  it("sanctioned: a moved note's box keeps its size but not its custom position (a row inserted above it)", async () => {
    const { parts, model } = await fixture();
    applyStructuralEdit(model, { kind: "insertRows", sheetId: summaryOf(model).id, at: 1, count: 1 });
    const bytes = await writeXlsx(model);
    const written = await unzip(bytes);
    const shapes = vmlShapes(textOf(written, F.vml.partName));
    const b5 = shapes.find((shape) => shape.includes("<x:Row>4</x:Row>"))!;
    expect(b5).toContain("width:150pt;height:75pt");
    expect(b5).toContain(`id="${F.notes.B4.shapeId}"`);
    expect(b5).toContain(`    ${defaultNoteAnchor({ row: 4, col: 1 }).slice(0, 4).join(", ")}, `);
    expect(b5).not.toBe(vmlShapes(textOf(parts, F.vml.partName))[1]);
    expect(checkboxShape(textOf(written, F.vml.partName))).toBe(checkboxShape(textOf(parts, F.vml.partName)));
    const comments = textOf(written, F.comments.partName);
    expect(comments).toContain(`<comment ref="B3" authorId="0" shapeId="0" xr:uid="${F.notes.B2.uid}">`);
    expect(comments).toContain(`<comment ref="D3" authorId="1" shapeId="0" xr:uid="${F.threads.D2.threadId}">`);
    expect(textOf(written, F.threadedComments.partName)).not.toContain(`ref="D2"`);
    expect(textOf(written, F.sheets.Summary.partName)).toContain(`<hyperlink ref="A7:B7" r:id="rId1"`);
    const reopened = await open(bytes);
    expect(summaryOf(reopened).notes!.map((note) => [note.noteId, note.box])).toEqual([
      ["n-B3", undefined],
      ["n-B5", F.notes.B4.box],
    ]);
    expect(summaryOf(reopened).threads!.map((thread) => thread.ref)).toEqual([
      { row: 2, col: 3 },
      { row: 2, col: 4 },
    ]);
  });

  it("sanctioned: a message carrying <mentions> re-emits it while unedited and loses it once edited", async () => {
    const mentions = `<mentions><mention mentionpersonId="${F.persons[1]!.personId}" mentionId="{AAAAAAAA-0000-4000-8000-000000000001}" startIndex="0" length="5"/></mentions>`;
    const bytes = await withPart(F.threadedComments.partName, (xml) => xml.replace("<text>From the Q3 ledger.</text>", `<text>From the Q3 ledger.</text>${mentions}`));
    const model = await open(bytes);
    const thread = summaryOf(model).threads![0]!;
    expect(thread.messages[1]!.mentions).toBe(mentions);
    replyToThread(model, summaryOf(model).id, thread.threadId, "Another", "You");
    expect(textOf(await unzip(await writeXlsx(model)), F.threadedComments.partName)).toContain(`<text>From the Q3 ledger.</text>${mentions}</threadedComment>`);
    editThreadMessage(model, summaryOf(model).id, thread.threadId, thread.messages[1]!.messageId, "From the Q3 ledger, page 4.");
    const written = textOf(await unzip(await writeXlsx(model)), F.threadedComments.partName);
    expect(written).not.toContain("<mentions>");
    expect(written).toContain("<text>From the Q3 ledger, page 4.</text></threadedComment>");
  });
});

// --- sheets coming and going ---------------------------------------------------------------------

describe("duplicated, deleted and crossed sheets", () => {
  it("a duplicated sheet gets its own comment, VML and threaded parts with fresh thread ids; both sheets reopen with their threads", async () => {
    const model = await open(await buildSpec121LinksXlsxBytes(), "spec121-links.xlsx");
    const data = model.sheets.findIndex((sheet) => sheet.name === "Data");
    expect(model.sheets[data]!.threads).toHaveLength(1);
    duplicateSheet(model, data);
    const copy = model.sheets[data + 1]!;
    expect(copy.threads![0]!.threadId).not.toBe(model.sheets[data]!.threads![0]!.threadId);
    const bytes = await writeXlsx(model);
    expect(await packageIntegrity(bytes)).toEqual([]);
    const reopened = await open(bytes, "copy.xlsx");
    const ids = reopened.sheets.flatMap((sheet) => (sheet.threads ?? []).map((thread) => thread.threadId));
    expect(ids).toHaveLength(2);
    expect(new Set(ids).size).toBe(2);
    expect(fields(reopened.sheets[data + 1]!)).toEqual(fields(copy));
    const written = await unzip(bytes);
    const threadedParts = [...written.keys()].filter((name) => name.startsWith("xl/threadedComments/"));
    expect(threadedParts).toHaveLength(2);
  });

  it("a deleted sheet's comment parts go with it; the workbook's persons part stays", async () => {
    const { model } = await fixture();
    deleteSheet(model, 0);
    const bytes = await writeXlsx(model);
    const written = await unzip(bytes);
    for (const name of [F.comments.partName, F.vml.partName, F.threadedComments.partName, SUMMARY_RELS]) {
      expect(written.has(name), name).toBe(false);
    }
    expect(written.has(F.person.partName)).toBe(true);
    expect(await packageIntegrity(bytes)).toEqual([]);
  });

  it("a copy as .xlsx from scratch (stripRetention) writes the notes, threads, persons and links as valid parts", async () => {
    const { model } = await fixture();
    const copy = stripRetention({ ...model, format: "xlsx" });
    const bytes = await writeXlsx(copy);
    expect(await packageIntegrity(bytes)).toEqual([]);
    const reopened = await open(bytes, "copy.xlsx");
    const crossed = summaryOf(copy);
    expect(fields(summaryOf(reopened))).toEqual(fields(crossed));
    expect(textOf(await unzip(bytes), "[Content_Types].xml")).toContain(`<Default Extension="vml"`);
  });

  it("a model crossed from .ods (notes without source, no persons, url and place links) writes valid parts that read back field for field", async () => {
    const ods = await readSpreadsheetBytes(await buildSpec135AnnotationsOdsBytes(), "ods", "spec135-annotations.ods");
    const copy = stripRetention({ ...ods, format: "xlsx" });
    const bytes = await writeXlsx(copy);
    expect(await packageIntegrity(bytes)).toEqual([]);
    const reopened = await open(bytes, "crossed.xlsx");
    copy.sheets.forEach((sheet, index) => {
      const strip = (each: Sheet): unknown => {
        const all = fields(each) as { notes: { createdAt?: string }[] };
        return { ...all, notes: all.notes.map(({ createdAt: _date, ...note }) => note) };
      };
      expect(strip(reopened.sheets[index]!), sheet.name).toEqual(strip(sheet));
    });
  });
});

// --- goldens -----------------------------------------------------------------------------------

/** Excel's default note shape for B2 (shape 1 of idmap 1), as Excel 365
 * writes it — the golden vmlNoteGolden.ts reproduces. */
const EXCEL_NOTE_SHAPE_B2 = [
  `<v:shape id="_x0000_s1025" type="#_x0000_t202" style='position:absolute;`,
  `  margin-left:107.25pt;margin-top:7.5pt;width:108pt;height:59.25pt;z-index:1;`,
  `  visibility:hidden' fillcolor="#ffffe1" o:insetmode="auto">`,
  `  <v:fill color2="#ffffe1"/>`,
  `  <v:shadow color="black" obscured="t"/>`,
  `  <v:path o:connecttype="none"/>`,
  `  <v:textbox style='mso-direction-alt:auto'>`,
  `   <div style='text-align:left'></div>`,
  `  </v:textbox>`,
  `  <x:ClientData ObjectType="Note">`,
  `   <x:MoveWithCells/>`,
  `   <x:SizeWithCells/>`,
  `   <x:Anchor>`,
  `    2, 15, 0, 10, 4, 15, 4, 4</x:Anchor>`,
  `   <x:AutoFill>False</x:AutoFill>`,
  `   <x:Row>1</x:Row>`,
  `   <x:Column>1</x:Column>`,
  `  </x:ClientData>`,
  ` </v:shape>`,
].join("\r\n");

describe("goldens", () => {
  it("vmlNoteGolden: Excel's default note shape byte for byte — the golden text, and the fixture's B2, D2 and E2 shapes", async () => {
    expect(vmlNoteShapeFor({ row: 1, col: 1 }, { idmap: 1, n: 1, visible: false })).toBe(EXCEL_NOTE_SHAPE_B2);
    expect(defaultNoteAnchor({ row: 0, col: 0 })).toEqual([1, 15, 0, 10, 3, 15, 4, 4]);
    expect(defaultNoteAnchor({ row: 9, col: 2 })).toEqual([3, 15, 8, 10, 5, 15, 12, 4]);
    const shapes = vmlShapes(textOf((await fixture()).parts, F.vml.partName));
    expect(shapes[0]).toBe(EXCEL_NOTE_SHAPE_B2);
    expect(shapes[2]).toBe(vmlNoteShapeFor({ row: 1, col: 3 }, { idmap: 1, n: 3, visible: false }));
    expect(shapes[3]).toBe(vmlNoteShapeFor({ row: 1, col: 4 }, { idmap: 1, n: 4, visible: false }));
    const shown = vmlNoteShapeFor({ row: 3, col: 1 }, { idmap: 1, n: 2, visible: true, box: { widthPx: 200, heightPx: 100 } });
    expect(shown).toContain("width:150pt;height:75pt;z-index:2;\r\n  visibility:visible'");
    expect(shown).toContain("   <x:Column>1</x:Column>\r\n   <x:Visible/>\r\n  </x:ClientData>");
  });

  it("threadPlaceholder: the preamble, the tc= author and the comment element reproduce the fixture's placeholders byte for byte", async () => {
    expect(THREAD_PLACEHOLDER_PREAMBLE).toBe(F.placeholder.preamble);
    const { model, parts } = await fixture();
    const comments = textOf(parts, F.comments.partName);
    for (const [index, thread] of summaryOf(model).threads!.entries()) {
      expect(placeholderAuthor(thread.threadId)).toBe(F.authors[index + 1]);
      expect(placeholderText(thread)).toBe(
        `${F.placeholder.preamble}Comment:\n    ${thread.messages[0]!.text}${thread.messages
          .slice(1)
          .map((reply) => `\nReply:\n    ${reply.text}`)
          .join("")}`,
      );
      expect(comments).toContain(placeholderCommentXml(thread, index === 0 ? "D2" : "E2", index + 1));
    }
  });
});

// --- the human-captured Excel 365 workbook ------------------------------------------------------

/** A comment element's `<text>` element, as written. */
function textElement(element: string): string {
  return /<(?:\w+:)?text>[\s\S]*<\/(?:\w+:)?text>/.exec(element)![0];
}

/** Every legacy comment of a package: its author and its `<text>` element as written. */
function legacyComments(parts: Map<string, Uint8Array>): { part: string; ref: string; author: string; text: string; element: string }[] {
  return [...parts.keys()]
    .filter((name) => /^xl\/comments[^/]*\.xml$/.test(name))
    .flatMap((part) => {
      const xml = textOf(parts, part);
      const doc = parseXml(xml);
      const authors = Array.from(doc.getElementsByTagNameNS("*", "author")).map((author) => author.textContent ?? "");
      const elements = [...xml.matchAll(/<(?:\w+:)?comment\b[\s\S]*?<\/(?:\w+:)?comment>/g)].map((match) => match[0]);
      return Array.from(doc.getElementsByTagNameNS("*", "comment")).map((comment, index) => ({
        part,
        ref: comment.getAttribute("ref") ?? "",
        author: authors[Number(comment.getAttribute("authorId"))] ?? "",
        text: comment.getElementsByTagNameNS("*", "text")[0]?.textContent ?? "",
        element: elements[index] ?? "",
      }));
    });
}

describe(`the Excel 365 workbook (${EXCEL365})`, () => {
  const load = async (): Promise<{ bytes: Uint8Array; parts: Map<string, Uint8Array>; model: SpreadsheetModel }> => {
    const bytes = new Uint8Array(readFileSync(EXCEL365_PATH));
    return { bytes, parts: await unzip(bytes), model: await open(bytes, "spec135-excel365-threads.xlsx") };
  };

  it("reads every thread with its messages, its authors' display names through model.persons and its resolved state — without the fallback", async (ctx) => {
    ctx.skip(!existsSync(EXCEL365_PATH), EXCEL365_SKIP);
    const { parts, model } = await load();
    const threads = model.sheets.flatMap((sheet) => sheet.threads ?? []);
    const placeholders = legacyComments(parts).filter((comment) => comment.author.startsWith("tc="));
    expect(threads.length).toBeGreaterThan(0);
    expect(threads).toHaveLength(placeholders.length);
    const persons = new Map((model.persons ?? []).map((person) => [person.personId, person.displayName] as const));
    for (const thread of threads) {
      expect(thread.messages.length).toBeGreaterThan(0);
      for (const message of thread.messages) {
        expect(persons.get(message.personId)).toBe(message.author);
      }
    }
    expect(threads.some((thread) => thread.resolved)).toBe(true);
    expect(threads.some((thread) => thread.messages.length > 1)).toBe(true);
  });

  it("models its legacy note as a CellNote and no tc= placeholder as a note", async (ctx) => {
    ctx.skip(!existsSync(EXCEL365_PATH), EXCEL365_SKIP);
    const { parts, model } = await load();
    const notes = model.sheets.flatMap((sheet) => sheet.notes ?? []);
    const legacy = legacyComments(parts).filter((comment) => !comment.author.startsWith("tc="));
    expect(legacy.length).toBeGreaterThan(0);
    expect(notes).toHaveLength(legacy.length);
    expect(notes.some((note) => note.author.startsWith("tc="))).toBe(false);
  });

  it("regenerates each placeholder's author and <text> byte for byte", async (ctx) => {
    ctx.skip(!existsSync(EXCEL365_PATH), EXCEL365_SKIP);
    const { parts, model } = await load();
    const threads = new Map(model.sheets.flatMap((sheet) => (sheet.threads ?? []).map((thread) => [thread.threadId, thread] as const)));
    for (const comment of legacyComments(parts).filter((each) => each.author.startsWith("tc="))) {
      const thread = threads.get(comment.author.slice(3));
      expect(thread, comment.author).toBeDefined();
      expect(placeholderAuthor(thread!.threadId)).toBe(comment.author);
      expect(placeholderText(thread!)).toBe(comment.text);
      const prefix = /^<(\w+:)?comment/.exec(comment.element)![1] ?? "";
      expect(textElement(placeholderCommentXml(thread!, comment.ref, 0, prefix))).toBe(textElement(comment.element));
    }
  });

  it("downloads byte-identical part for part with no edit", async (ctx) => {
    ctx.skip(!existsSync(EXCEL365_PATH), EXCEL365_SKIP);
    const { bytes, model } = await load();
    expect(changedParts(await unzip(bytes), await unzip(await writeXlsx(model)))).toEqual([]);
  });
});
