// SPEC-135 "ods wire mapping" (ods/annotations.ts, the reader's and the
// table writer's annotation slots): spec135-annotations.ods read field by
// field — notes lifted out of preservedXml, each cell's first text:a a link,
// nothing dirty; the href table; repeated runs, office:name and the note
// paragraph text; writing — a note its cell's first child (verbatim, moved,
// generated), a repeated empty run split for a note, text:a wrapping with a
// title, internal, external and name targets, preserved text kept only
// while its link matches; the two-links delta; a no-edit download
// byte-identical part for part; add/edit/link → save → reopen field-level
// equality (ids and sources excluded) and byte stability; an .xlsx model's
// threads, notes and links downloaded as .ods; and the reader-upgrade
// helpers the coordinator's ReaderUpgrade calls.
import { describe, expect, it } from "vitest";
import { replyParagraphText } from "../../core/comments";
import { equalBytes } from "../../core/equality";
import { unzip } from "../../core/zip";
import { equalStructural } from "../../equality";
import { formatA1, getCell, parseA1 } from "../address";
import { addNote, editNote, removeHyperlink, setHyperlink, setNoteVisible } from "../annotations";
import { setCellEntry } from "../cells";
import { spreadsheetDownload } from "../download";
import { buildSpec135AnnotationsOdsBytes, SPEC135_ANNOTATIONS_ODS } from "../fixtures/build";
import type { CellHyperlink, CellNote, CellThread, Sheet, SpreadsheetModel } from "../model";
import { readSpreadsheetBytes } from "../open";
import { createBlankWorkbook, renameSheet } from "../sheets";
import { applyStructuralEdit } from "../structure";
import {
  liftOdsSheetNotes,
  linkedParagraphsXml,
  odsHyperlinkTarget,
  odsLinkHref,
  odsNoteXml,
  paragraphsMatchLink,
  readOdsAnnotation,
  retainedOdsAnnotations,
  upgradeOdsSheetAnnotations,
  upgradeOdsSheetLinks,
} from "./annotations";
import { writeOds } from "./writer";

const O = SPEC135_ANNOTATIONS_ODS;
const decoder = new TextDecoder();
const encoder = new TextEncoder();

const open = async (): Promise<SpreadsheetModel> => readSpreadsheetBytes(await buildSpec135AnnotationsOdsBytes(), "ods", O.title);
const reopen = (bytes: Uint8Array): Promise<SpreadsheetModel> => readSpreadsheetBytes(bytes, "ods", O.title);
const sheetNamed = (model: SpreadsheetModel, name: string): Sheet => model.sheets.find((sheet) => sheet.name === name)!;
const at = (a1: string): { row: number; col: number } => {
  const { row, col } = parseA1(a1)!;
  return { row, col };
};

/** Marks a table rewritten — as any edit on it leaves it (sheet.dirty). */
function rewrite(model: SpreadsheetModel, name: string): void {
  model.sheets = model.sheets.map((sheet) => (sheet.name === name ? { ...sheet, dirty: true } : sheet));
}

async function contentXml(bytes: Uint8Array): Promise<string> {
  return decoder.decode((await unzip(bytes)).get("content.xml")!);
}

/** The fixture with content.xml edited (each replacement must apply). */
async function patched(edits: [string, string][]): Promise<SpreadsheetModel> {
  const parts = await unzip(await buildSpec135AnnotationsOdsBytes());
  let text = decoder.decode(parts.get("content.xml"));
  for (const [from, to] of edits) {
    expect(text.includes(from), from).toBe(true);
    text = text.replace(from, to);
  }
  parts.set("content.xml", encoder.encode(text));
  return readSpreadsheetBytes(new Uint8Array(0), "ods", O.title, parts);
}

/** A table's text as written (content.xml sliced at the table). */
function tableXml(content: string, name: string): string {
  const start = content.indexOf(`<table:table table:name="${name}"`);
  return content.slice(start, content.indexOf("</table:table>", start));
}

/** The cell element written at `a1` of a table's text, walking repeats. */
function cellXmlAt(table: string, a1: string): string | undefined {
  const target = at(a1);
  let row = 0;
  for (const rowMatch of table.matchAll(/<table:table-row\b([^>]*)>([\s\S]*?)<\/table:table-row>/g)) {
    const rows = Number(/table:number-rows-repeated="(\d+)"/.exec(rowMatch[1]!)?.[1] ?? 1);
    if (target.row >= row && target.row < row + rows) {
      let col = 0;
      for (const cell of rowMatch[2]!.matchAll(/<table:(?:covered-)?table-cell\b[^>]*?(?:\/>|>[\s\S]*?<\/table:(?:covered-)?table-cell>)/g)) {
        const cols = Number(/^<[^>]*table:number-columns-repeated="(\d+)"/.exec(cell[0])?.[1] ?? 1);
        if (target.col >= col && target.col < col + cols) {
          return cell[0];
        }
        col += cols;
      }
      return undefined;
    }
    row += rows;
  }
  return undefined;
}

/** Note text as plain paragraphs, one per line. */
const paragraphs = (text: string): string => text.split("\n").map((line) => `<text:p>${line}</text:p>`).join("");

/** A thread message. */
const message = (messageId: string, author: string, text: string, createdAt: string): CellThread["messages"][number] => ({ messageId, personId: `{P-${author}}`, author, text, createdAt });

/** A note or link without its id and source — the fields a round trip reproduces. */
const noteFields = ({ noteId: _id, source: _source, ...note }: CellNote): unknown => note;
const linkFields = ({ linkId: _id, source: _source, ...link }: CellHyperlink): unknown => link;

describe("reading spec135-annotations.ods", () => {
  it("each annotation is a note — creator, date, paragraphs, display, size, name — lifted out of the cell's preservedXml", async () => {
    const model = await open();
    const summary = sheetNamed(model, "Summary");
    const notes = summary.notes ?? [];
    expect(notes.map((note) => formatA1(note.ref))).toEqual(["B2", "B4"]);
    for (const [index, fact] of [O.notes.B2, O.notes.B4].entries()) {
      const note = notes[index]!;
      const expected: CellNote = {
        noteId: `n-${fact.a1}`,
        ref: at(fact.a1),
        author: fact.author,
        createdAt: fact.createdAt,
        text: fact.text,
        box: fact.box,
        source: { odsAnnotation: fact.xml, odsRef: at(fact.a1) },
      };
      if (fact.visible) {
        expected.visible = true;
      }
      expect(note).toEqual(expected);
      const cell = getCell(summary, at(fact.a1).row, at(fact.a1).col)!;
      expect(cell.value).toBe(O.values.Summary[fact.a1]);
      expect(cell.preservedXml).toBeUndefined();
    }
    expect(sheetNamed(model, "Data").notes).toBeUndefined();
    expect(model.sheets.some((sheet) => sheet.threads !== undefined)).toBe(false);
  });

  it("each cell's first text:a is a link over that cell — external with its title, internal as canonical place text — and no preserved text", async () => {
    const model = await open();
    const summary = sheetNamed(model, "Summary");
    expect(summary.hyperlinks).toEqual(
      [O.links.A7, O.links.B7, O.links.C9].map((fact) => {
        const ref = at(fact.a1);
        const link: CellHyperlink = { linkId: `h-${fact.a1}`, range: { top: ref.row, left: ref.col, bottom: ref.row, right: ref.col }, target: fact.target };
        if ("tooltip" in fact) {
          link.tooltip = fact.tooltip;
        }
        return link;
      }),
    );
    for (const fact of [O.links.A7, O.links.B7, O.links.C9]) {
      const cell = getCell(summary, at(fact.a1).row, at(fact.a1).col)!;
      expect(cell).toEqual({ col: at(fact.a1).col, value: O.values.Summary[fact.a1] });
    }
    expect(sheetNamed(model, "Data").hyperlinks).toBeUndefined();
  });

  it("reading sets nothing dirty", async () => {
    const model = await open();
    expect(model.sheets.map((sheet) => [sheet.dirty, sheet.annotationsDirty])).toEqual([
      [undefined, undefined],
      [undefined, undefined],
    ]);
  });

  it("maps hrefs: #Sheet.A1, #'My Sheet'.A1:B2, #$Sheet.$A$1 and a name are places; anything else a url, a dangerous scheme kept", () => {
    expect(odsHyperlinkTarget("#Sheet2.A1")).toEqual({ kind: "place", location: "Sheet2!A1" });
    expect(odsHyperlinkTarget("#'My Sheet'.A1:B2")).toEqual({ kind: "place", location: "'My Sheet'!A1:B2" });
    expect(odsHyperlinkTarget("#$Sheet2.$A$1")).toEqual({ kind: "place", location: "Sheet2!$A$1" });
    expect(odsHyperlinkTarget("#My Sheet.B3")).toEqual({ kind: "place", location: "'My Sheet'!B3" });
    expect(odsHyperlinkTarget("#MyName")).toEqual({ kind: "place", location: "MyName" });
    expect(odsHyperlinkTarget("https://example.com/a#b")).toEqual({ kind: "url", href: "https://example.com/a#b" });
    expect(odsHyperlinkTarget("mailto:team@example.com")).toEqual({ kind: "url", href: "mailto:team@example.com" });
    expect(odsHyperlinkTarget("javascript:alert(1)")).toEqual({ kind: "url", href: "javascript:alert(1)" });
    expect(odsHyperlinkTarget("../other.ods")).toEqual({ kind: "url", href: "../other.ods" });
  });

  it("a dangerous link read from the file stays in the model and in the rewritten file", async () => {
    const model = await patched([[`xlink:href="#Data.A1"`, `xlink:href="javascript:alert(1)"`]]);
    const summary = sheetNamed(model, "Summary");
    expect(summary.hyperlinks![1]!.target).toEqual({ kind: "url", href: "javascript:alert(1)" });
    rewrite(model, "Summary");
    const content = await contentXml(await writeOds(model));
    expect(cellXmlAt(tableXml(content, "Summary"), "B7")).toContain(`<text:a xlink:type="simple" xlink:href="javascript:alert(1)">See data</text:a>`);
  });

  it("a repeated cell gives every repetition the note (office:name naming the first); note paragraphs expand text:s, text:tab and text:line-break", async () => {
    const annotation = `<office:annotation office:name="Pair"><dc:creator>Q</dc:creator><text:p>a<text:s text:c="2"/>b<text:tab/>c<text:line-break/>d</text:p><text:p/></office:annotation>`;
    const model = await patched([
      [
        `<table:table-row table:style-name="ro1"><table:table-cell table:number-columns-repeated="16384"/></table:table-row>`,
        `<table:table-row table:style-name="ro1"><table:table-cell table:number-columns-repeated="2" office:value-type="string" calcext:value-type="string">${annotation}<text:p>x</text:p></table:table-cell><table:table-cell table:number-columns-repeated="16382"/></table:table-row>`,
      ],
    ]);
    const notes = sheetNamed(model, "Summary").notes!;
    expect(notes.map((note) => [note.noteId, formatA1(note.ref), note.author, note.text, note.createdAt, note.box])).toEqual([
      ["n-B2", "B2", O.notes.B2.author, O.notes.B2.text, O.notes.B2.createdAt, O.notes.B2.box],
      ["Pair", "A3", "Q", "a  b\tc\nd\n", undefined, undefined],
      ["n-B3", "B3", "Q", "a  b\tc\nd\n", undefined, undefined],
      ["n-B4", "B4", O.notes.B4.author, O.notes.B4.text, O.notes.B4.createdAt, O.notes.B4.box],
    ]);
    expect(getCell(sheetNamed(model, "Summary"), 2, 1)).toEqual({ col: 1, value: "x" });
  });

  it("a cell holding only a note is not stored; a link with an unmodelled attribute keeps its paragraph as preserved text, its link modelled", async () => {
    const model = await patched([
      [
        `<table:table-row table:style-name="ro1"><table:table-cell table:number-columns-repeated="16384"/></table:table-row>`,
        `<table:table-row table:style-name="ro1"><table:table-cell table:number-columns-repeated="3"/><table:table-cell><office:annotation><text:p>Lonely</text:p></office:annotation></table:table-cell><table:table-cell table:number-columns-repeated="16380"/></table:table-row>`,
      ],
      [`xlink:href="https://example.com/" xlink:type="simple"`, `xlink:href="https://example.com/" xlink:type="simple" office:target-frame-name="_blank"`],
    ]);
    const summary = sheetNamed(model, "Summary");
    expect(getCell(summary, 2, 3)).toBeUndefined();
    expect(summary.notes!.map((note) => [formatA1(note.ref), note.author, note.text])).toContainEqual(["D3", "", "Lonely"]);
    expect(getCell(summary, 6, 0)!.preservedText).toContain(`office:target-frame-name="_blank"`);
    expect(summary.hyperlinks![0]).toMatchObject({ range: { top: 6, left: 0, bottom: 6, right: 0 }, target: { kind: "url", href: "https://example.com/" }, tooltip: O.links.A7.title });
  });

  it("readOdsAnnotation reads svg sizes in any unit and fails soft", () => {
    expect(readOdsAnnotation(`<office:annotation svg:width="5.08cm" svg:height="25.4mm"/>`, "")).toEqual({ author: "", text: "", visible: false, box: { widthPx: 192, heightPx: 96 } });
    expect(readOdsAnnotation(`<office:annotation office:display="false"><dc:creator> A </dc:creator><text:list><text:list-item><text:p>i</text:p></text:list-item></text:list></office:annotation>`, "")).toEqual({
      author: " A ",
      text: "i",
      visible: false,
    });
    expect(readOdsAnnotation(`<office:annotation>`, "")).toBeNull();
    expect(readOdsAnnotation(`<text:p/>`, "")).toBeNull();
  });
});

describe("writing notes and links into .ods", () => {
  it("a no-edit download is byte-identical part for part", async () => {
    const bytes = await buildSpec135AnnotationsOdsBytes();
    const model = await reopen(bytes);
    const { bytes: written } = await spreadsheetDownload(model, "ods");
    const left = await unzip(bytes);
    const right = await unzip(written);
    expect([...right.keys()]).toEqual([...left.keys()]);
    for (const [name, data] of left) {
      expect(equalBytes(right.get(name)!, data), name).toBe(true);
    }
  });

  it("a rewritten table writes each unchanged note verbatim as its cell's first child, and each link around its text", async () => {
    const model = await open();
    rewrite(model, "Summary");
    const table = tableXml(await contentXml(await writeOds(model)), "Summary");
    for (const note of [O.notes.B2, O.notes.B4]) {
      expect(cellXmlAt(table, note.a1)).toBe(`<table:table-cell office:value-type="float" office:value="${O.values.Summary[note.a1]}" calcext:value-type="float">${note.xml}<text:p>${O.values.Summary[note.a1]}</text:p></table:table-cell>`);
    }
    expect(cellXmlAt(table, "A7")).toBe(
      `<table:table-cell office:value-type="string" calcext:value-type="string"><text:p><text:a xlink:type="simple" xlink:href="https://example.com/" office:title="Open the example site">Example site</text:a></text:p></table:table-cell>`,
    );
    expect(cellXmlAt(table, "B7")).toContain(`<text:p><text:a xlink:type="simple" xlink:href="#$Data.A1">See data</text:a></text:p>`);
  });

  it("the sanctioned delta: a cell that carried two links keeps the first once its table is rewritten", async () => {
    const model = await open();
    rewrite(model, "Summary");
    const written = await writeOds(model);
    expect(cellXmlAt(tableXml(await contentXml(written), "Summary"), "C9")).toBe(
      `<table:table-cell office:value-type="string" calcext:value-type="string"><text:p><text:a xlink:type="simple" xlink:href="https://one.example/">One and Two</text:a></text:p></table:table-cell>`,
    );
    const again = sheetNamed(await reopen(written), "Summary");
    expect(getCell(again, 8, 2)!.value).toBe(O.values.Summary.C9);
    expect(again.hyperlinks!.filter((link) => link.range.top === 8).map((link) => link.target)).toEqual([O.links.C9.target]);
  });

  it("a renamed destination sheet: the place location follows (the sheet holding it rewritten) and is written from the model", async () => {
    const model = await open();
    renameSheet(model, 1, "Facts");
    const summary = sheetNamed(model, "Summary");
    expect(summary.hyperlinks![1]!.target).toEqual({ kind: "place", location: "Facts!A1" });
    expect(summary.dirty).toBe(true);
    const cell = cellXmlAt(tableXml(await contentXml(await writeOds(model)), "Summary"), "B7")!;
    expect(cell).toContain(`<text:p><text:a xlink:type="simple" xlink:href="#$Facts.A1">See data</text:a></text:p>`);
  });

  it("a new note on an unstored cell splits the repeated empty run and is generated per the wire mapping", async () => {
    const model = await open();
    const summary = sheetNamed(model, "Summary");
    addNote(model, summary.id, at("D4"), "Check  this\nline two & <more>", "You");
    const table = tableXml(await contentXml(await writeOds(model)), "Summary");
    const row4 = /<table:table-row table:style-name="ro1">(?:(?!<\/table:table-row>)[\s\S])*Costs[\s\S]*?<\/table:table-row>/.exec(table)![0];
    expect(row4).toContain(
      `<table:table-cell/><table:table-cell><office:annotation xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:creator>You</dc:creator>` +
        `<text:p>Check<text:s text:c="2"/>this</text:p><text:p>line two &amp; &lt;more&gt;</text:p></office:annotation></table:table-cell><table:table-cell table:number-columns-repeated="16380"/>`,
    );
  });

  it("a generated note carries office:display, a wire name, its size in cm and its date", () => {
    const note: CellNote = { noteId: "Annotation 1", ref: { row: 0, col: 0 }, author: "Ann", createdAt: "2026-03-01T09:00:00", text: "a\nb", visible: true, box: { widthPx: 153.6, heightPx: 48 } };
    expect(odsNoteXml(note, paragraphs)).toBe(
      `<office:annotation xmlns:dc="http://purl.org/dc/elements/1.1/" office:display="true" office:name="Annotation 1" svg:width="4.064cm" svg:height="1.27cm">` +
        `<dc:creator>Ann</dc:creator><dc:date>2026-03-01T09:00:00</dc:date><text:p>a</text:p><text:p>b</text:p></office:annotation>`,
    );
    expect(odsNoteXml({ noteId: "n-B2-2", ref: { row: 1, col: 1 }, author: "", text: "" }, paragraphs)).toBe(`<office:annotation xmlns:dc="http://purl.org/dc/elements/1.1/"><text:p></text:p></office:annotation>`);
  });

  it("a moved note keeps its source and size but not its caption position (the ODS face of the sanctioned delta)", async () => {
    const model = await open();
    const summary = sheetNamed(model, "Summary");
    expect(applyStructuralEdit(model, { kind: "insertRows", sheetId: summary.id, at: 1, count: 1 })).toEqual({ ok: true });
    const moved = sheetNamed(model, "Summary").notes!.find((note) => formatA1(note.ref) === "B3")!;
    expect(moved.source?.odsAnnotation).toBe(O.notes.B2.xml);
    const written = await writeOds(model);
    const cell = cellXmlAt(tableXml(await contentXml(written), "Summary"), "B3")!;
    expect(cell).toContain(
      `<office:annotation draw:style-name="gr1" draw:text-style-name="P1" svg:width="1.6in" svg:height="0.5in"><dc:creator>Alex Kim</dc:creator>`,
    );
    const again = sheetNamed(await reopen(written), "Summary").notes!;
    expect(again.map((note) => [formatA1(note.ref), note.text, note.box])).toEqual([
      ["B3", O.notes.B2.text, O.notes.B2.box],
      ["B5", O.notes.B4.text, O.notes.B4.box],
    ]);
  });

  it("a shown/hidden toggle regenerates the element with office:display", async () => {
    const model = await open();
    const summary = sheetNamed(model, "Summary");
    setNoteVisible(model, summary.id, "n-B2", true);
    const cell = cellXmlAt(tableXml(await contentXml(await writeOds(model)), "Summary"), "B2")!;
    expect(cell).toContain(`<office:annotation xmlns:dc="http://purl.org/dc/elements/1.1/" office:display="true" svg:width="4.064cm" svg:height="1.27cm"><dc:creator>Alex Kim</dc:creator><dc:date>2026-03-01T09:00:00</dc:date>`);
  });

  it("preserved text is kept while its first text:a matches the model's link, regenerated once the link is removed", async () => {
    const model = await patched([[`xlink:href="https://example.com/" xlink:type="simple"`, `xlink:href="https://example.com/" xlink:type="simple" office:target-frame-name="_blank"`]]);
    rewrite(model, "Summary");
    const verbatim = cellXmlAt(tableXml(await contentXml(await writeOds(model)), "Summary"), "A7")!;
    expect(verbatim).toContain(`office:target-frame-name="_blank"`);
    removeHyperlink(model, sheetNamed(model, "Summary").id, "h-A7");
    expect(cellXmlAt(tableXml(await contentXml(await writeOds(model)), "Summary"), "A7")).toBe(
      `<table:table-cell office:value-type="string" calcext:value-type="string"><text:p>Example site</text:p></table:table-cell>`,
    );
  });

  it("place targets are written as # + an ODS address, names as # + the name, urls with their subaddress; linked paragraphs wrap each one", () => {
    expect(odsLinkHref({ kind: "place", location: "Data!A10" })).toBe("#$Data.A10");
    expect(odsLinkHref({ kind: "place", location: "'My Sheet'!A1:B2" })).toBe("#$'My Sheet'.A1:.B2");
    expect(odsLinkHref({ kind: "place", location: "MyName" })).toBe("#MyName");
    expect(odsLinkHref({ kind: "url", href: "https://example.com/doc", subaddress: "part" })).toBe("https://example.com/doc#part");
    expect(linkedParagraphsXml(`<text:p>a</text:p><text:p/><text:p>b</text:p>`, { target: { kind: "url", href: "https://x.example/?a=1&b=2" }, tooltip: `Say "hi"` })).toBe(
      `<text:p><text:a xlink:type="simple" xlink:href="https://x.example/?a=1&amp;b=2" office:title="Say &quot;hi&quot;">a</text:a></text:p><text:p/>` +
        `<text:p><text:a xlink:type="simple" xlink:href="https://x.example/?a=1&amp;b=2" office:title="Say &quot;hi&quot;">b</text:a></text:p>`,
    );
    expect(paragraphsMatchLink(`<text:p>plain</text:p>`, undefined)).toBe(true);
    expect(paragraphsMatchLink(`<text:p><text:a xlink:href="#Data.A1">x</text:a></text:p>`, { target: { kind: "place", location: "Data!A1" } })).toBe(true);
    expect(paragraphsMatchLink(`<text:p><text:a xlink:href="#Data.A1">x</text:a></text:p>`, { target: { kind: "place", location: "Facts!A1" } })).toBe(false);
    expect(paragraphsMatchLink(`<text:p><text:a xlink:href="#Data.A1" office:title="t">x</text:a></text:p>`, { target: { kind: "place", location: "Data!A1" } })).toBe(false);
  });
});

describe("round trip through .ods", () => {
  it("add, edit and link → save → reopen reproduces every modelled field; writing the reopened model is byte-stable", async () => {
    const model = await open();
    const summary = sheetNamed(model, "Summary");
    addNote(model, summary.id, at("D4"), "Added\nTwo lines", "You");
    editNote(model, summary.id, "n-B2", "Edited note");
    setHyperlink(model, summary.id, { top: 3, left: 0, bottom: 3, right: 0 }, { kind: "place", location: "Data!A10" }, "Go to the target");
    setHyperlink(model, summary.id, { top: 0, left: 1, bottom: 0, right: 1 }, { kind: "place", location: "MyName" });
    setHyperlink(model, summary.id, { top: 0, left: 0, bottom: 0, right: 0 }, { kind: "url", href: "mailto:team@example.com" });
    const before = sheetNamed(model, "Summary");
    const written = await writeOds(model);
    const reopened = await reopen(written);
    const after = sheetNamed(reopened, "Summary");
    expect(equalStructural((after.notes ?? []).map(noteFields), (before.notes ?? []).map(noteFields))).toBe(true);
    expect((after.notes ?? []).map(noteFields)).toEqual((before.notes ?? []).map(noteFields));
    // C9's second link is the sanctioned delta; every link the model holds comes back.
    expect(equalStructural((after.hyperlinks ?? []).map(linkFields), (before.hyperlinks ?? []).map(linkFields))).toBe(true);
    expect(after.dirty).toBeUndefined();
    // Byte-stable: the reopened model written as it is, and rebuilt from the model.
    expect(equalBytes(await writeOds(reopened), written)).toBe(true);
    const rebuilt = await reopen(written);
    rebuilt.sheets = rebuilt.sheets.map((sheet) => ({ ...sheet, dirty: true }));
    const again = await writeOds(rebuilt);
    expect(tableXml(await contentXml(again), "Summary")).toBe(tableXml(await contentXml(written), "Summary"));
  });
});

describe("an .xlsx model downloaded as .ods", () => {
  /** A workbook with a resolved thread (two replies), a note and links. */
  function xlsxModel(): SpreadsheetModel {
    const model = createBlankWorkbook();
    setCellEntry(model, 0, at("A1"), { value: "Check" });
    setCellEntry(model, 0, at("B1"), { value: 1200 });
    setCellEntry(model, 0, at("A2"), { value: "Report" });
    setCellEntry(model, 0, at("B2"), { value: "Q3" });
    const sheet = model.sheets[0]!;
    model.sheets = [
      {
        ...sheet,
        threads: [
          {
            threadId: "{T1}",
            ref: at("A1"),
            resolved: true,
            messages: [
              message("{T1}", "Sam Lee", "Where does this number come from?", "2026-03-01T09:00:00.00"),
              message("{T2}", "Alex Kim", "From the Q3 ledger.", "2026-03-01T10:00:00.00"),
              message("{T3}", "Sam Lee", "Thanks, that helps.", "2026-03-01T11:00:00.00"),
            ],
          },
        ],
        notes: [{ noteId: "n-B1", ref: at("B1"), author: "Alex Kim", text: "Alex Kim:\nQ3 figure, unaudited.", runs: [{ text: "Alex Kim:", font: { bold: true } }, { text: "\nQ3 figure, unaudited." }], source: { xlsxComment: "<comment/>" } }],
        hyperlinks: [
          { linkId: "{L1}", range: { top: 1, left: 0, bottom: 1, right: 1 }, target: { kind: "url", href: "https://example.com/report" }, tooltip: "The report", source: { relId: "rId1" } },
        ],
        annotationsDirty: true,
      },
    ];
    return model;
  }

  it("a thread becomes a note — the opening message, then one replyParagraphText paragraph per reply, resolved dropped — and notes and links convert field for field", async () => {
    const model = xlsxModel();
    const { bytes } = await spreadsheetDownload(model, "ods");
    const back = (await reopen(bytes)).sheets[0]!;
    const thread = model.sheets[0]!.threads![0]!;
    expect((back.notes ?? []).map(noteFields)).toEqual([
      { ref: at("A1"), author: "Sam Lee", createdAt: "2026-03-01T09:00:00.00", text: [thread.messages[0]!.text, ...thread.messages.slice(1).map(replyParagraphText)].join("\n") },
      { ref: at("B1"), author: "Alex Kim", text: "Alex Kim:\nQ3 figure, unaudited." },
    ]);
    expect(back.threads).toBeUndefined();
    // A range link reaches each of its cells (an .ods link is a cell's text:a).
    expect((back.hyperlinks ?? []).map(linkFields)).toEqual([
      { range: { top: 1, left: 0, bottom: 1, right: 0 }, target: { kind: "url", href: "https://example.com/report" }, tooltip: "The report" },
      { range: { top: 1, left: 1, bottom: 1, right: 1 }, target: { kind: "url", href: "https://example.com/report" }, tooltip: "The report" },
    ]);
    const content = await contentXml(bytes);
    expect(content).toContain(`<text:p>Alex Kim (2026-03-01T10:00:00.00): From the Q3 ledger.</text:p>`);
    expect(content).toMatch(/<office:document-content [^>]*xmlns:xlink="http:\/\/www\.w3\.org\/1999\/xlink"/);
  });
});

describe("the reader-upgrade helpers", () => {
  const dates = { dateSystem: "1900" as const };

  it("liftOdsSheetNotes moves each office:annotation out of preservedXml into a note at the cell's current position, copy-on-write", async () => {
    const model = await open();
    const prefix = model.odsContentEnvelope!.prefix;
    const xml = O.notes.B2.xml;
    const rows = [
      { index: 2, cells: [{ col: 1, value: 1200, preservedXml: `${xml}<text:list/>` }, { col: 2, value: null, preservedXml: xml }] },
      { index: 5, cells: [{ col: 0, value: "kept" }] },
    ];
    const lifted = liftOdsSheetNotes({ sheet: { rows }, prefix });
    expect(lifted.rows).toEqual([{ index: 2, cells: [{ col: 1, value: 1200, preservedXml: `<text:list/>` }] }, rows[1]]);
    expect(lifted.rows[1]).toBe(rows[1]);
    expect(rows[0]!.cells).toHaveLength(2);
    expect(lifted.notes.map((note) => [note.noteId, formatA1(note.ref), note.text, note.source?.odsAnnotation, note.source?.odsRef])).toEqual([
      ["n-B3", "B3", O.notes.B2.text, xml, at("B3")],
      ["n-C3", "C3", O.notes.B2.text, xml, at("C3")],
    ]);
    const untouched = [{ index: 0, cells: [{ col: 0, value: "x" }] }];
    expect(liftOdsSheetNotes({ sheet: { rows: untouched }, prefix }).rows).toBe(untouched);
  });

  it("retainedOdsAnnotations reads the retained table text's notes and links at source addresses, with each cell's source value", async () => {
    const model = await open();
    const summary = sheetNamed(model, "Summary");
    const retained = retainedOdsAnnotations(summary.sourceXml, model.odsContentEnvelope!.prefix, dates);
    expect(retained.annotations.map((each) => [formatA1(each.ref), each.xml, each.sourceValue])).toEqual([
      ["B2", O.notes.B2.xml, 1200],
      ["B4", O.notes.B4.xml, 800],
    ]);
    expect(retained.links.map((each) => [formatA1(each.ref), each.hyperlink, each.sourceValue])).toEqual(summary.hyperlinks!.map((link) => [link.linkId.slice(2), link, O.values.Summary[link.linkId.slice(2) as "A7"]]));
  });

  it("upgradeOdsSheetAnnotations: notes from preservedXml, else from the source where the value is unchanged; links only where the value is unchanged", async () => {
    const model = await open();
    const prefix = model.odsContentEnvelope!.prefix;
    const summary = sheetNamed(model, "Summary");
    // A pre-SPEC-135 dirty record: B4's annotation still in preservedXml (moved
    // with a row inserted above it, to B5); B2 never captured (a pre-SPEC-122
    // record); A7 edited since; B7 and C9 unchanged.
    const stored: Pick<Sheet, "rows" | "sourceXml"> = {
      sourceXml: summary.sourceXml,
      rows: [
        { index: 0, cells: [{ col: 0, value: "Metric" }, { col: 1, value: "Value" }] },
        { index: 1, cells: [{ col: 0, value: "Revenue" }, { col: 1, value: 1200 }] },
        { index: 4, cells: [{ col: 0, value: "Costs" }, { col: 1, value: 800, preservedXml: O.notes.B4.xml }] },
        { index: 6, cells: [{ col: 0, value: "Edited" }, { col: 1, value: "See data" }] },
        { index: 8, cells: [{ col: 2, value: "One and Two" }] },
      ],
    };
    const upgraded = upgradeOdsSheetAnnotations({ sheet: stored, prefix, dates, remapLocation: (location) => location.replace("Data!", "Facts!") });
    expect(upgraded.rows[2]).toEqual({ index: 4, cells: [{ col: 0, value: "Costs" }, { col: 1, value: 800 }] });
    // B4's element is lifted at B5 and its retained twin (at B4, now empty) not resurrected.
    expect(upgraded.notes.map((note) => [formatA1(note.ref), note.text])).toEqual([
      ["B2", O.notes.B2.text],
      ["B5", O.notes.B4.text],
    ]);
    expect(upgraded.hyperlinks.map((link) => [formatA1({ row: link.range.top, col: link.range.left }), link.target])).toEqual([
      ["B7", { kind: "place", location: "Facts!A1" }],
      ["C9", O.links.C9.target],
    ]);
    expect(upgradeOdsSheetLinks({ sheet: stored, prefix, dates })[0]!.target).toEqual(O.links.B7.target);
  });

  it("upgradeOdsSheetLinks: a linked cell moved by a structural edit keeps its link from its own preservedText, and its old address gains none", async () => {
    const model = await open();
    const prefix = model.odsContentEnvelope!.prefix;
    const summary = sheetNamed(model, "Summary");
    // A previous reader's dirty record: a row inserted above row 7 moved A7
    // and B7 (their verbatim paragraphs with them) to row 8; C9 kept its
    // place; a value equal to A7's source value was typed into the new A7
    // (no paragraphs) — it must not pick up the moved link.
    const stored: Pick<Sheet, "rows" | "sourceXml"> = {
      sourceXml: summary.sourceXml,
      rows: [
        { index: 6, cells: [{ col: 0, value: "Example site" }] },
        { index: 7, cells: [{ col: 0, value: "Example site", preservedText: O.links.A7.paragraph }, { col: 1, value: "See data", preservedText: O.links.B7.paragraph }] },
        { index: 8, cells: [{ col: 2, value: "One and Two", preservedText: O.links.C9.paragraph }] },
      ],
    };
    const links = upgradeOdsSheetLinks({ sheet: stored, prefix, dates, remapLocation: (location) => location.replace("Data!", "Facts!") });
    expect(links.map((link) => [formatA1({ row: link.range.top, col: link.range.left }), link.target, link.tooltip])).toEqual([
      ["A8", O.links.A7.target, O.links.A7.tooltip],
      ["B8", { kind: "place", location: "Facts!A1" }, undefined],
      ["C9", O.links.C9.target, undefined],
    ]);
    // The same read through the whole step; the cells keep their paragraphs,
    // which the writer re-emits verbatim while they match the model's link.
    const upgraded = upgradeOdsSheetAnnotations({ sheet: stored, prefix, dates });
    expect(upgraded.hyperlinks.map((link) => link.linkId)).toEqual(["h-A8", "h-B8", "h-C9"]);
    expect(upgraded.rows).toBe(stored.rows);
  });
});
