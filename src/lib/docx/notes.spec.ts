// SPEC-031 — footnotes and endnotes through the .docx reader and writer:
// footnotes.xml/endnotes.xml parts with w:footnoteReference/
// w:endnoteReference marks in the body, and w:footnotePr/w:endnotePr
// numbering settings on the body sectPr.
import { describe, it, expect } from "vitest";
import { parseDocx, type ParsedDocxDocument } from "./reader";
import { writeDocx } from "./writer";
import { unzip } from "../core/zip";
import {
  createDocument,
  createNoteRun,
  type DocumentModel,
  type NoteContent,
  type NoteKind,
  type Paragraph,
} from "../documentModel";
import { collectNoteRefs, noteById } from "../notes";
import { blockText } from "../textEditing";

const decoder = new TextDecoder();

const RESOLVED_DEFAULT_FORMATTING = { fontFamily: "Calibri", fontSizePt: 11 };

function paragraph(text: string): Paragraph {
  return {
    type: "paragraph",
    runs: text.length > 0 ? [{ text, formatting: { ...RESOLVED_DEFAULT_FORMATTING } }] : [],
    formatting: {},
  };
}

function note(noteId: string, kind: NoteKind, text: string): NoteContent {
  return { noteId, kind, blocks: [paragraph(text)] };
}

/** A model with one body paragraph carrying a reference mark per note. */
function modelWithNotes(notes: NoteContent[]): DocumentModel {
  const model = createDocument("spec-031-docx");
  const runs = notes.flatMap((entry, index) => [
    { text: `Body ${index} `, formatting: { ...RESOLVED_DEFAULT_FORMATTING } },
    createNoteRun(entry.noteId, entry.kind, { ...RESOLVED_DEFAULT_FORMATTING }),
  ]);
  model.sections = [{ blocks: [{ type: "paragraph", runs, formatting: {} }] }];
  model.notes = notes;
  return model;
}

async function writeAndReadBack(model: DocumentModel): Promise<{
  parsed: ParsedDocxDocument;
  parts: Map<string, Uint8Array>;
}> {
  const { bytes } = await writeDocx(model);
  const parts = await unzip(bytes);
  const parsed = await parseDocx(bytes);
  return { parsed, parts };
}

function parsedModel(parsed: ParsedDocxDocument): DocumentModel {
  const model = createDocument("spec-031-docx-reparsed");
  model.sections = parsed.sections;
  model.notes = parsed.notes;
  return model;
}

describe("footnote round trip (SPEC-031 AC-21)", () => {
  it("writes footnotes.xml with w:footnoteReference in the body and reads it back", async () => {
    const model = modelWithNotes([note("n1", "footnote", "First note"), note("n2", "footnote", "Second note")]);
    const { parsed, parts } = await writeAndReadBack(model);

    expect(parts.has("word/footnotes.xml")).toBe(true);
    const documentXml = decoder.decode(parts.get("word/document.xml")!);
    expect(documentXml).toContain('<w:footnoteReference w:id="1"/>');
    expect(documentXml).toContain('<w:footnoteReference w:id="2"/>');
    const footnotesXml = decoder.decode(parts.get("word/footnotes.xml")!);
    expect(footnotesXml).toContain('w:type="separator"');
    expect(footnotesXml).toContain('w:type="continuationSeparator"');
    expect(footnotesXml).toContain("<w:footnoteRef/>");
    const contentTypes = decoder.decode(parts.get("[Content_Types].xml")!);
    expect(contentTypes).toContain("/word/footnotes.xml");
    const rels = decoder.decode(parts.get("word/_rels/document.xml.rels")!);
    expect(rels).toContain("relationships/footnotes");

    const reparsed = parsedModel(parsed);
    const refs = collectNoteRefs(reparsed, "footnote");
    expect(refs).toHaveLength(2);
    expect(blockText(noteById(reparsed, refs[0].noteId)!.blocks[0])).toBe("First note");
    expect(blockText(noteById(reparsed, refs[1].noteId)!.blocks[0])).toBe("Second note");
    // The re-read marks stay superscript one-character atoms.
    const body = reparsed.sections[0].blocks[0] as Paragraph;
    const marks = body.runs.filter((run) => run.note);
    expect(marks).toHaveLength(2);
    expect(marks[0].formatting.vertAlign).toBe("superscript");
    expect(marks[0].text).toHaveLength(1);
  });

  it("survives a second round trip without duplicating marker runs", async () => {
    const model = modelWithNotes([note("n1", "footnote", "Stable note")]);
    const first = await writeAndReadBack(model);
    const second = await writeAndReadBack(parsedModel(first.parsed));
    const refs = collectNoteRefs(parsedModel(second.parsed), "footnote");
    expect(refs).toHaveLength(1);
    const content = second.parsed.notes.find((entry) => entry.noteId === refs[0].noteId)!;
    expect(blockText(content.blocks[0])).toBe("Stable note");
  });

  it("round-trips multi-paragraph and formatted note content", async () => {
    const content: NoteContent = {
      noteId: "n1",
      kind: "footnote",
      blocks: [
        {
          type: "paragraph",
          runs: [{ text: "Bold note", formatting: { ...RESOLVED_DEFAULT_FORMATTING, bold: true } }],
          formatting: {},
        },
        paragraph("Second paragraph"),
      ],
    };
    const model = modelWithNotes([content]);
    const { parsed } = await writeAndReadBack(model);
    const reparsed = parsedModel(parsed);
    const refs = collectNoteRefs(reparsed, "footnote");
    const readBack = noteById(reparsed, refs[0].noteId)!;
    expect(readBack.blocks).toHaveLength(2);
    const firstRun = (readBack.blocks[0] as Paragraph).runs[0];
    expect(firstRun.formatting.bold).toBe(true);
    expect(blockText(readBack.blocks[1])).toBe("Second paragraph");
  });
});

describe("endnote round trip (SPEC-031 AC-22)", () => {
  it("writes endnotes.xml with w:endnoteReference and reads it back", async () => {
    const model = modelWithNotes([note("e1", "endnote", "The endnote")]);
    const { parsed, parts } = await writeAndReadBack(model);

    expect(parts.has("word/endnotes.xml")).toBe(true);
    const documentXml = decoder.decode(parts.get("word/document.xml")!);
    expect(documentXml).toContain('<w:endnoteReference w:id="1"/>');
    const endnotesXml = decoder.decode(parts.get("word/endnotes.xml")!);
    expect(endnotesXml).toContain("<w:endnoteRef/>");

    const reparsed = parsedModel(parsed);
    const refs = collectNoteRefs(reparsed, "endnote");
    expect(refs).toHaveLength(1);
    expect(refs[0].kind).toBe("endnote");
    expect(blockText(noteById(reparsed, refs[0].noteId)!.blocks[0])).toBe("The endnote");
  });

  it("keeps both streams separate through a round trip", async () => {
    const model = modelWithNotes([note("n1", "footnote", "Foot"), note("e1", "endnote", "End")]);
    const { parsed } = await writeAndReadBack(model);
    const reparsed = parsedModel(parsed);
    expect(collectNoteRefs(reparsed, "footnote")).toHaveLength(1);
    expect(collectNoteRefs(reparsed, "endnote")).toHaveLength(1);
  });
});

describe("note numbering settings round trip (SPEC-031 AC-17..20)", () => {
  it("writes w:footnotePr on the body sectPr and reads it back", async () => {
    const model = modelWithNotes([note("n1", "footnote", "Note")]);
    model.footnoteSettings = { numberFormat: "lowerRoman", startAt: 3, restart: "page" };
    const { parsed, parts } = await writeAndReadBack(model);
    const documentXml = decoder.decode(parts.get("word/document.xml")!);
    expect(documentXml).toContain(
      '<w:footnotePr><w:numFmt w:val="lowerRoman"/><w:numStart w:val="3"/><w:numRestart w:val="eachPage"/></w:footnotePr>',
    );
    expect(parsed.footnoteSettings).toEqual({ numberFormat: "lowerRoman", startAt: 3, restart: "page" });
  });

  it("maps the symbol format to chicago and back, and section restarts to eachSect", async () => {
    const model = modelWithNotes([note("n1", "footnote", "Note"), note("e1", "endnote", "End")]);
    model.footnoteSettings = { numberFormat: "symbol", restart: "section" };
    model.endnoteSettings = { numberFormat: "decimal" };
    const { parsed, parts } = await writeAndReadBack(model);
    const documentXml = decoder.decode(parts.get("word/document.xml")!);
    expect(documentXml).toContain('<w:numFmt w:val="chicago"/>');
    expect(documentXml).toContain('<w:numRestart w:val="eachSect"/>');
    expect(documentXml).toContain('<w:endnotePr><w:numFmt w:val="decimal"/></w:endnotePr>');
    expect(parsed.footnoteSettings).toEqual({ numberFormat: "symbol", restart: "section" });
    expect(parsed.endnoteSettings).toEqual({ numberFormat: "decimal" });
  });

  it("emits no note parts or settings for a document without notes", async () => {
    const model = createDocument("spec-031-none");
    model.sections = [{ blocks: [paragraph("Plain")] }];
    const { parsed, parts } = await writeAndReadBack(model);
    expect(parts.has("word/footnotes.xml")).toBe(false);
    expect(parts.has("word/endnotes.xml")).toBe(false);
    expect(parsed.notes).toEqual([]);
    expect(parsed.footnoteSettings).toBeUndefined();
  });

  it("drops an unreferenced note on save (AC-10's file-format face)", async () => {
    const model = modelWithNotes([note("n1", "footnote", "Kept")]);
    model.notes = [...model.notes!, note("orphan", "footnote", "Gone")];
    const { parts } = await writeAndReadBack(model);
    const footnotesXml = decoder.decode(parts.get("word/footnotes.xml")!);
    expect(footnotesXml).toContain("Kept");
    expect(footnotesXml).not.toContain("Gone");
  });
});
